//! PDF → Word (.docx) conversion in pure Rust.
//!
//! Uses `lopdf` to walk each page's content stream (tracking text position and
//! font size via `Tm`/`Tf`, and images via `Do` + the CTM from `cm`), and
//! `docx-rs` to emit the Word document. Headings are detected by font-size
//! ratio; embedded images are re-encoded to PNG and interleaved by Y position.

use std::collections::BTreeMap;
use std::io::Cursor;

use docx_rs::{BreakType, Docx, Paragraph, Pic, Run};
use lopdf::content::Content;
use lopdf::{Document, Encoding, Object, ObjectId};

/// One reconstructed line of text.
#[derive(Debug, Clone)]
struct Line {
    text: String,
    /// Baseline Y (PDF coordinates, origin bottom-left).
    y: f32,
    x: f32,
    font_size: f32,
}

/// One embedded image, decoded to PNG bytes.
struct ExtractedImage {
    x: f32,
    y: f32,
    width: u32,
    height: u32,
    png: Vec<u8>,
}

struct TextRun {
    text: String,
    x: f32,
    y: f32,
    font_size: f32,
}

/// A semantic block in the output document.
enum Block {
    Title(String),
    Heading(String, &'static str),
    Paragraph(Vec<String>),
    Image(ExtractedImage),
    PageBreak,
}

/// Convert a PDF (bytes) into a .docx (bytes), emitting progress updates.
pub fn pdf_to_docx<F>(data: &[u8], mut on_progress: F) -> Result<Vec<u8>, String>
where
    F: FnMut(&str, &str),
{
    let doc = Document::load_mem(data).map_err(|e| format!("无法解析 PDF: {e}"))?;
    let pages = doc.get_pages();
    let total = pages.len();

    on_progress("加载 PDF 文档", "");

    let mut page_lines: Vec<Vec<Line>> = Vec::new();
    let mut page_images: Vec<Vec<ExtractedImage>> = Vec::new();

    for (num, page_id) in pages {
        on_progress("解析页面", &format!("第 {num} / {total} 页"));
        let (lines, images) = parse_page(&doc, page_id);
        page_lines.push(lines);
        page_images.push(images);
    }

    let total_lines: usize = page_lines.iter().map(|p| p.len()).sum();
    let image_count: usize = page_images.iter().map(|p| p.len()).sum();
    if total_lines == 0 && image_count == 0 {
        return Err("此 PDF 不含可提取的文本或图片（可能是扫描件，暂不支持 OCR）".into());
    }
    if image_count > 0 {
        on_progress("提取图片", &format!("{image_count} 张"));
    }

    on_progress("生成 Word 文档", "");
    let blocks = classify(&page_lines, &page_images);
    generate(blocks)
}

// --- Content-stream parsing -------------------------------------------------

fn parse_page(doc: &Document, page_id: ObjectId) -> (Vec<Line>, Vec<ExtractedImage>) {
    let mut runs: Vec<TextRun> = Vec::new();
    let mut images: Vec<ExtractedImage> = Vec::new();

    let fonts = doc.get_page_fonts(page_id).unwrap_or_default();
    let mut encodings: BTreeMap<Vec<u8>, Encoding> = BTreeMap::new();
    for (name, font_dict) in &fonts {
        if let Ok(enc) = font_dict.get_font_encoding(doc) {
            encodings.insert(name.clone(), enc);
        }
    }

    let content_data = doc.get_page_content(page_id);
    let content = match Content::decode(&content_data) {
        Ok(c) => c,
        Err(_) => return (Vec::new(), Vec::new()),
    };

    let mut font_size = 0f32;
    let mut current_encoding: Option<&Encoding> = None;
    let mut tm = [1f32, 0.0, 0.0, 1.0, 0.0, 0.0]; // text matrix
    let mut tlm = tm; // text line matrix
    let mut leading = 0f32;
    let mut ctm = [1f32, 0.0, 0.0, 1.0, 0.0, 0.0]; // current transformation matrix
    let mut ctm_stack: Vec<[f32; 6]> = Vec::new();

    for op in &content.operations {
        match op.operator.as_str() {
            "BT" => {
                tm = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];
                tlm = tm;
            }
            "Tf" => {
                if let Some(name) = op.operands.first().and_then(obj_name) {
                    font_size = op.operands.get(1).map(obj_float).unwrap_or(0.0);
                    current_encoding = encodings.get(name);
                }
            }
            "Tm" => {
                if let Some(m) = parse_matrix(&op.operands) {
                    tm = m;
                    tlm = m;
                }
            }
            "Td" | "TD" => {
                let tx = op.operands.first().map(obj_float).unwrap_or(0.0);
                let ty = op.operands.get(1).map(obj_float).unwrap_or(0.0);
                tlm[4] += tx;
                tlm[5] += ty;
                tm = tlm;
                if op.operator == "TD" {
                    leading = -ty;
                }
            }
            "T*" => {
                tlm[5] -= leading;
                tm = tlm;
            }
            "TL" => leading = op.operands.first().map(obj_float).unwrap_or(0.0),
            "Tj" | "TJ" => {
                if let Some(enc) = current_encoding {
                    if let Some(text) = decode_text(enc, &op.operands) {
                        let (x, y) = apply_ctm(&ctm, tm[4], tm[5]);
                        runs.push(TextRun { text, x, y, font_size });
                    }
                }
            }
            "cm" => {
                if let Some(m) = parse_matrix(&op.operands) {
                    ctm = concat(&ctm, &m);
                }
            }
            "q" => ctm_stack.push(ctm),
            "Q" => ctm = ctm_stack.pop().unwrap_or([1.0, 0.0, 0.0, 1.0, 0.0, 0.0]),
            "Do" => {
                if let Some(name) = op.operands.first().and_then(obj_name) {
                    if let Some(img) = extract_image(doc, page_id, name, &ctm) {
                        images.push(img);
                    }
                }
            }
            _ => {}
        }
    }

    (group_lines(runs), images)
}

fn group_lines(runs: Vec<TextRun>) -> Vec<Line> {
    if runs.is_empty() {
        return Vec::new();
    }
    let mut sorted = runs;
    sorted.sort_by(|a, b| b.y.partial_cmp(&a.y).unwrap_or(std::cmp::Ordering::Equal));

    let mut lines: Vec<Line> = Vec::new();
    let mut current: Vec<TextRun> = Vec::new();
    let mut current_y = sorted[0].y;

    for run in sorted {
        let tol = (run.font_size * 0.4).max(2.0);
        if (run.y - current_y).abs() > tol {
            push_line(&mut lines, &mut current, current_y);
            current_y = run.y;
        }
        current.push(run);
    }
    push_line(&mut lines, &mut current, current_y);
    lines
}

fn push_line(lines: &mut Vec<Line>, current: &mut Vec<TextRun>, y: f32) {
    if current.is_empty() {
        return;
    }
    current.sort_by(|a, b| a.x.partial_cmp(&b.x).unwrap_or(std::cmp::Ordering::Equal));
    let text = current
        .iter()
        .map(|r| r.text.as_str())
        .collect::<Vec<_>>()
        .join(" ")
        .split_whitespace()
        .collect::<Vec<_>>()
        .join(" ");
    if text.is_empty() {
        current.clear();
        return;
    }
    let sizes: Vec<f32> = current.iter().map(|r| r.font_size).collect();
    let font_size = median(&sizes);
    let x = current.iter().map(|r| r.x).fold(f32::INFINITY, f32::min);
    lines.push(Line { text, y, x, font_size });
    current.clear();
}

fn decode_text(enc: &Encoding, operands: &[Object]) -> Option<String> {
    let mut out = String::new();
    let mut found = false;
    for operand in operands {
        match operand {
            Object::String(bytes, _) => {
                if enc.write_to_string(bytes, &mut out).is_ok() {
                    found = true;
                }
            }
            Object::Array(arr) => {
                for item in arr {
                    match item {
                        Object::String(bytes, _) => {
                            if enc.write_to_string(bytes, &mut out).is_ok() {
                                found = true;
                            }
                        }
                        Object::Integer(i) if *i < -100 => out.push(' '),
                        _ => {}
                    }
                }
                found = true;
            }
            _ => {}
        }
    }
    if !found {
        return None;
    }
    let trimmed = out.trim().to_string();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed)
    }
}

// --- Image extraction -------------------------------------------------------

fn extract_image(doc: &Document, page_id: ObjectId, name: &[u8], ctm: &[f32; 6]) -> Option<ExtractedImage> {
    let (res_dict, _ids) = doc.get_page_resources(page_id).ok()?;
    let xobject = res_dict?.get(b"XObject").ok()?.as_dict().ok()?;
    let xobj_id = xobject.get(name).ok()?.as_reference().ok()?;
    let stream = doc.get_object(xobj_id).ok()?.as_stream().ok()?;

    let dict = &stream.dict;
    let width = dict.get(b"Width").ok().and_then(obj_int)? as u32;
    let height = dict.get(b"Height").ok().and_then(obj_int)? as u32;
    let bpc = dict.get(b"BitsPerComponent").ok().and_then(obj_int).unwrap_or(8) as u32;
    let colorspace = dict.get(b"ColorSpace").ok().and_then(obj_name).map(|s| s.to_vec());

    let filter_is_jpeg = match dict.get(b"Filter").ok() {
        Some(Object::Name(n)) => n == b"DCTDecode" || n == b"DCT",
        Some(Object::Array(arr)) => arr.iter().any(|o| matches!(o, Object::Name(n) if n == b"DCTDecode")),
        _ => false,
    };

    let png = if filter_is_jpeg {
        image_to_png(&stream.content)?
    } else {
        let raw = stream.decompressed_content().ok()?;
        raw_to_png(&raw, width, height, bpc, colorspace.as_deref())?
    };

    Some(ExtractedImage { x: ctm[4], y: ctm[5], width, height, png })
}

fn image_to_png(bytes: &[u8]) -> Option<Vec<u8>> {
    let img = image::load_from_memory(bytes).ok()?;
    let mut out = Cursor::new(Vec::new());
    img.write_to(&mut out, image::ImageFormat::Png).ok()?;
    Some(out.into_inner())
}

fn raw_to_png(raw: &[u8], w: u32, h: u32, bpc: u32, colorspace: Option<&[u8]>) -> Option<Vec<u8>> {
    if bpc != 8 {
        return None;
    }
    let mut out = Cursor::new(Vec::new());
    let ok = match colorspace {
        Some(cs) if cs == b"DeviceGray" || cs == b"G" => {
            if raw.len() < (w * h) as usize {
                return None;
            }
            image::GrayImage::from_raw(w, h, raw[..(w * h) as usize].to_vec())?
                .write_to(&mut out, image::ImageFormat::Png)
                .is_ok()
        }
        _ => {
            if raw.len() < (w * h * 3) as usize {
                return None;
            }
            image::RgbImage::from_raw(w, h, raw[..(w * h * 3) as usize].to_vec())?
                .write_to(&mut out, image::ImageFormat::Png)
                .is_ok()
        }
    };
    if ok {
        Some(out.into_inner())
    } else {
        None
    }
}

// --- Classification ---------------------------------------------------------

fn classify(page_lines: &[Vec<Line>], page_images: &[Vec<ExtractedImage>]) -> Vec<Block> {
    let all_sizes: Vec<f32> = page_lines
        .iter()
        .flatten()
        .map(|l| l.font_size)
        .filter(|s| *s > 0.0)
        .collect();
    let body_size = mode(&all_sizes).unwrap_or(12.0);

    let mut blocks: Vec<Block> = Vec::new();
    let mut title_consumed = false;
    let title_line = page_lines
        .first()
        .and_then(|p| p.iter().max_by(|a, b| a.font_size.partial_cmp(&b.font_size).unwrap_or(std::cmp::Ordering::Equal)));

    for (pi, lines) in page_lines.iter().enumerate() {
        // Merge lines and images into a single ordered list (top → bottom).
        enum Item<'a> {
            Line(&'a Line),
            Image(&'a ExtractedImage),
        }
        impl Item<'_> {
            fn top_y(&self) -> f32 {
                match self {
                    Item::Line(l) => l.y + l.font_size,
                    Item::Image(i) => i.y + i.height as f32,
                }
            }
        }
        let mut items: Vec<Item> = lines.iter().map(Item::Line).collect();
        if let Some(imgs) = page_images.get(pi) {
            items.extend(imgs.iter().map(Item::Image));
        }
        items.sort_by(|a, b| b.top_y().partial_cmp(&a.top_y()).unwrap_or(std::cmp::Ordering::Equal));

        let mut paragraph: Vec<String> = Vec::new();
        let flush_para = |blocks: &mut Vec<Block>, paragraph: &mut Vec<String>| {
            if !paragraph.is_empty() {
                blocks.push(Block::Paragraph(std::mem::take(paragraph)));
            }
        };

        for item in items {
            match item {
                Item::Image(img) => {
                    flush_para(&mut blocks, &mut paragraph);
                    blocks.push(Block::Image(ExtractedImage {
                        x: img.x,
                        y: img.y,
                        width: img.width,
                        height: img.height,
                        png: img.png.clone(),
                    }));
                }
                Item::Line(line) => {
                    let ratio = if body_size > 0.0 { line.font_size / body_size } else { 1.0 };
                    let is_title = !title_consumed
                        && title_line.map_or(false, |t| std::ptr::eq(t, line) && ratio >= 1.2);
                    if is_title {
                        flush_para(&mut blocks, &mut paragraph);
                        blocks.push(Block::Title(line.text.clone()));
                        title_consumed = true;
                    } else if ratio >= 1.08 {
                        flush_para(&mut blocks, &mut paragraph);
                        let style = if ratio >= 1.3 {
                            "Heading1"
                        } else if ratio >= 1.15 {
                            "Heading2"
                        } else {
                            "Heading3"
                        };
                        blocks.push(Block::Heading(line.text.clone(), style));
                    } else {
                        paragraph.push(line.text.clone());
                    }
                }
            }
        }
        flush_para(&mut blocks, &mut paragraph);
        if pi + 1 < page_lines.len() {
            blocks.push(Block::PageBreak);
        }
    }
    blocks
}

// --- DOCX generation --------------------------------------------------------

fn generate(blocks: Vec<Block>) -> Result<Vec<u8>, String> {
    let mut docx = Docx::new();
    for block in blocks {
        match block {
            Block::Title(text) => {
                docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_text(text)).style("Title"));
            }
            Block::Heading(text, style) => {
                docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_text(text)).style(style));
            }
            Block::Paragraph(lines) => {
                docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_text(lines.join(" "))));
            }
            Block::Image(img) => {
                let pic = Pic::new_with_dimensions(img.png, img.width, img.height);
                docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_image(pic)));
            }
            Block::PageBreak => {
                docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_break(BreakType::Page)));
            }
        }
    }

    let xml = docx.build();
    let mut cursor = Cursor::new(Vec::new());
    xml.pack(&mut cursor).map_err(|e| format!("生成 Word 文档失败: {e}"))?;
    Ok(cursor.into_inner())
}

// --- Small helpers ----------------------------------------------------------

fn obj_float(o: &Object) -> f32 {
    match o {
        Object::Real(r) => *r,
        Object::Integer(i) => *i as f32,
        _ => 0.0,
    }
}

fn obj_int(o: &Object) -> Option<i64> {
    match o {
        Object::Integer(i) => Some(*i),
        _ => None,
    }
}

fn obj_name(o: &Object) -> Option<&[u8]> {
    match o {
        Object::Name(n) => Some(n.as_slice()),
        _ => None,
    }
}

fn parse_matrix(operands: &[Object]) -> Option<[f32; 6]> {
    if operands.len() < 6 {
        return None;
    }
    Some([
        obj_float(&operands[0]),
        obj_float(&operands[1]),
        obj_float(&operands[2]),
        obj_float(&operands[3]),
        obj_float(&operands[4]),
        obj_float(&operands[5]),
    ])
}

/// Concatenate a `cm` matrix onto the CTM (pre-multiply: new = m × ctm).
fn concat(ctm: &[f32; 6], m: &[f32; 6]) -> [f32; 6] {
    let [a, b, c, d, e, f] = *m;
    let [a2, b2, c2, d2, e2, f2] = *ctm;
    [
        a * a2 + b * c2,
        a * b2 + b * d2,
        c * a2 + d * c2,
        c * b2 + d * d2,
        e * a2 + f * c2 + e2,
        e * b2 + f * d2 + f2,
    ]
}

fn apply_ctm(ctm: &[f32; 6], x: f32, y: f32) -> (f32, f32) {
    (ctm[0] * x + ctm[2] * y + ctm[4], ctm[1] * x + ctm[3] * y + ctm[5])
}

fn median(vals: &[f32]) -> f32 {
    if vals.is_empty() {
        return 0.0;
    }
    let mut sorted: Vec<f32> = vals.to_vec();
    sorted.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let mid = sorted.len() / 2;
    if sorted.len() % 2 == 0 {
        (sorted[mid - 1] + sorted[mid]) / 2.0
    } else {
        sorted[mid]
    }
}

fn mode(vals: &[f32]) -> Option<f32> {
    let mut freq: BTreeMap<i64, (f32, usize)> = BTreeMap::new();
    for &v in vals {
        let key = (v * 2.0).round() as i64;
        let entry = freq.entry(key).or_insert((v, 0));
        entry.1 += 1;
    }
    let max_count = freq.values().map(|(_, c)| *c).max()?;
    // Among the most frequent sizes, prefer the smallest — body text is
    // typically the smallest and most common size.
    freq.into_values()
        .filter(|(_, c)| *c == max_count)
        .map(|(v, _)| v)
        .min_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn converts_text_pdf() {
        let path = concat!(env!("CARGO_MANIFEST_DIR"), "/../testdata/sample.pdf");
        let data = std::fs::read(path).unwrap();
        let docx = pdf_to_docx(&data, |_, _| {}).expect("conversion should succeed");
        assert!(docx.starts_with(b"PK"), "docx should be a zip");
    }

    #[test]
    fn converts_image_pdf() {
        let path = concat!(env!("CARGO_MANIFEST_DIR"), "/../testdata/img-sample.pdf");
        let data = std::fs::read(path).unwrap();
        let docx = pdf_to_docx(&data, |_, _| {}).expect("conversion should succeed");
        assert!(docx.starts_with(b"PK"), "docx should be a zip");
    }
}
