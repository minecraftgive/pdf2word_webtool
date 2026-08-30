# PDF → Word (Tauri + Rust)

A **small, self-contained desktop app** that converts PDF to editable Word
(`.docx`), with the PDF conversion logic written in pure Rust. This is a
companion to the Bun/TypeScript version in `../pdf2word` — same UI, ~5× smaller.

## Bundle size

| Artifact | Size |
| --- | --- |
| `.app` (macOS) | ~14 MB |
| `.dmg` (installer) | ~4.2 MB |
| release binary | ~13.5 MB |

(Compared to ~65–89 MB for the Bun single-binary build, and ~150 MB+ for Electron.)

## How it works

- **PDF parsing** — `lopdf` walks each page's content stream, tracking text
  position/font size via `Tm`/`Tf` and images via `Do` + the CTM from `cm`.
  Text is decoded through the font's `ToUnicode`/encoding (via `lopdf`'s
  `Encoding`), which handles all font encodings.
- **Reconstruction** — text runs are grouped into lines, lines into paragraphs,
  headings detected by font-size ratio, images re-encoded to PNG (`image` crate)
  and interleaved by Y position.
- **Generation** — `docx-rs` emits the `.docx`.
- **Shell** — `tauri` 2 hosts the existing HTML/CSS/JS frontend in the native
  webview; conversion runs as a Tauri command with live progress events.

## Build

Requirements: Rust (stable) + the platform's Tauri prerequisites
(Xcode CLT on macOS, WebView2 + MSVC on Windows, webkit2gtk on Linux).

```bash
cd src-tauri
cargo run                 # debug
cargo tauri build         # release bundle (.app / .dmg on macOS)
```

From the project root, `bunx @tauri-apps/cli build` works too (with `cargo` on
`PATH`).

## Layout

```
src/            frontend (index.html, style.css, app.js) — Tauri IPC
src-tauri/
  src/convert.rs   PDF → docx conversion (pure Rust)
  src/lib.rs       convert_pdf command + progress events
  src/main.rs      entry point
  tauri.conf.json  window + bundle config
  icons/           generated icon set
testdata/        sample PDFs for tests (cargo test)
```

## Notes & limitations

- Scanned (image-only) PDFs return a clear "OCR not supported" error.
- Text is reconstructed as flowing paragraphs + headings + images, not a
  pixel-perfect reproduction of columns/tables.
- `cargo test` runs the conversion against `testdata/` sample PDFs.
