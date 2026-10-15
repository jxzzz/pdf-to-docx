# PDF → Word

Convert PDF files to Word (.docx) **entirely in your browser** — no upload, no server, no tracking.

在线将 PDF 转换为 Word（.docx），全部在浏览器本地完成，文件不会上传。

**Try it:** https://jxzzz.github.io/pdf-to-docx/

## How it works

1. [pdf.js](https://mozilla.github.io/pdf.js/) extracts the text of every page in your browser.
2. Text items are grouped into lines and paragraphs by their position on the page.
3. The [docx](https://docx.js.org/) library builds a `.docx` file and your browser downloads it.

## Limitations

Text content only (paragraphs + page breaks). Images, tables, columns and exact
formatting are not preserved — for those, use a server-based converter.

## Run locally

```bash
python3 -m http.server  # then open http://localhost:8000
```

## Test

```bash
node test/grouping.test.mjs
```

## License

MIT
