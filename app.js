import { normalizeItems, groupItemsIntoLines, linesToParagraphs, lineText } from './lib.mjs';

const PDFJS_VER = '4.10.38';
let pdfjs = null;
async function loadPdfjs() {
  if (!pdfjs) {
    pdfjs = await import(
      `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFJS_VER}/pdf.min.mjs`
    );
    pdfjs.GlobalWorkerOptions.workerSrc =
      `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFJS_VER}/pdf.worker.min.mjs`;
  }
  return pdfjs;
}

const $ = (s) => document.querySelector(s);
const drop = $('#drop');
const input = $('#file');
const statusEl = $('#status');
const errorEl = $('#error');

drop.addEventListener('click', () => input.click());
drop.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    input.click();
  }
});
drop.addEventListener('dragover', (e) => {
  e.preventDefault();
  drop.classList.add('drag');
});
drop.addEventListener('dragleave', () => drop.classList.remove('drag'));
drop.addEventListener('drop', (e) => {
  e.preventDefault();
  drop.classList.remove('drag');
  const f = e.dataTransfer.files[0];
  if (f) convert(f);
});
input.addEventListener('change', () => {
  if (input.files[0]) convert(input.files[0]);
  input.value = '';
});

// A stray drop outside the dropzone would otherwise navigate the browser
// to the dropped file, killing the app.
for (const type of ['dragover', 'drop']) {
  window.addEventListener(type, (e) => e.preventDefault());
}

function setStatus(msg) {
  errorEl.hidden = true;
  statusEl.textContent = msg;
  statusEl.hidden = !msg;
}

function showError(msg) {
  statusEl.hidden = true;
  errorEl.textContent = msg;
  errorEl.hidden = false;
}

let busy = false;

async function convert(file) {
  if (busy) return;
  if (file.type !== 'application/pdf' && !/\.pdf$/i.test(file.name)) {
    showError('Please choose a .pdf file / 请选择 PDF 文件');
    return;
  }
  if (typeof docx === 'undefined') {
    showError('The docx library failed to load — check your network and reload / 转换组件加载失败，请检查网络后刷新');
    return;
  }
  busy = true;
  setStatus(`Reading ${file.name}…`);
  try {
    const blob = await pdfFileToDocx(file);
    download(blob, file.name.replace(/\.pdf$/i, '') + '.docx');
    setStatus(`✓ Converted ${file.name} → .docx`);
  } catch (e) {
    if (e?.name === 'PasswordException') {
      showError('This PDF is password-protected and cannot be converted / 该 PDF 已加密，无法转换');
    } else {
      showError('Conversion failed: ' + (e?.message || e));
    }
  } finally {
    busy = false;
  }
}

async function pdfFileToDocx(file) {
  const { getDocument } = await loadPdfjs();
  // cMapUrl is required for CJK PDFs whose fonts rely on predefined CMaps.
  const task = getDocument({
    data: await file.arrayBuffer(),
    cMapUrl: `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VER}/cmaps/`,
    cMapPacked: true,
  });
  try {
    const pdf = await task.promise;

    const children = [];
    let foundText = false;
    for (let p = 1; p <= pdf.numPages; p++) {
      setStatus(`Extracting text — page ${p} of ${pdf.numPages}…`);
      const page = await pdf.getPage(p);
      const { items } = await page.getTextContent();
      const paras = linesToParagraphs(
        groupItemsIntoLines(normalizeItems(items)).map((l) => ({
          ...l,
          text: lineText(l.items),
        }))
      );
      if (paras.length) foundText = true;
      for (const text of paras) children.push(new docx.Paragraph(text));
      if (p < pdf.numPages) {
        children.push(new docx.Paragraph({ children: [new docx.PageBreak()] }));
      }
    }

    if (!foundText) {
      throw new Error(
        'No extractable text (scanned PDF?). This tool needs a text-based PDF / ' +
          '没有可提取的文字（可能是扫描件），本工具只支持文本型 PDF'
      );
    }

    const doc = new docx.Document({ sections: [{ children }] });
    return docx.Packer.toBlob(doc);
  } finally {
    // Tears down the per-conversion pdf.js Web Worker; without this every
    // conversion leaks a worker holding the whole PDF.
    await task.destroy();
  }
}

function download(blob, name) {
  const a = document.createElement('a');
  const url = URL.createObjectURL(blob);
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  // Deferred: an immediate revoke races the blob fetch in WebKit.
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}
