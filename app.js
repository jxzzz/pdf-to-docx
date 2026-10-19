import { normalizeItems, groupItemsIntoLines, linesToParagraphs, lineText, getTemplate, pageGeometry } from './lib.mjs';

const PDFJS_VER = '4.10.38';
let pdfjs = null;
async function loadPdfjs() {
  if (!pdfjs) {
    pdfjs = await import(`https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFJS_VER}/pdf.min.mjs`);
    pdfjs.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFJS_VER}/pdf.worker.min.mjs`;
  }
  return pdfjs;
}

const $ = (s) => document.querySelector(s);
const drop = $('#drop');
const input = $('#file');
const statusEl = $('#status');
const errorEl = $('#error');
const preview = $('#preview');
const templateSelect = $('#template');
const fidelity = $('#fidelity');
let mode = 'pdf-to-word';
let busy = false;

for (const button of document.querySelectorAll('[data-mode]')) {
  button.addEventListener('click', () => {
    mode = button.dataset.mode;
    document.querySelectorAll('[data-mode]').forEach((b) => b.classList.toggle('active', b === button));
    input.accept = mode === 'pdf-to-word' ? 'application/pdf,.pdf' : '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    drop.querySelector('strong').textContent = mode === 'pdf-to-word' ? 'Drop a PDF here' : 'Drop a Word file here';
    drop.querySelector('.hint').textContent = mode === 'pdf-to-word' ? '拖入 PDF 文件，或点击选择' : '拖入 Word 文件，或点击选择';
    preview.hidden = true;
    templateSelect.closest('.template-picker').hidden = mode === 'pdf-to-word';
    fidelity.closest('.fidelity-picker').hidden = mode !== 'pdf-to-word';
    setStatus('');
  });
}

drop.addEventListener('click', () => input.click());
drop.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
});
drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('drag'); });
drop.addEventListener('dragleave', () => drop.classList.remove('drag'));
drop.addEventListener('drop', (e) => {
  e.preventDefault();
  drop.classList.remove('drag');
  if (e.dataTransfer.files[0]) convert(e.dataTransfer.files[0]);
});
input.addEventListener('change', () => {
  if (input.files[0]) convert(input.files[0]);
  input.value = '';
});
for (const type of ['dragover', 'drop']) window.addEventListener(type, (e) => e.preventDefault());
templateSelect.addEventListener('change', () => preview.className = getTemplate(templateSelect.value).className);
$('#print').addEventListener('click', () => window.print());

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

async function convert(file) {
  if (busy) return;
  const isPdf = file.type === 'application/pdf' || /\.pdf$/i.test(file.name);
  const isDocx = file.type.includes('wordprocessingml') || /\.docx$/i.test(file.name);
  if ((mode === 'pdf-to-word' && !isPdf) || (mode === 'word-to-pdf' && !isDocx)) {
    showError(mode === 'pdf-to-word' ? 'Please choose a .pdf file / 请选择 PDF 文件' : 'Please choose a .docx file / 请选择 Word 文件');
    return;
  }
  if ((mode === 'pdf-to-word' && typeof docx === 'undefined') ||
      (mode === 'word-to-pdf' && typeof mammoth === 'undefined')) {
    showError('Conversion components failed to load — check your network and reload / 转换组件加载失败，请检查网络后刷新');
    return;
  }
  busy = true;
  try {
    if (mode === 'pdf-to-word') {
      setStatus(`Reading ${file.name}…`);
      const blob = fidelity.checked ? await pdfFileToDocxFidelity(file) : await pdfFileToDocx(file);
      download(blob, file.name.replace(/\.pdf$/i, '') + '.docx');
      setStatus(`✓ Converted ${file.name} → .docx`);
    } else {
      setStatus(`Preparing preview for ${file.name}…`);
      const html = await wordFileToHtml(file);
      preview.innerHTML = html.value || '<p>No readable text found.</p>';
      preview.className = getTemplate(templateSelect.value).className;
      preview.hidden = false;
      setStatus('Preview ready. Choose a style, then print / 预览已生成，可选择样式后打印');
    }
  } catch (e) {
    if (e?.name === 'PasswordException') {
      showError('This PDF is password-protected and cannot be converted / 该 PDF 已加密，无法转换');
    } else {
      showError((mode === 'pdf-to-word' ? 'Conversion failed: ' : 'Preview failed: ') + (e?.message || e));
    }
  } finally { busy = false; }
}

async function wordFileToHtml(file) {
  return mammoth.convertToHtml({ arrayBuffer: await file.arrayBuffer() });
}

async function pdfFileToDocxFidelity(file) {
  const { getDocument } = await loadPdfjs();
  const task = getDocument({
    data: await file.arrayBuffer(),
    cMapUrl: `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VER}/cmaps/`,
    cMapPacked: true,
  });
  try {
    const pdf = await task.promise;
    // One section per page so each page keeps its own size (mixed A4/Letter).
    const sections = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      setStatus(`Rendering page ${p} of ${pdf.numPages}…`);
      const page = await pdf.getPage(p);
      // ponytail: fixed 2× scale, JPEG q0.92 — bump to 3×/PNG if fine print blurs.
      const viewport = page.getViewport({ scale: 2 });
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
      const blob = await new Promise((r) => canvas.toBlob(r, 'image/jpeg', 0.92));
      const g = pageGeometry(viewport);
      sections.push({
        properties: { page: { size: { width: g.twipW, height: g.twipH }, margin: { top: 0, right: 0, bottom: 0, left: 0 } } },
        children: [new docx.Paragraph({
          children: [new docx.ImageRun({ data: await blob.arrayBuffer(), transformation: { width: g.pxW, height: g.pxH } })],
        })],
      });
      page.cleanup();
    }
    return docx.Packer.toBlob(new docx.Document({ sections }));
  } finally { await task.destroy(); }
}

async function pdfFileToDocx(file) {
  const { getDocument } = await loadPdfjs();
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
      const paras = linesToParagraphs(groupItemsIntoLines(normalizeItems(items)).map((l) => ({ ...l, text: lineText(l.items) })));
      if (paras.length) foundText = true;
      for (const text of paras) children.push(new docx.Paragraph(text));
      if (p < pdf.numPages) children.push(new docx.Paragraph({ children: [new docx.PageBreak()] }));
    }
    if (!foundText) throw new Error('No extractable text (scanned PDF?). This tool needs a text-based PDF / 没有可提取的文字');
    return docx.Packer.toBlob(new docx.Document({ sections: [{ children }] }));
  } finally { await task.destroy(); }
}

function download(blob, name) {
  const a = document.createElement('a');
  const url = URL.createObjectURL(blob);
  a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}
