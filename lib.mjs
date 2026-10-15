// Pure text-layout logic, shared between the app and tests. No DOM here.

// Map pdf.js TextItems ({str, transform, width, height}) to the plain
// {str, x, y, width, height} shape the grouping functions expect.
// pdf.js keeps the position in transform[4]/transform[5]; there is no x/y.
export function normalizeItems(items) {
  return items
    .filter((i) => typeof i.str === 'string' && i.transform)
    .map((i) => ({ ...i, x: i.transform[4], y: i.transform[5] }));
}

// Group text items into visual lines by y-coordinate (y grows upward).
export function groupItemsIntoLines(items) {
  const lines = [];
  for (const it of items) {
    if (!it.str || !it.str.trim()) continue;
    const tol = Math.max(2, (it.height || 10) * 0.5);
    const line = lines.find((l) => Math.abs(l.y - it.y) <= tol);
    if (line) {
      line.items.push(it);
      line.height = Math.max(line.height, it.height || 10);
    } else {
      lines.push({ y: it.y, height: it.height || 10, items: [it] });
    }
  }
  lines.sort((a, b) => b.y - a.y); // top of page first
  for (const l of lines) l.items.sort((a, b) => a.x - b.x);
  return lines;
}

// Characters that are illegal in XML 1.0 — pdf.js passes C0 controls through
// for CID fonts without a ToUnicode map, and docx writes them raw, which
// makes Word refuse to open the file.
const BAD_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g;

// Join one line's items into text, inserting a space when the horizontal
// gap between items is wide enough to be a word break.
export function lineText(items) {
  let out = '';
  let penX = null;
  for (const it of items) {
    const gap = penX === null ? Infinity : it.x - penX;
    if (out && gap > it.height * 0.15 && !/\s$/.test(out) && !/^\s/.test(it.str)) out += ' ';
    out += it.str;
    penX = it.x + (it.width || 0);
  }
  return out.replace(BAD_XML, '').replace(/\s+/g, ' ').trim();
}

const CJK = /[　-鿿豈-﫿＀-￯]/;

function joinLines(a, b) {
  // CJK text has no spaces between words — don't inject one at the join.
  if (CJK.test(a.slice(-1)) || CJK.test(b[0])) return a + b;
  return a + ' ' + b;
}

// Merge lines (top→bottom, each {y, height, text}) into paragraphs:
// ponytail: gap-based merge only — columns/tables/indent-based paragraph
// detection would need real layout analysis; server-side conversion is the
// upgrade path.
export function linesToParagraphs(lines) {
  const paras = [];
  let prev = null;
  for (const line of lines) {
    const gap = prev ? prev.y - line.y : Infinity;
    if (prev && gap <= Math.max(prev.height, line.height) * 1.6) {
      paras[paras.length - 1] = joinLines(paras[paras.length - 1], line.text);
    } else {
      paras.push(line.text);
    }
    prev = line;
  }
  return paras;
}
