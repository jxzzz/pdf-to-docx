// One runnable check for the layout logic: node test/grouping.test.mjs
import { normalizeItems, groupItemsIntoLines, linesToParagraphs, lineText } from '../lib.mjs';

const it = (str, x, y, width = str.length * 5, height = 10) => ({ str, x, y, width, height });
let failed = 0;
const eq = (actual, expected, label) => {
  const a = JSON.stringify(actual), b = JSON.stringify(expected);
  if (a === b) console.log(`ok: ${label}`);
  else { failed++; console.error(`FAIL: ${label}\n  expected ${b}\n  actual   ${a}`); }
};

// Line grouping: same y clusters, output sorted top→bottom, items left→right
const lines = groupItemsIntoLines([
  it('world', 30, 50), it('Hello', 0, 50),
  it('second', 0, 35), it('para', 30, 35),
]);
eq(lines.map((l) => l.items.map((i) => i.str)), [['Hello', 'world'], ['second', 'para']], 'groups and sorts lines');

// Word spacing via x-gap between items
eq(lineText([it('Hello', 0, 0, 25), it('world', 40, 0)]), 'Hello world', 'gap becomes space');
eq(lineText([it('Hel', 0, 0, 15), it('lo', 15, 0)]), 'Hello', 'no gap, no space');

// Paragraph merge: small gap joins, big gap breaks
const mk = (y, text, height = 10) => ({
  y, height, text,
});
eq(linesToParagraphs([mk(100, 'first'), mk(90, 'second'), mk(70, 'new para')]),
  ['first second', 'new para'], 'gap threshold splits paragraphs');

// CJK join without injected space
eq(linesToParagraphs([mk(100, '第一行'), mk(90, '第二行')]), ['第一行第二行'], 'CJK lines join without space');
eq(linesToParagraphs([mk(100, 'hello'), mk(90, 'world')]), ['hello world'], 'Latin lines join with space');

// Whitespace-only items are dropped
eq(groupItemsIntoLines([it('   ', 0, 50), it('text', 0, 50)]).length, 1, 'whitespace items skipped');

// pdf.js raw shape: position lives in transform, not x/y
const raw = [
  { str: 'world', width: 25, height: 10, transform: [1, 0, 0, 1, 30, 50] },
  { str: 'Hello', width: 25, height: 10, transform: [1, 0, 0, 1, 0, 50] },
  { str: 'second line', width: 50, height: 10, transform: [1, 0, 0, 1, 0, 30] },
];
const norm = normalizeItems(raw);
const rawParas = linesToParagraphs(
  groupItemsIntoLines(norm).map((l) => ({ ...l, text: lineText(l.items) }))
);
eq(rawParas, ['Hello world', 'second line'], 'pdf.js-shaped items via normalizeItems');

// C0 control chars are stripped (they would make document.xml invalid XML)
eq(lineText([it('a\u0001b\u0000c', 0, 0)]), 'abc', 'C0 control chars stripped');

if (failed) { console.error(`\n${failed} test(s) failed`); process.exit(1); }
console.log('\nall passed');
