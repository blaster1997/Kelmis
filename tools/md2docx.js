// Markdown -> Word-Bericht im klassischen Berichtsstil (docx-js).
// Unterstützt: # / ## / ### Überschriften (automatisch nummeriert; "{-}" am Ende = ohne Nummer),
// Absätze mit **fett** / *kursiv* / [Link](URL), Listen (-, 1.), GFM-Tabellen,
// "Tab: Titel" (Tabellenbeschriftung vor einer Tabelle), ![Titel](pfad){breite_cm} (Abbildung),
// "Quelle: ..." (Quellenzeile), :::box Titel ... ::: (gerahmter Hinweis), > Hinweis,
// :::toc, :::abbildungsverzeichnis, :::tabellenverzeichnis, ---pagebreak---.
// Aufruf: node md2docx.js input.md output.docx "Kopfzeile" cover.json [seitenkarte.json]
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun, ExternalHyperlink,
  Header, Footer, AlignmentType, LevelFormat, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageNumber, PageBreak, TabStopType, TableLayoutType, HeadingLevel, LineRuleType,
} = require("docx");

const [, , inFile, outFile, headerTitle = "", coverFile, pageMapFile] = process.argv;
const md = fs.readFileSync(inFile, "utf8").replace(/<!--[\s\S]*?-->/g, "");
const baseDir = path.dirname(path.resolve(inFile));
const cover = coverFile ? JSON.parse(fs.readFileSync(coverFile, "utf8")) : null;
let firstChapter = false;
const pageMap = pageMapFile && fs.existsSync(pageMapFile) ? JSON.parse(fs.readFileSync(pageMapFile, "utf8")) : {};

// ---------- Gestaltung ----------
const FONT = "Arial";
const C = { text: "000000", grey: "595959", line: "808080", head: "D9D9D9", box: "F2F2F2", link: "0563C1" };
const SZ = { body: 20, table: 18, small: 16, h1: 28, h2: 24, h3: 21 };
const PAGE_W = 11906, PAGE_H = 16838, ML = 1418, MR = 1134; // A4, links 2,5 cm, rechts 2,0 cm
const CW = PAGE_W - ML - MR;
const LINE = 288; // Zeilenabstand 1,2

// ---------- Inline ----------
function runs(text, base = {}) {
  const linkRe = /\[([^\]]+)\]\((https?:[^)]+)\)/g;
  if (/\[([^\]]+)\]\((https?:[^)]+)\)/.test(text)) {
    const parts = []; let last = 0, m;
    while ((m = linkRe.exec(text)) !== null) {
      if (m.index > last) parts.push(...runs(text.slice(last, m.index), base));
      parts.push(new ExternalHyperlink({ link: m[2], children: [new TextRun({ text: m[1], ...base, color: C.link, underline: {} })] }));
      last = m.index + m[0].length;
    }
    if (last < text.length) parts.push(...runs(text.slice(last), base));
    return parts;
  }
  const out = []; const re = /(\*\*[^*]+\*\*|\*[^*\s][^*]*\*)/g; let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    const t = m[0];
    out.push(t.startsWith("**") ? new TextRun({ text: t.slice(2, -2), bold: true, ...base }) : new TextRun({ text: t.slice(1, -1), italics: true, ...base }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...base }));
  return out.length ? out : [new TextRun({ text: "", ...base })];
}
const clean = (s) => s.replace(/`/g, "").trim();
const plain = (s) => clean(s || "").replace(/\*\*/g, "").replace(/\*/g, "").replace(/\[([^\]]+)\]\([^)]+\)/g, "$1");

// ---------- Nummerierung der Überschriften (auch für das Inhaltsverzeichnis) ----------
const lines = md.split("\n");
const firstH1 = lines.findIndex((l) => /^#\s/.test(l));
function numberHeadings() {
  const cnt = [0, 0, 0]; const res = {};
  lines.forEach((l, idx) => {
    const m = l.match(/^(#{1,3})\s+(.*)$/);
    if (!m || idx === firstH1) return;
    const lvl = m[1].length; let t = m[2]; const un = /\{-\}\s*$/.test(t); t = t.replace(/\s*\{-\}\s*$/, "");
    if (un) { res[idx] = { lvl, num: "", text: plain(t) }; return; }
    cnt[lvl - 1]++; for (let k = lvl; k < 3; k++) cnt[k] = 0;
    res[idx] = { lvl, num: cnt.slice(0, lvl).join("."), text: plain(t) };
  });
  return res;
}
const H = numberHeadings();

const figs = [], tabs = [];
lines.forEach((l) => {
  const mi = l.match(/^!\[(.*)\]\((.*)\)/); if (mi) figs.push(plain(mi[1]));
  const mt = l.match(/^Tab:\s*(.*)$/); if (mt) tabs.push(plain(mt[1]));
});
const marks = []; // Texte, deren Seitenzahl im 2. Durchlauf gesucht wird

// ---------- Bausteine ----------
const children = [];
const bLine = { style: BorderStyle.SINGLE, size: 4, color: C.line };

const para = (t, o = {}) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 140, line: LINE }, children: runs(clean(t)), ...o });

let numInstance = 0, listBreak = true;
function listItem(text, ordered, level) {
  if (ordered && level === 0 && listBreak) { numInstance++; listBreak = false; }
  if (!ordered && level === 0) listBreak = true;
  return new Paragraph({
    numbering: ordered ? { reference: "numbers", level, instance: numInstance } : { reference: "bullets", level },
    alignment: AlignmentType.LEFT, spacing: { after: 60, line: LINE }, children: runs(clean(text)),
  });
}

let figNo = 0, tabNo = 0;
function caption(kind, no, title) {
  marks.push(`${kind} ${no}: ${title}`);
  return new Paragraph({ keepNext: kind === "Tabelle", spacing: { before: kind === "Tabelle" ? 160 : 60, after: kind === "Tabelle" ? 80 : 40 },
    children: [new TextRun({ text: `${kind} ${no}: `, bold: true, size: SZ.small }), new TextRun({ text: title, size: SZ.small })] });
}

function table(rows) {
  const header = rows[0], body = rows.slice(1), n = header.length;
  const CH = 100, PAD = 240;
  const minW = header.map((_, i) => Math.min(Math.max(...rows.map((r) => Math.max(0, ...plain(r[i]).split(/[\s/–-]+/).map((w) => w.length)))) * CH + PAD, CW * 0.42));
  const desire = header.map((_, i) => { const l = rows.map((r) => plain(r[i]).length); return Math.min(Math.max(l.reduce((a, b) => a + b, 0) / l.length, 4), 70); });
  const sumMin = minW.reduce((a, b) => a + b, 0), sumDes = desire.reduce((a, b) => a + b, 0);
  const widths = sumMin >= CW ? minW.map((w) => Math.floor(w * CW / sumMin)) : minW.map((w, i) => Math.floor(w + (CW - sumMin) * desire[i] / sumDes));
  widths[widths.length - 1] += CW - widths.reduce((a, b) => a + b, 0);
  const keep = rows.length <= 12;
  const isNum = (s) => { const t = plain(s); return !/^\d{4}$/.test(t) && /^[-+−~]?\s?(ca\.\s|rd\.\s)?[\d.,]+\s?(%|€|T€|Mio\.? €|m²|ha|WE|P\.?)?$/.test(t); };
  const cell = (txt, w, head, last) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: head ? { type: ShadingType.CLEAR, color: "auto", fill: C.head } : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 }, verticalAlign: VerticalAlign.TOP,
    borders: { top: bLine, bottom: bLine, left: bLine, right: bLine },
    children: [new Paragraph({ alignment: !head && isNum(txt) ? AlignmentType.RIGHT : AlignmentType.LEFT, keepNext: keep && !last,
      spacing: { after: 0, line: 240 }, children: runs(clean(txt), head ? { size: SZ.table, bold: true } : { size: SZ.table }) })],
  });
  return new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: widths, layout: TableLayoutType.FIXED,
    rows: [new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, widths[i], true, false)) }),
      ...body.map((r, ri) => new TableRow({ cantSplit: true, children: Array.from({ length: n }, (_, i) => cell(r[i] || "", widths[i], false, ri === body.length - 1)) }))],
  });
}

function box(title, blines) {
  listBreak = true;
  const inner = [];
  if (title) inner.push(new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: title, bold: true })] }));
  for (const l of blines) {
    const m = l.match(/^(\s*)([-*]|\d+\.)\s+(.*)$/);
    if (m) inner.push(listItem(m[3], /\d+\./.test(m[2]), m[1].length >= 2 ? 1 : 0));
    else if (l.trim()) inner.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 80, line: LINE }, children: runs(clean(l)) }));
  }
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [CW], rows: [new TableRow({ cantSplit: true, children: [new TableCell({
    width: { size: CW, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: "auto", fill: C.box },
    margins: { top: 120, bottom: 100, left: 160, right: 160 }, borders: { top: bLine, bottom: bLine, left: bLine, right: bLine }, children: inner })] })] });
}

function image(title, file, widthCm) {
  const buf = fs.readFileSync(path.resolve(baseDir, file));
  const wpx = buf.readUInt32BE(16), hpx = buf.readUInt32BE(20);
  const wPt = Math.min((widthCm || 15) / 2.54 * 72, CW / 20), hPt = wPt * hpx / wpx;
  figNo++;
  return [new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 120, after: 40, line: 240, lineRule: LineRuleType.AUTO },
      children: [new ImageRun({ type: "png", data: buf, transformation: { width: Math.round(wPt * 96 / 72), height: Math.round(hPt * 96 / 72) }, altText: { title, description: title, name: path.basename(file) } })] }),
    caption("Abbildung", figNo, title)];
}

function listing(kind, items) {
  const title = kind === "Abbildung" ? "Abbildungsverzeichnis" : "Tabellenverzeichnis";
  marks.push(title);
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, children: [new TextRun(title)] }));
  items.forEach((t, k) => {
    const label = `${kind} ${k + 1}: ${t}`;
    children.push(new Paragraph({ spacing: { after: 60 }, indent: { left: 1418, hanging: 1418 },
      tabStops: [{ type: TabStopType.LEFT, position: 1418 }, { type: TabStopType.RIGHT, position: CW, leader: "dot" }],
      children: [new TextRun({ text: `${kind} ${k + 1}:`, size: SZ.table }), new TextRun({ text: `\t${t}\t${pageMap[label] || ""}`, size: SZ.table })] }));
  });
}

// ---------- Parser ----------
for (let i = 0; i < lines.length; ) {
  const line = lines[i];
  if (!/^(\s*)([-*]|\d+\.)\s+/.test(line) && line.trim()) listBreak = true;
  if (i === firstH1) { i++; continue; }
  if (/^---pagebreak---\s*$/.test(line)) { children.push(new Paragraph({ children: [new PageBreak()] })); i++; continue; }
  if (/^:::toc\s*$/.test(line)) {
    children.push(new Paragraph({ spacing: { after: 240 }, children: [new TextRun({ text: "Inhaltsverzeichnis", bold: true, size: SZ.h1 })] }));
    Object.keys(H).map(Number).sort((a, b) => a - b).forEach((idx) => {
      const h = H[idx]; if (h.lvl > 2) return;
      const pg = pageMap[h.text] != null ? String(pageMap[h.text]) : "";
      const ind = h.lvl === 1 ? 0 : 567;
      children.push(new Paragraph({ spacing: { before: h.lvl === 1 ? 120 : 20, after: 20 }, indent: { left: ind + 680, hanging: 680 },
        tabStops: [{ type: TabStopType.LEFT, position: ind + 680 }, { type: TabStopType.RIGHT, position: CW, leader: "dot" }],
        children: [new TextRun({ text: h.num, bold: h.lvl === 1 }), new TextRun({ text: `\t${h.text}\t${pg}`, bold: h.lvl === 1 })] }));
    });
    for (const [t, dir] of [["Abbildungsverzeichnis", ":::abbildungsverzeichnis"], ["Tabellenverzeichnis", ":::tabellenverzeichnis"]]) {
      if (md.includes(dir)) children.push(new Paragraph({ spacing: { before: 120, after: 20 }, indent: { left: 680, hanging: 680 },
        tabStops: [{ type: TabStopType.LEFT, position: 680 }, { type: TabStopType.RIGHT, position: CW, leader: "dot" }],
        children: [new TextRun({ text: `\t${t}\t${pageMap[t] || ""}`, bold: true })] }));
    }
    i++; continue;
  }
  if (/^:::abbildungsverzeichnis\s*$/.test(line)) { listing("Abbildung", figs); i++; continue; }
  if (/^:::tabellenverzeichnis\s*$/.test(line)) { listing("Tabelle", tabs); i++; continue; }
  if (/^:::box/.test(line)) {
    const title = line.replace(/^:::box\s*/, "").trim(); const buf = []; i++;
    while (i < lines.length && !/^:::\s*$/.test(lines[i])) { buf.push(lines[i]); i++; }
    i++; children.push(box(title, buf)); children.push(new Paragraph({ spacing: { after: 100 }, children: [] })); continue;
  }
  let m;
  if ((m = line.match(/^(#{1,3})\s+(.*)$/))) {
    const h = H[i];
    marks.push(h.text);
    const brk = h.lvl === 1 && (!(cover && cover.chapterBreaks === false) || !firstChapter);
    if (h.lvl === 1) firstChapter = true;
    const hl = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3][h.lvl - 1];
    children.push(new Paragraph({ heading: hl, pageBreakBefore: brk, ...(h.lvl === 1 && !brk ? { spacing: { before: 480, after: 240 } } : {}), tabStops: [{ type: TabStopType.LEFT, position: 680 }],
      children: [new TextRun(h.num ? `${h.num}\t${h.text}` : h.text)] }));
    i++; continue;
  }
  if ((m = line.match(/^Tab:\s*(.*)$/))) { tabNo++; children.push(caption("Tabelle", tabNo, plain(m[1]))); i++; continue; }
  if ((m = line.match(/^!\[(.*)\]\((.*)\)(\{(\d+)\})?/))) { children.push(...image(plain(m[1]), m[2], m[4] ? +m[4] : 15)); i++; continue; }
  if (/^Quelle:/.test(line)) { children.push(new Paragraph({ spacing: { after: 220 }, children: runs(clean(line), { size: SZ.small, color: C.grey }) })); i++; continue; }
  if (/^\|/.test(line)) {
    const rows = [];
    while (i < lines.length && /^\|/.test(lines[i])) {
      const cells = lines[i].trim().replace(/^\||\|$/g, "").split("|").map((s) => s.trim());
      if (!cells.every((c) => /^:?-+:?$/.test(c))) rows.push(cells); i++;
    }
    children.push(table(rows));
    const next = lines.slice(i).find((l) => l.trim()) || "";
    if (!next.startsWith("Quelle:")) children.push(new Paragraph({ spacing: { after: 160 }, children: [] }));
    else children.push(new Paragraph({ spacing: { after: 40 }, children: [] }));
    continue;
  }
  if (/^>\s?/.test(line)) {
    const buf = []; while (i < lines.length && /^>\s?/.test(lines[i])) { buf.push(lines[i].replace(/^>\s?/, "")); i++; }
    children.push(box("", buf)); children.push(new Paragraph({ spacing: { after: 100 }, children: [] })); continue;
  }
  if ((m = line.match(/^(\s*)([-*]|\d+\.)\s+(.*)$/))) { children.push(listItem(m[3], /\d+\./.test(m[2]), m[1].length >= 2 ? 1 : 0)); i++; continue; }
  if (/^\s*$/.test(line) || /^---\s*$/.test(line)) { i++; continue; }
  const buf = [line]; i++;
  while (i < lines.length && lines[i].trim() && !/^(#|\||>|\s*[-*]\s|\s*\d+\.\s|:::|!\[|---|Tab:|Quelle:)/.test(lines[i])) { buf.push(lines[i]); i++; }
  const t = clean(buf.join(" "));
  let j = i; while (j < lines.length && !lines[j].trim()) j++;
  const beforeFloat = j < lines.length && /^(Tab:|!\[)/.test(lines[j]) && t.length < 400;
  children.push(para(t, { keepNext: (/:$/.test(t) && t.length < 160) || beforeFloat }));
}

// ---------- Deckblatt ----------
const coverKids = [];
if (cover) {
  coverKids.push(new Paragraph({ spacing: { before: 2200, after: 160 }, children: [new TextRun({ text: cover.overline || "", size: 22, color: C.grey })] }));
  coverKids.push(new Paragraph({ spacing: { after: 160 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.text, space: 10 } }, children: [new TextRun({ text: cover.title, bold: true, size: 52 })] }));
  coverKids.push(new Paragraph({ spacing: { before: 200, after: 120, line: 320 }, children: [new TextRun({ text: cover.subtitle || "", size: 28 })] }));
  coverKids.push(new Paragraph({ spacing: { before: 600, after: 2400 }, children: [new TextRun({ text: cover.doctype || "", bold: true, size: 28 })] }));
  for (const [k, v] of cover.meta || []) coverKids.push(new Paragraph({ spacing: { after: 60 }, indent: { left: 2268, hanging: 2268 }, tabStops: [{ type: TabStopType.LEFT, position: 2268 }], children: [new TextRun({ text: k, size: 20, color: C.grey }), new TextRun({ text: "\t" + v, size: 20 })] }));
  if (cover.note) { coverKids.push(new Paragraph({ spacing: { before: 600 }, children: [] })); coverKids.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { line: 276 }, children: runs(cover.note, { size: 18, color: C.grey }) })); }
}

// ---------- Dokument ----------
const sectionPage = { size: { width: PAGE_W, height: PAGE_H }, margin: { top: 1418, bottom: 1247, left: ML, right: MR, header: 709, footer: 567 } };
const doc = new Document({
  creator: "", lastModifiedBy: "", title: cover ? cover.title : "", description: "",
  hyphenation: { autoHyphenation: true, hyphenationZone: 357, consecutiveHyphenLimit: 2 },
  styles: {
    default: { document: { run: { font: FONT, size: SZ.body, color: C.text }, paragraph: { spacing: { line: LINE } } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: SZ.h1, bold: true, color: C.text }, paragraph: { spacing: { before: 0, after: 240 }, outlineLevel: 0, keepNext: true } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: SZ.h2, bold: true, color: C.text }, paragraph: { spacing: { before: 300, after: 120 }, outlineLevel: 1, keepNext: true } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: SZ.h3, bold: true, color: C.text }, paragraph: { spacing: { before: 220, after: 80 }, outlineLevel: 2, keepNext: true } },
    ],
  },
  numbering: { config: [
    { reference: "bullets", levels: [
      { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 357, hanging: 357 } } } },
      { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 714, hanging: 357 } } } } ] },
    { reference: "numbers", levels: [
      { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 357, hanging: 357 } } } },
      { level: 1, format: LevelFormat.LOWER_LETTER, text: "%2)", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 714, hanging: 357 } } } } ] },
  ] },
  sections: [
    ...(cover ? [{ properties: { page: sectionPage }, children: coverKids }] : []),
    { properties: { page: sectionPage },
      headers: { default: new Header({ children: [new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.line, space: 4 } }, tabStops: [{ type: TabStopType.RIGHT, position: CW }],
        children: [new TextRun({ text: headerTitle, size: SZ.small, color: C.grey }), new TextRun({ text: "\t" + ((cover && cover.headerRight) || ""), size: SZ.small, color: C.grey })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: ["Seite ", PageNumber.CURRENT], size: SZ.small, color: C.grey })] })] }) },
      children },
  ],
});
fs.writeFileSync(outFile + ".marks.json", JSON.stringify(marks));
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(outFile, b); console.log("geschrieben:", outFile); });
