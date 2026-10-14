// Markdown -> gelayoutetes Word (docx-js). Single source: Markdown im Repo.
// Unterstützt: #/##/### Überschriften, Absätze mit **fett**/*kursiv*, Listen (-, 1.),
// GFM-Tabellen, > Infobox, :::box Titel ... :::, :::kpi ... :::, ![Bildunterschrift](pfad){breite_cm},
// ---pagebreak---, <!-- kommentare --> werden ignoriert.
// Aufruf: node md2docx.js input.md output.docx "Kopfzeilentitel" [cover.json]
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, ExternalHyperlink, LevelFormat, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, TableOfContents, TabStopType,
  TableLayoutType,
} = require("docx");

const [, , inFile, outFile, headerTitle = "", coverFile, tocMapFile] = process.argv;
const tocMap = tocMapFile && fs.existsSync(tocMapFile) ? JSON.parse(fs.readFileSync(tocMapFile, "utf8")) : {};
const tocEntries = []; // für Pass 1: Liste der Überschriften
const md = fs.readFileSync(inFile, "utf8").replace(/<!--[\s\S]*?-->/g, "");
const baseDir = path.dirname(path.resolve(inFile));
const cover = coverFile ? JSON.parse(fs.readFileSync(coverFile, "utf8")) : null;

// ---------- Design-Tokens ----------
const C = {
  navy: "0D366B", blue: "2A78D6", blueLight: "CDE2FB", blueTint: "EEF5FD",
  ink: "1F2328", ink2: "52514E", muted: "8A8984", rule: "D0D7DE",
  zebra: "F5F7FA", white: "FFFFFF", accent: "EB6834",
};
const FONT = "Inter";
const PAGE_W = 11906, PAGE_H = 16838, MARGIN_L = 1247, MARGIN_R = 1247; // A4, 2,2 cm
const CONTENT_W = PAGE_W - MARGIN_L - MARGIN_R; // 9412 DXA

// ---------- Inline-Parser: **fett**, *kursiv* ----------
function runs(text, base = {}) {
  // Markdown-Links [Text](URL) -> Hyperlink
  const linkRe = /\[([^\]]+)\]\((https?:[^)]+)\)/g;
  if (linkRe.test(text)) {
    const parts = [];
    let last = 0, m;
    linkRe.lastIndex = 0;
    while ((m = linkRe.exec(text)) !== null) {
      if (m.index > last) parts.push(...runs(text.slice(last, m.index), base));
      parts.push(new ExternalHyperlink({ link: m[2], children: [new TextRun({ text: m[1], ...base, color: "2A78D6", underline: {} })] }));
      last = m.index + m[0].length;
    }
    if (last < text.length) parts.push(...runs(text.slice(last), base));
    return parts;
  }
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*\s][^*]*\*)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    const tok = m[0];
    if (tok.startsWith("**")) out.push(new TextRun({ text: tok.slice(2, -2), bold: true, ...base }));
    else out.push(new TextRun({ text: tok.slice(1, -1), italics: true, ...base }));
    last = m.index + tok.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...base }));
  return out.length ? out : [new TextRun({ text: "", ...base })];
}
const clean = (s) => s.replace(/`/g, "").trim();

// ---------- Bausteine ----------
const children = [];
const thin = { style: BorderStyle.SINGLE, size: 4, color: C.rule };
const none = { style: BorderStyle.NONE, size: 0, color: C.white };

function para(text, opts = {}) {
  const t = clean(text);
  const pseudoHead = t.startsWith("**") && t.length < 90;
  return new Paragraph({ children: runs(t, opts.run || {}), keepNext: pseudoHead, spacing: { after: 120, line: 288 }, ...opts.p });
}

function heading(level, text) {
  const map = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3 };
  const p = { heading: map[level], children: runs(clean(text)) };
  if (level === 1) p.pageBreakBefore = true;
  return new Paragraph(p);
}

let numInstance = 0, listBreak = true;
function listItem(text, ordered, level) {
  if (ordered && level === 0 && listBreak) { numInstance++; listBreak = false; }
  if (!ordered && level === 0) listBreak = true;
  return new Paragraph({
    numbering: ordered ? { reference: "numbers", level, instance: numInstance } : { reference: "bullets", level },
    children: runs(clean(text)),
    spacing: { after: 60, line: 276 },
  });
}

function table(rows) {
  const header = rows[0];
  const body = rows.slice(1);
  const n = header.length;
  // Spaltenbreiten: Mindestbreite nach längstem Wort, Rest nach mittlerer Textlänge
  const CH = 96, PAD = 230;
  const plain = (t) => clean(t || "").replace(/\*\*/g, "").replace(/\*/g, "");
  const minW = header.map((_, i) => {
    const longest = Math.max(...rows.map((r) => Math.max(0, ...plain(r[i]).split(/[\s/–-]+/).map((w) => w.length))));
    return Math.min(longest * CH + PAD, CONTENT_W * 0.42);
  });
  const desire = header.map((_, i) => {
    const l = rows.map((r) => plain(r[i]).length);
    return Math.min(Math.max(l.reduce((a, b) => a + b, 0) / l.length, 4), 70);
  });
  let widths;
  const sumMin = minW.reduce((a, b) => a + b, 0);
  if (sumMin >= CONTENT_W) widths = minW.map((w) => Math.floor(w * CONTENT_W / sumMin));
  else {
    const rest = CONTENT_W - sumMin, sd = desire.reduce((a, b) => a + b, 0);
    widths = minW.map((w, i) => Math.floor(w + rest * desire[i] / sd));
  }
  widths[widths.length - 1] += CONTENT_W - widths.reduce((a, b) => a + b, 0);
  const keepTogether = rows.length <= 12;
  const isNum = (s) => { const t = clean(s).replace(/\*\*/g, ""); return !/^\d{4}$/.test(t) && /^[-+−≈~]?\s?[\d.,]+\s?(%|€|T€|Mio\.? €|m²|ha|WE|P\.?)?$/.test(t); };
  const mkCell = (txt, w, isHead, rowIdx, last) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: isHead ? C.navy : rowIdx % 2 ? C.zebra : C.white },
    margins: { top: 70, bottom: 70, left: 100, right: 100 },
    verticalAlign: VerticalAlign.CENTER,
    borders: { top: thin, bottom: thin, left: none, right: none },
    children: [new Paragraph({
      alignment: !isHead && isNum(txt) ? AlignmentType.RIGHT : AlignmentType.LEFT,
      keepNext: keepTogether && !last,
      spacing: { after: 0, line: 252 },
      children: runs(clean(txt), isHead ? { bold: true, color: C.white, size: 17 } : { size: 17 }),
    })],
  });
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({ tableHeader: true, children: header.map((h, i) => mkCell(h, widths[i], true, 0, false)) }),
      ...body.map((r, ri) => new TableRow({ cantSplit: true, children: Array.from({ length: n }, (_, i) => mkCell(r[i] || "", widths[i], false, ri, ri === body.length - 1)) })),
    ],
  });
}

function box(title, lines) {
  listBreak = true;
  const inner = [];
  if (title) inner.push(new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: title.toUpperCase(), bold: true, color: C.navy, size: 17, characterSpacing: 20 })] }));
  for (const l of lines) {
    const m = l.match(/^(\s*)([-*]|\d+\.)\s+(.*)$/);
    if (m) inner.push(listItem(m[3], /\d+\./.test(m[2]), m[1].length >= 2 ? 1 : 0));
    else if (l.trim()) inner.push(new Paragraph({ spacing: { after: 80, line: 276 }, children: runs(clean(l)) }));
  }
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [new TableCell({
      width: { size: CONTENT_W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: C.blueTint },
      margins: { top: 160, bottom: 120, left: 220, right: 220 },
      borders: { top: none, bottom: none, right: none, left: { style: BorderStyle.SINGLE, size: 24, color: C.blue } },
      children: inner,
    })] })],
  });
}

function kpis(lines) {
  // Zeilen: "Wert | Beschriftung"
  const items = lines.filter((l) => l.includes("|")).map((l) => l.split("|").map((s) => s.trim()));
  const n = items.length;
  const gap = 120;
  const w = Math.floor((CONTENT_W - gap * (n - 1)) / n);
  const cells = [];
  const widths = [];
  items.forEach(([val, label], i) => {
    cells.push(new TableCell({
      width: { size: w, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: C.navy },
      margins: { top: 160, bottom: 160, left: 160, right: 120 },
      borders: { top: none, bottom: none, left: none, right: none },
      children: [
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: val, bold: true, color: C.white, size: val.length > 9 ? 24 : 30 })] }),
        new Paragraph({ spacing: { after: 0, line: 240 }, children: [new TextRun({ text: label, color: C.blueLight, size: 15 })] }),
      ],
    }));
    widths.push(w);
    if (i < n - 1) {
      cells.push(new TableCell({ width: { size: gap, type: WidthType.DXA }, borders: { top: none, bottom: none, left: none, right: none }, children: [new Paragraph("")] }));
      widths.push(gap);
    }
  });
  widths[widths.length - 1] += CONTENT_W - widths.reduce((a, b) => a + b, 0);
  return new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: widths, rows: [new TableRow({ children: cells })] });
}

function image(caption, file, widthCm) {
  const full = path.resolve(baseDir, file);
  const buf = fs.readFileSync(full);
  // PNG-Größe lesen
  const wpx = buf.readUInt32BE(16), hpx = buf.readUInt32BE(20);
  const maxWpt = (widthCm || 16) / 2.54 * 72;
  const wPt = Math.min(maxWpt, CONTENT_W / 20);
  const hPt = wPt * hpx / wpx;
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 },
      children: [new ImageRun({ type: "png", data: buf, transformation: { width: Math.round(wPt * 96 / 72), height: Math.round(hPt * 96 / 72) }, altText: { title: caption, description: caption, name: path.basename(file) } })] }),
    new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: caption, italics: true, color: C.ink2, size: 16 })] }),
  ];
}

// ---------- Parser ----------
const lines = md.split("\n");
let i = 0;
let docTitle = "";
while (i < lines.length) {
  const line = lines[i];
  if (!/^(\s*)([-*]|\d+\.)\s+/.test(line) && line.trim()) listBreak = true;
  if (/^---pagebreak---\s*$/.test(line)) { children.push(new Paragraph({ children: [new PageBreak()] })); i++; continue; }
  if (/^:::toc\s*$/.test(line)) {
    children.push(new Paragraph({ spacing: { before: 120, after: 240 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.blue, space: 6 } }, children: [new TextRun({ text: "Inhalt", bold: true, size: 34, color: C.navy, font: "Inter Display" })] }));
    // Überschriften vorab einsammeln
    for (const l of lines) {
      const mm = l.match(/^(#{1,2})\s+(.*)$/);
      if (mm && !(mm[1].length === 1 && l === lines.find((x) => /^#\s/.test(x)))) {
        const t = clean(mm[2]).replace(/\*\*/g, "");
        const lvl = mm[1].length;
        const pg = tocMap[t] != null ? String(tocMap[t]) : "00";
        children.push(new Paragraph({
          spacing: { before: lvl === 1 ? 140 : 20, after: 20 },
          indent: { left: lvl === 1 ? 0 : 360 },
          tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W, leader: "dot" }],
          children: [new TextRun({ text: t, bold: lvl === 1, color: lvl === 1 ? C.navy : C.ink, size: lvl === 1 ? 20 : 18 }), new TextRun({ text: "\t" + pg, bold: lvl === 1, color: C.ink2, size: lvl === 1 ? 20 : 18 })],
        }));
      }
    }
    i++; continue;
  }
  if (/^:::(box|kpi)/.test(line)) {
    const kind = line.match(/^:::(box|kpi)/)[1];
    const title = line.replace(/^:::(box|kpi)\s*/, "").trim();
    const buf = []; i++;
    while (i < lines.length && !/^:::\s*$/.test(lines[i])) { buf.push(lines[i]); i++; }
    i++;
    children.push(kind === "kpi" ? kpis(buf) : box(title, buf));
    children.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
    continue;
  }
  let m;
  if ((m = line.match(/^(#{1,3})\s+(.*)$/))) {
    const lvl = m[1].length;
    if (lvl === 1 && !docTitle) { docTitle = m[2]; i++; continue; } // erster H1 = Dokumenttitel (Deckblatt)
    tocEntries.push(clean(m[2]).replace(/\*\*/g, ""));
    children.push(heading(lvl, m[2])); i++; continue;
  }
  if ((m = line.match(/^!\[(.*)\]\((.*)\)(\{(\d+)\})?/))) { children.push(...image(m[1], m[2], m[4] ? +m[4] : 16)); i++; continue; }
  if (/^\|/.test(line)) {
    const rows = [];
    while (i < lines.length && /^\|/.test(lines[i])) {
      const cells = lines[i].trim().replace(/^\||\|$/g, "").split("|").map((s) => s.trim());
      if (!cells.every((c) => /^:?-+:?$/.test(c))) rows.push(cells);
      i++;
    }
    children.push(table(rows));
    children.push(new Paragraph({ spacing: { after: 160 }, children: [] }));
    continue;
  }
  if (/^>\s?/.test(line)) {
    const buf = [];
    while (i < lines.length && /^>\s?/.test(lines[i])) { buf.push(lines[i].replace(/^>\s?/, "")); i++; }
    children.push(box("", buf));
    children.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
    continue;
  }
  if ((m = line.match(/^(\s*)([-*]|\d+\.)\s+(.*)$/))) {
    children.push(listItem(m[3], /\d+\./.test(m[2]), m[1].length >= 2 ? 1 : 0)); i++; continue;
  }
  if (/^\s*$/.test(line) || /^---\s*$/.test(line)) { i++; continue; }
  // Absatz (mehrzeilig zusammenführen)
  const buf = [line];
  i++;
  while (i < lines.length && lines[i].trim() && !/^(#|\||>|\s*[-*]\s|\s*\d+\.\s|:::|!\[|---)/.test(lines[i])) { buf.push(lines[i]); i++; }
  children.push(para(buf.join(" ")));
}

// ---------- Deckblatt ----------
const coverChildren = [];
if (cover) {
  const band = (fill, h) => new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
    rows: [new TableRow({ height: { value: h, rule: "exact" }, children: [new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: "auto", fill }, borders: { top: none, bottom: none, left: none, right: none }, children: [new Paragraph("")] })] })],
  });
  coverChildren.push(new Paragraph({ spacing: { before: 600, after: 120 }, children: [new TextRun({ text: cover.kicker || "", bold: true, color: C.blue, size: 20, characterSpacing: 40 })] }));
  coverChildren.push(band(C.navy, 120));
  coverChildren.push(new Paragraph({ spacing: { before: 480, after: 240, line: 312 }, children: [new TextRun({ text: cover.title, bold: true, color: C.navy, size: 56, font: "Inter Display" })] }));
  coverChildren.push(new Paragraph({ spacing: { after: 600, line: 300 }, children: [new TextRun({ text: cover.subtitle || "", color: C.ink2, size: 26 })] }));
  for (const [k, v] of cover.meta || []) {
    coverChildren.push(new Paragraph({ spacing: { after: 60 }, tabStops: [{ type: TabStopType.LEFT, position: 2600 }], children: [new TextRun({ text: k, bold: true, color: C.navy, size: 18 }), new TextRun({ text: "\t" + v, color: C.ink, size: 18 })] }));
  }
  coverChildren.push(new Paragraph({ spacing: { before: 2600 }, children: [] }));
  if (cover.note) coverChildren.push(box("", [cover.note]));
  coverChildren.push(new Paragraph({ children: [new PageBreak()] }));
}

// ---------- Dokument ----------
const doc = new Document({
  creator: "EMRIQ Kelmis",
  title: cover ? cover.title : docTitle,
  styles: {
    default: { document: { run: { font: FONT, size: 19, color: C.ink } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 34, bold: true, color: C.navy, font: "Inter Display" },
        paragraph: { spacing: { before: 120, after: 240 }, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.blue, space: 6 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 25, bold: true, color: C.navy, font: "Inter Display" },
        paragraph: { spacing: { before: 320, after: 140 }, outlineLevel: 1, keepNext: true } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 21, bold: true, color: C.blue },
        paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 2, keepNext: true } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "▪", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 260 } }, run: { color: C.blue } } },
        { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 800, hanging: 260 } } } },
      ] },
      { reference: "numbers", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 300 } }, run: { bold: true, color: C.navy } } },
        { level: 1, format: LevelFormat.LOWER_LETTER, text: "%2)", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 800, hanging: 300 } } } },
      ] },
    ],
  },
  sections: [
    ...(cover ? [{
      properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: 1300, bottom: 1200, left: MARGIN_L, right: MARGIN_R } } },
      children: coverChildren.slice(0, -1),
    }] : []),
    {
      properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: 1400, bottom: 1300, left: MARGIN_L, right: MARGIN_R, header: 600, footer: 600 } } },
      headers: { default: new Header({ children: [new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.rule, space: 4 } },
        tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
        children: [new TextRun({ text: headerTitle, color: C.ink2, size: 15 }), new TextRun({ text: "\t" + (cover && cover.headerRight || ""), color: C.muted, size: 15 })],
      })] }) },
      footers: { default: new Footer({ children: [new Paragraph({
        tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
        children: [new TextRun({ text: (cover && cover.footerLeft) || "", color: C.muted, size: 15 }),
          new TextRun({ children: ["\tSeite ", PageNumber.CURRENT, " von ", PageNumber.TOTAL_PAGES], color: C.ink2, size: 15 })],
      })] }) },
      children,
    },
  ],
});

fs.writeFileSync(outFile + ".headings.json", JSON.stringify(tocEntries));
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(outFile, b); console.log("geschrieben:", outFile); });
