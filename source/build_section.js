/**
 * build_section.js
 * ----------------
 * Builds the review script for one course section, plus production files.
 *
 * Input:  a section content file (JSON), see source/content/section-01.json
 * Output:
 *   1. <out.docx>   Review document: Hebrew author notes, then per lesson a
 *                   Slide | Narration table, worksheets, quizzes, answer keys.
 *   2. <narrationDir>/lesson-<id>.txt   Narration only, one file per lesson,
 *                   ready to paste into an AI voice tool.
 *   3. <narrationDir>/worksheet-<id>.txt   Worksheet text for exercises.
 *
 * Usage:
 *   node build_section.js <content.json> <out.docx> <narrationDir>
 *
 * JSON shape (all text fields are plain strings):
 * {
 *   "section": 1, "title": "The Problem", "version": "1", "date": "20.11.2026",
 *   "notes_he": ["..."],
 *   "lessons": [{
 *     "id": "1.2", "title": "...", "type": "Video|Exercise|Quiz", "preview": true,
 *     "slides": [{ "slide": ["SLIDE 1", "..."], "narration": ["...", "..."] }],
 *     "worksheet": { "title": "...", "instructions": "...", "items": ["..."],
 *                    "answers": ["..."], "part2": { "title": "...", "steps": ["..."] } },
 *     "quiz": [{ "q": "...", "options": ["..."], "answer": 0, "why": "..." }]
 *   }]
 * }
 */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, LevelFormat, PageBreak,
} = require("docx");

const [, , SRC, OUT, NARR_DIR] = process.argv;
if (!SRC || !OUT || !NARR_DIR) {
  console.error("Usage: node build_section.js <content.json> <out.docx> <narrationDir>");
  process.exit(1);
}
const C = JSON.parse(fs.readFileSync(SRC, "utf8"));
const FONT = "Arial";
const WPM = 140; // assumed speaking rate for length estimates (words per minute)

// ---------- helpers ----------
const he = (text, bold = false) => new Paragraph({
  bidirectional: true, alignment: AlignmentType.RIGHT, spacing: { after: 100 },
  children: [new TextRun({ text, bold, rightToLeft: true, font: FONT })],
});
const heb = (text) => new Paragraph({
  bidirectional: true, alignment: AlignmentType.RIGHT, spacing: { after: 60 },
  numbering: { reference: "rtl-bullets", level: 0 },
  children: [new TextRun({ text, rightToLeft: true, font: FONT })],
});
const en = (text, bold = false, italics = false) => new Paragraph({
  spacing: { after: 100 }, children: [new TextRun({ text, bold, italics, font: FONT })],
});
const enb = (text) => new Paragraph({
  spacing: { after: 60 }, numbering: { reference: "bullets", level: 0 },
  children: [new TextRun({ text, font: FONT })],
});
const h = (text, level) => new Paragraph({
  heading: [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3][level - 1],
  children: [new TextRun({ text })],
});
const cell = (lines, width, header = false) => new TableCell({
  width: { size: width, type: WidthType.DXA },
  shading: header ? { type: ShadingType.CLEAR, color: "auto", fill: "D9E2F3" } : undefined,
  margins: { top: 80, bottom: 80, left: 110, right: 110 },
  children: (Array.isArray(lines) ? lines : [lines]).map(t => new Paragraph({
    spacing: { after: 80 }, children: [new TextRun({ text: t, bold: header, font: FONT, size: 20 })],
  })),
});
const SW = [3000, 6026];
const words = (arr) => arr.join(" ").split(/\s+/).filter(Boolean).length;

// ---------- build ----------
fs.mkdirSync(NARR_DIR, { recursive: true });
const children = [
  new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun(`Section ${C.section}: ${C.title}`)] }),
  en(`Goal-First Planning for New Managers | Draft scripts | Version ${C.version} | ${C.date}`, false, true),
  he("הערות למחבר", true),
  ...C.notes_he.map(heb),
];

let totalWords = 0;
C.lessons.forEach((L, idx) => {
  if (idx > 0) children.push(new Paragraph({ children: [new PageBreak()] }));
  const narr = (L.slides || []).flatMap(s => s.narration);
  const w = words(narr);
  totalWords += w;
  const meta = [`Type: ${L.type}`, `Free preview: ${L.preview ? "Yes" : "No"}`];
  if (w) meta.push(`Narration: about ${w} words, about ${Math.max(1, Math.round(w / WPM))} min at ${WPM} words per minute (estimate)`);
  children.push(h(`Lesson ${L.id}: ${L.title}`, 1), en(meta.join(" | "), false, true));

  if (L.slides && L.slides.length) {
    children.push(new Table({
      width: { size: 9026, type: WidthType.DXA }, columnWidths: SW,
      rows: [
        new TableRow({ tableHeader: true, children: [cell("Slide", SW[0], true), cell("Narration", SW[1], true)] }),
        ...L.slides.map(s => new TableRow({ children: [cell(s.slide, SW[0]), cell(s.narration, SW[1])] })),
      ],
    }));
    fs.writeFileSync(path.join(NARR_DIR, `lesson-${L.id}.txt`),
      `Lesson ${L.id}: ${L.title}\n\n` + narr.join("\n\n") + "\n");
  }

  if (L.worksheet) {
    const W = L.worksheet;
    const lines = [`${W.title}`, "", W.instructions, ""];
    children.push(h(`Worksheet: ${W.title}`, 2), en(W.instructions));
    W.items.forEach((q, i) => { children.push(en(`${i + 1}. ${q}`)); lines.push(`${i + 1}. ${q}`); });
    if (W.part2) {
      children.push(h(W.part2.title, 3));
      lines.push("", W.part2.title);
      W.part2.steps.forEach(s => { children.push(enb(s)); lines.push(`- ${s}`); });
    }
    children.push(h("Answer key", 3));
    lines.push("", "Answer key");
    W.answers.forEach((a, i) => { children.push(en(`${i + 1}. ${a}`)); lines.push(`${i + 1}. ${a}`); });
    fs.writeFileSync(path.join(NARR_DIR, `worksheet-${L.id}.txt`), lines.join("\n") + "\n");
  }

  if (L.quiz) {
    children.push(h("Quiz questions", 2), en("Multiple choice. The correct answer is marked, with the explanation shown to the student after answering.", false, true));
    L.quiz.forEach((q, i) => {
      children.push(en(`${i + 1}. ${q.q}`, true));
      q.options.forEach((o, j) => children.push(enb(`${j === q.answer ? "[Correct] " : ""}${o}`)));
      children.push(en(`Why: ${q.why}`, false, true));
    });
  }
});

children.splice(3 + C.notes_he.length, 0,
  heb(`סך הכל קריינות בפרק: כ-${totalWords} מילים, כלומר בערך ${Math.round(totalWords / WPM)} דקות בקצב של ${WPM} מילים לדקה. זו הערכה, וצריך למדוד בהקלטה.`));

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: "Title", name: "Title", basedOn: "Normal", run: { font: FONT, size: 40, bold: true, color: "1F3864" },
        paragraph: { spacing: { before: 0, after: 120 }, keepNext: true } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 30, bold: true, color: "1F3864" },
        paragraph: { spacing: { before: 240, after: 120 }, keepNext: true, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 25, bold: true, color: "2F5496" },
        paragraph: { spacing: { before: 200, after: 80 }, keepNext: true, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 22, bold: true, color: "2F5496" },
        paragraph: { spacing: { before: 160, after: 60 }, keepNext: true, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 500, hanging: 300 } } } }] },
      { reference: "rtl-bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.RIGHT,
        style: { paragraph: { indent: { right: 500, hanging: 300 } } } }] },
    ],
  },
  sections: [{ properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children }],
});

Packer.toBuffer(doc).then(buf => { fs.writeFileSync(OUT, buf); console.log("written:", OUT, "| narration words:", totalWords); });
