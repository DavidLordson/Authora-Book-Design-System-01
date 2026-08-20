#!/usr/bin/env node
/**
 * Builds a complete formatted .docx from the classified manuscript JSON
 * (produced by a scripts/convert-source-manuscript*.py importer) using the
 * Authora design tokens in src/design-system.js.
 *
 * Each chapter is emitted as its own Word *section* (type: ODD_PAGE), which
 * is what actually makes Word start every chapter on an odd (recto) page —
 * inserting a blank page where needed — rather than just the next page.
 * All flowing body prose (paragraphs, drop-cap text, block quotes) is fully
 * justified; centered elements (titles, epigraphs, headings) are untouched.
 *
 * Usage: node scripts/build-book.js <elements.json> [output.docx]
 */
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, Footer, PageNumber,
  HorizontalPositionAlign, VerticalPositionAlign, FrameAnchorType,
  HeightRule, DropCapType, HeadingLevel, TableOfContents, FrameWrap, SectionType,
} = require("docx");

const { COLORS, FONTS, SIZES, SPACING, PAGE } = require("../src/design-system");

// ─── Shared style builders (mirrors scripts/render-sample-preview.js) ───

function run(text, extra = {}) {
  return new TextRun({ text, font: FONTS.serif, size: SIZES.bodyText, color: COLORS.body, ...extra });
}

function makeChapterLabel(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: SPACING.chapterOpenerBefore, after: SPACING.chapterLabelAfter },
    children: [run(text, { size: SIZES.chapterLabel, italics: true })],
  });
}

function makeChapterTitle(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    heading: HeadingLevel.HEADING_1,
    spacing: { after: SPACING.chapterTitleAfter },
    children: [new TextRun({ text, font: FONTS.script, size: SIZES.chapterTitle, color: COLORS.primary })],
  });
}

function makeEpigraph(text) {
  const quoted = /^[“"]/.test(text) ? text : `“${text}”`;
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: SPACING.epigraphBefore, after: SPACING.epigraphAfter },
    children: [run(quoted, { size: SIZES.epigraph, italics: true })],
  });
}

function makeEpigraphCitation(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: SPACING.epigraphCitationAfter },
    children: [run(text, { size: SIZES.epigraphCitation, italics: true })],
  });
}

function makeTagline(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: SPACING.epigraphCitationAfter },
    children: [run(text, { size: SIZES.chapterSubtitle, bold: true, italics: true })],
  });
}

function makeBodyParagraph(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: SPACING.bodyParagraphAfter },
    children: [run(text)],
  });
}

function makeDropCapParagraph(text) {
  const dropLetter = text.charAt(0);
  const rest = text.slice(1);
  const dropCapPara = new Paragraph({
    frame: {
      type: "alignment",
      alignment: { x: HorizontalPositionAlign.LEFT, y: VerticalPositionAlign.TOP },
      anchor: { horizontal: FrameAnchorType.TEXT, vertical: FrameAnchorType.TEXT },
      width: 800, height: 800, rule: HeightRule.AUTO, wrap: FrameWrap.AROUND,
      dropCap: DropCapType.DROP, lines: SPACING.dropCapLines,
    },
    children: [new TextRun({ text: dropLetter, font: FONTS.serif, color: COLORS.primary })],
  });
  const restPara = new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: SPACING.bodyParagraphAfter },
    children: [run(rest)],
  });
  return [dropCapPara, restPara];
}

// Accepts either the legacy { text, citation } shape or { lines: string[],
// citation } for multi-line quotes (e.g. a four-line prayer with no
// citation at all).
function makeBlockQuote(item) {
  const lines = item.lines || [item.text];
  const hasCitation = !!item.citation;
  const paras = lines.map((line, idx) => {
    const isFirst = idx === 0;
    const isLast = idx === lines.length - 1;
    let quoted = line;
    if (isFirst && !/^[“"]/.test(quoted)) quoted = "“" + quoted;
    if (isLast && !/[”"]$/.test(quoted)) quoted = quoted + "”";
    return new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      indent: { left: 360, right: 360 },
      spacing: {
        before: isFirst ? SPACING.blockQuoteBefore : 0,
        after: isLast ? (hasCitation ? 60 : SPACING.blockQuoteAfter) : 0,
      },
      children: [run(quoted, { size: SIZES.epigraph, italics: true })],
    });
  });
  if (hasCitation) {
    paras.push(
      new Paragraph({
        indent: { left: 360 },
        spacing: { after: SPACING.blockQuoteAfter },
        children: [run(item.citation, { size: SIZES.epigraphCitation, italics: true })],
      })
    );
  }
  return paras;
}

// Faux small-caps section heading (matches the reference's manual Word
// trick: enlarged first letter of each word, smaller true caps for the
// rest) — used both for "TAKE THESE STEPS" / "A PRAYER" style dividers and
// for ordinary in-chapter subheadings.
function makeSectionHeading(text) {
  const words = text.toUpperCase().split(" ");
  const runs = [];
  words.forEach((word, wi) => {
    if (wi > 0) runs.push(run(" ", { size: SIZES.sectionHeading, bold: true }));
    runs.push(run(word[0], { size: SIZES.sectionHeadingCap, bold: true }));
    if (word.length > 1) runs.push(run(word.slice(1), { size: SIZES.sectionHeading, bold: true }));
  });
  return new Paragraph({ spacing: { before: SPACING.sectionHeadingBefore, after: SPACING.sectionHeadingAfter }, children: runs });
}

function makeNumberedItem(n, text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { left: SPACING.bulletIndent, hanging: SPACING.bulletHanging },
    spacing: { after: SPACING.bulletAfter },
    children: [run(`${n}.\t`), run(text)],
  });
}

function makePrayerLine(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: SPACING.bodyParagraphAfter },
    children: [run(text, { italics: true })],
  });
}

function makeSectionBreak() {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: SPACING.sectionBreakBefore, after: SPACING.sectionBreakAfter },
    children: [run("✶")],
  });
}

function makeFooter() {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ font: FONTS.serif, size: SIZES.footer, color: COLORS.muted, children: [PageNumber.CURRENT] })],
      }),
    ],
  });
}

// ─── Front / back matter ───

function buildTitlePage(el) {
  const paras = [];
  if (el.series) {
    paras.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 3200, after: 400 }, children: [run(el.series, { italics: true })] }));
  }
  paras.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: el.series ? 0 : 3200, after: 300 },
      children: [new TextRun({ text: el.title, font: FONTS.script, size: SIZES.chapterTitle + 12, color: COLORS.primary })],
    })
  );
  paras.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 1600 }, children: [run(el.tagline, { italics: true })] }));
  paras.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [run(el.author)] }));
  paras.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [run(el.publisher)] }));
  return paras;
}

function buildCopyrightPage(el) {
  const paras = [new Paragraph({ spacing: { before: 2000, after: 200 }, children: [run(el.lines[0], { bold: true })] })];
  for (const line of el.lines.slice(1)) {
    paras.push(new Paragraph({ spacing: { after: 160 }, children: [run(line, { size: SIZES.footnote })] }));
  }
  return paras;
}

function buildTOC() {
  return [
    new Paragraph({ spacing: { before: 2000, after: 300 }, alignment: AlignmentType.CENTER, children: [run("Contents", { size: SIZES.chapterSubtitle, bold: true })] }),
    new TableOfContents("Contents", { hyperlink: true, headingStyleRange: "1-1" }),
  ];
}

// ─── Chapter body assembly ───

function buildChapter(chapter) {
  const paras = [makeChapterLabel(chapter.label), makeChapterTitle(chapter.title)];

  if (chapter.epigraph) {
    if (chapter.citation) {
      paras.push(makeEpigraph(chapter.epigraph));
      paras.push(makeEpigraphCitation(chapter.citation));
    } else {
      paras.push(makeEpigraph(chapter.epigraph));
      if (chapter.tagline) paras.push(makeTagline(chapter.tagline));
    }
  }

  let dropCapPending = true;
  for (const item of chapter.body) {
    switch (item.type) {
      case "paragraph":
        if (dropCapPending) {
          paras.push(...makeDropCapParagraph(item.text));
          dropCapPending = false;
        } else {
          paras.push(makeBodyParagraph(item.text));
        }
        break;
      case "block-quote":
        paras.push(...makeBlockQuote(item));
        break;
      case "subheading":
        paras.push(makeSectionHeading(item.text));
        break;
      case "section-break":
        paras.push(makeSectionBreak());
        break;
      case "steps":
        paras.push(makeSectionHeading(item.heading));
        item.items.forEach((t, idx) => paras.push(makeNumberedItem(idx + 1, t)));
        break;
      case "prayer":
        paras.push(makeSectionHeading(item.heading));
        item.lines.forEach((t) => paras.push(makePrayerLine(t)));
        break;
    }
  }
  return paras;
}

function buildBackMatter(el) {
  const paras = [];
  for (const item of el.items) {
    if (item.type !== "paragraph") continue;
    const isHeading = ["About the Author", "Also in the Association Series", "Authora Press"].includes(item.text);
    if (isHeading) {
      paras.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: item.text === "About the Author" ? 2000 : 600, after: 200 },
          children: [run(item.text, { bold: true, size: SIZES.chapterSubtitle })],
        })
      );
    } else {
      paras.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, children: [run(item.text, { italics: true })] }));
    }
  }
  return paras;
}

// ─── Main ───

async function build(inputPath, outputPath) {
  const elements = JSON.parse(fs.readFileSync(inputPath, "utf-8"));

  const basePageProps = {
    size: { width: PAGE.width, height: PAGE.height },
    margin: {
      top: PAGE.marginTop, bottom: PAGE.marginBottom,
      left: PAGE.marginInside, right: PAGE.marginOutside,
      header: PAGE.headerDistance, footer: PAGE.footerDistance,
    },
  };

  // Front matter (title/copyright/TOC) all share one ordinary section; each
  // chapter then gets its own section typed ODD_PAGE, which is what makes
  // Word insert a blank page when needed so the chapter starts recto.
  // Back matter closes out in a final ordinary next-page section.
  const sections = [];
  let frontMatter = [];

  for (const el of elements) {
    if (el.type === "title-page") frontMatter.push(...buildTitlePage(el));
    else if (el.type === "copyright-page") frontMatter.push(...buildCopyrightPage(el));
    else if (el.type === "toc") frontMatter.push(...buildTOC());
    else if (el.type === "chapter") {
      sections.push({
        properties: { page: basePageProps, type: SectionType.ODD_PAGE },
        footers: { default: makeFooter() },
        children: buildChapter(el),
      });
    } else if (el.type === "back-matter" && el.items.length) {
      sections.push({
        properties: { page: basePageProps, type: SectionType.NEXT_PAGE },
        footers: { default: makeFooter() },
        children: buildBackMatter(el),
      });
    }
  }

  sections.unshift({
    properties: { page: basePageProps },
    footers: { default: makeFooter() },
    children: frontMatter,
  });

  const doc = new Document({
    creator: "Authora Book Design System",
    description: "Formatted manuscript",
    features: { updateFields: true },
    styles: {
      default: { document: { run: { font: FONTS.serif, size: SIZES.bodyText, color: COLORS.body } } },
      paragraphStyles: [
        {
          id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONTS.script, size: SIZES.chapterTitle, color: COLORS.primary },
          paragraph: { alignment: AlignmentType.CENTER },
        },
      ],
    },
    sections,
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  console.log(`Written to ${outputPath}`);
}

if (require.main === module) {
  const [inputPath, outputPath] = process.argv.slice(2);
  if (!inputPath) {
    console.error("Usage: node scripts/build-book.js <elements.json> [output.docx]");
    process.exit(1);
  }
  build(inputPath, outputPath || "sample/formatted-book.docx").catch((err) => {
    console.error(err);
    process.exit(1);
  });
}

module.exports = { build };
