#!/usr/bin/env node
/**
 * One-off sample generator — proves out the Authora design tokens against
 * real manuscript content (Chapter One of "Blessed by Association") before
 * any general-purpose parser/CLI gets built.
 *
 * Unlike the Mrs Violet system, the reference book ("The Wonders of Prayer
 * and Fasting") does NOT give each chapter its own vertically-centered
 * opener spread — the chapter label, title, scripture epigraph, and body
 * text all flow down a single top-anchored page, pushed down by a fixed
 * space-before on the chapter title paragraph.
 */
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, Footer, PageNumber,
  BorderStyle, HorizontalPositionAlign, VerticalPositionAlign,
  FrameAnchorType, HeightRule, DropCapType,
} = require("docx");

const { COLORS, FONTS, SIZES, SPACING, PAGE } = require("../src/design-system");

function bodyStyle() {
  return { font: FONTS.serif, size: SIZES.bodyText, color: COLORS.body };
}

function makeChapterLabel(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 2030, after: SPACING.chapterLabelAfter },
    children: [
      new TextRun({ text, font: FONTS.serif, size: SIZES.chapterLabel, color: COLORS.primary, italics: true }),
    ],
  });
}

function makeChapterTitle(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: SPACING.chapterTitleAfter },
    children: [
      new TextRun({ text, font: FONTS.script, size: SIZES.chapterTitle, color: COLORS.primary }),
    ],
  });
}

function makeEpigraph(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: SPACING.epigraphBefore, after: SPACING.epigraphAfter },
    children: [
      new TextRun({ text, font: FONTS.serif, size: SIZES.epigraph, color: COLORS.primary, italics: true }),
    ],
  });
}

function makeEpigraphCitation(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: SPACING.epigraphCitationAfter },
    children: [
      new TextRun({ text, font: FONTS.serif, size: SIZES.epigraphCitation, color: COLORS.primary, italics: true }),
    ],
  });
}

// Faux small-caps section heading, matching the reference exactly: first
// letter of each significant word rendered larger, the rest of the word in
// smaller true caps — a manual Word trick, not a real small-caps font
// feature.
function makeSectionHeading(text) {
  const words = text.toUpperCase().split(" ");
  const runs = [];
  words.forEach((word, wi) => {
    if (wi > 0) runs.push(new TextRun({ text: " ", font: FONTS.serif, size: SIZES.sectionHeading, bold: true }));
    runs.push(new TextRun({ text: word[0], font: FONTS.serif, size: SIZES.sectionHeadingCap, bold: true }));
    if (word.length > 1) {
      runs.push(new TextRun({ text: word.slice(1), font: FONTS.serif, size: SIZES.sectionHeading, bold: true }));
    }
  });
  return new Paragraph({
    spacing: { before: SPACING.sectionHeadingBefore, after: SPACING.sectionHeadingAfter },
    children: runs,
  });
}

function makeBodyParagraph(text) {
  return new Paragraph({ spacing: { after: SPACING.bodyParagraphAfter }, children: [new TextRun({ ...bodyStyle(), text })] });
}

function makeDropCapParagraph(text) {
  const dropLetter = text.charAt(0);
  const rest = text.slice(1);
  const dropCapPara = new Paragraph({
    frame: {
      type: "alignment",
      alignment: { x: HorizontalPositionAlign.LEFT, y: VerticalPositionAlign.TOP },
      anchor: { horizontal: FrameAnchorType.TEXT, vertical: FrameAnchorType.TEXT },
      width: 800, height: 800, rule: HeightRule.AUTO,
      dropCap: DropCapType.DROP, lines: SPACING.dropCapLines,
    },
    children: [new TextRun({ text: dropLetter, font: FONTS.serif, color: COLORS.primary })],
  });
  const restPara = new Paragraph({ spacing: { after: SPACING.bodyParagraphAfter }, children: [new TextRun({ ...bodyStyle(), text: rest })] });
  return [dropCapPara, restPara];
}

function makeBlockQuote(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    indent: { left: 360, right: 360 },
    spacing: { before: SPACING.blockQuoteBefore, after: SPACING.blockQuoteAfter },
    children: [new TextRun({ text, font: FONTS.serif, size: SIZES.epigraph, color: COLORS.primary, italics: true })],
  });
}

function makeCitation(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    indent: { left: 360 },
    spacing: { after: SPACING.bodyParagraphAfter },
    children: [new TextRun({ text, font: FONTS.serif, size: SIZES.epigraphCitation, color: COLORS.primary, italics: true })],
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

async function build() {
  const children = [
    makeChapterLabel("Chapter One"),
    makeChapterTitle("Who Are Your Friends?"),
    makeEpigraph("“Be not deceived: evil communications corrupt good manners.”"),
    makeEpigraphCitation("1 Corinthians 15:33, KJV"),
    ...makeDropCapParagraph(
      "You come under curse or under blessing depending on who you associate with. That single sentence, if you receive it, will reorder your life."
    ),
    makeBodyParagraph(
      "We like to believe we are strong enough to keep close company with anything and remain unaffected. Paul says do not be deceived. The word he uses for communications is not only speech; it is fellowship, exchange, the daily traffic of two lives. What is exchanged in that traffic either corrupts or builds. Nobody stays neutral."
    ),
    makeBodyParagraph(
      "Look at how the Bible frames it in Proverbs. It does not say the companion of fools will make a mistake. It says he shall be destroyed. That is the language of collapse, of a life pulled down. Meanwhile the one who walks with wise men shall be wise — not shall learn wisdom, shall become wise. Wisdom becomes his nature by proximity."
    ),
    makeBlockQuote(
      "“And Joshua the son of Nun was full of the spirit of wisdom; for Moses had laid his hands upon him: and the children of Israel hearkened unto him, and did as the LORD commanded Moses.”"
    ),
    makeCitation("Deuteronomy 34:9, KJV"),
    makeBodyParagraph(
      "Joshua did not attend a school of wisdom. Joshua carried Moses' bag. He stood at the door of the tent when the cloud came down. He was the young man who did not leave when everybody else went back to the camp. Years of nearness produced a man full of the spirit of wisdom, and a whole nation followed him."
    ),
    makeSectionHeading("Take These Steps"),
    makeBodyParagraph(
      "Write the names of the five people who have the most access to your ear."
    ),
  ];

  const doc = new Document({
    creator: "Authora Book Design System",
    description: "Design token sample — Chapter One opener",
    styles: { default: { document: { run: { font: FONTS.serif, size: SIZES.bodyText, color: COLORS.body } } } },
    sections: [
      {
        properties: {
          page: {
            size: { width: PAGE.width, height: PAGE.height },
            margin: {
              top: PAGE.marginTop, bottom: PAGE.marginBottom,
              left: PAGE.marginInside, right: PAGE.marginOutside,
              header: PAGE.headerDistance, footer: PAGE.footerDistance,
            },
          },
        },
        footers: { default: makeFooter() },
        children,
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync("sample/chapter-one-sample.docx", buffer);
  console.log("Written to sample/chapter-one-sample.docx");
}

build().catch((err) => { console.error(err); process.exit(1); });
