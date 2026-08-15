// Authora Book Design System — canonical style tokens
// Extracted directly from the reference sample "The Wonders of Prayer and
// Fasting" (classic black-and-white literary non-fiction/devotional style).
// Measured from the PDF itself (page geometry, embedded font list, span
// sizes) — not eyeballed.

// The interior of the reference book carries no color at all: pure black
// text on white. Color only appears on the cover (lavender/periwinkle
// background, muted blue accent).
const COLORS = {
  body: "000000",
  primary: "000000",
  muted: "000000",       // footer/page numbers are plain black, not gray
  background: "FFFFFF",
  coverBackground: "D6D2EC", // lavender/periwinkle, sampled from reference cover
  coverAccent: "8FA8D6",     // muted cornflower blue (photo/watercolor accent)
};

const FONTS = {
  serif: "Palatino Linotype",       // body text, scripture epigraphs, section headings
  serifFallback: "Georgia",
  script: "Great Vibes",             // chapter titles (calligraphic display)
  scriptFallback: "Brush Script MT",
};

// All sizes in half-points (1 pt = 2 half-points)
const SIZES = {
  bodyText: 25,             // 12.5 pt — matches the reference's Palatino Roman body
  chapterLabel: 28,          // 14 pt   ("Chapter One", italic)
  chapterTitle: 52,          // 26 pt   (Great Vibes script)
  chapterSubtitle: 26,       // 13 pt   (bold italic description line, when present)
  epigraph: 25,              // 12.5 pt (scripture quote, italic)
  epigraphCitation: 25,      // 12.5 pt (citation line under a scripture quote)
  sectionHeading: 25,        // 12.5 pt (faux small-caps: see makeSectionHeading)
  sectionHeadingCap: 30,     // 15 pt   (the enlarged first letter of a small-caps heading)
  footer: 25,                // 12.5 pt — plain Palatino, no size reduction in the reference
  footnote: 21,               // 10.5 pt
  caption: 20,                // 10 pt
};

// Spacing in DXA (twentieths of a point; 1440 = 1 inch)
const SPACING = {
  bodyParagraphAfter: 160,       // 8 pt between block paragraphs (no first-line indent)
  bulletIndent: 400,
  bulletHanging: 260,
  bulletAfter: 80,
  chapterLabelAfter: 120,        // 6 pt, "Chapter One" -> title
  chapterTitleAfter: 200,        // 10 pt, title -> subtitle/epigraph
  chapterSubtitleAfter: 240,     // 12 pt
  epigraphBefore: 200,
  epigraphAfter: 60,
  epigraphCitationAfter: 320,    // 16 pt, citation -> first body paragraph
  sectionHeadingBefore: 300,     // 15 pt
  sectionHeadingAfter: 150,      // 7.5 pt
  blockQuoteBefore: 200,
  blockQuoteAfter: 200,
  sectionBreakBefore: 400,
  sectionBreakAfter: 400,
  dropCapLines: 3,                // number of body lines the drop cap spans
};

// 6 x 9 in trim (432 x 648 pt), matching the reference sample exactly
const PAGE = {
  width: 8640,   // 6 in
  height: 12960, // 9 in
  marginTop: 1152,     // ~0.8 in
  marginBottom: 1080,  // 0.75 in
  marginInside: 1080,  // 0.75 in
  marginOutside: 1080, // 0.75 in
  headerDistance: 576,
  footerDistance: 576,
};

module.exports = { COLORS, FONTS, SIZES, SPACING, PAGE };
