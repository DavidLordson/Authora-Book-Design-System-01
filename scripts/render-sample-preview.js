#!/usr/bin/env node
/**
 * Renders the same content/tokens as build-sample-page.js to a PNG, using
 * the real embedded fonts, so the design can be reviewed as an image.
 * (LibreOffice's headless docx->pdf conversion is non-functional in this
 * sandbox, so Chromium + the exact CSS equivalent of the docx tokens is
 * used instead as the visual proof.)
 */
const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");
const { COLORS, FONTS, SIZES, SPACING, PAGE } = require("../src/design-system");

const DPI = 150;
const TWIP_TO_PX = DPI / 1440;
const HALFPT_TO_PX = DPI / 144; // 1pt = 1/72in; half-pt size -> pt -> px

const px = (twips) => (twips * TWIP_TO_PX).toFixed(2) + "px";
const fpx = (halfPoints) => (halfPoints * HALFPT_TO_PX).toFixed(2) + "px";

const fontsDir = path.resolve(__dirname, "../assets/fonts");
const fontFace = (family, file, weight = "normal", style = "normal") => `
@font-face {
  font-family: "${family}";
  src: url("file://${fontsDir}/${file}");
  font-weight: ${weight};
  font-style: ${style};
}`;

const css = `
${fontFace("Palatino Linotype", "palatinolinotype_roman.ttf")}
${fontFace("Palatino Linotype", "palatinolinotype_bold.ttf", "bold")}
${fontFace("Palatino Linotype", "palatinolinotype_italic.ttf", "normal", "italic")}
${fontFace("Palatino Linotype", "palatinolinotype_bolditalic.ttf", "bold", "italic")}
${fontFace("Great Vibes", "GreatVibes-Regular.ttf")}

* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; background: #808080; }
.page {
  width: ${px(PAGE.width)};
  height: ${px(PAGE.height)};
  background: #${COLORS.background};
  color: #${COLORS.body};
  font-family: "${FONTS.serif}", ${FONTS.serifFallback};
  font-size: ${fpx(SIZES.bodyText)};
  line-height: 1.35;
  padding: ${px(PAGE.marginTop)} ${px(PAGE.marginOutside)} ${px(PAGE.marginBottom)} ${px(PAGE.marginInside)};
  position: relative;
}
.chapter-label {
  text-align: center;
  font-style: italic;
  font-size: ${fpx(SIZES.chapterLabel)};
  margin: ${px(SPACING.chapterOpenerBefore)} 0 ${px(SPACING.chapterLabelAfter)} 0;
}
.chapter-title {
  text-align: center;
  font-family: "Great Vibes", ${FONTS.scriptFallback};
  font-size: ${fpx(SIZES.chapterTitle)};
  font-weight: normal;
  margin: 0 0 ${px(SPACING.chapterTitleAfter)} 0;
}
.epigraph {
  text-align: center;
  font-style: italic;
  margin: ${px(SPACING.epigraphBefore)} 0 ${px(SPACING.epigraphAfter)} 0;
}
.epigraph-citation {
  text-align: center;
  font-style: italic;
  margin: 0 0 ${px(SPACING.epigraphCitationAfter)} 0;
}
p.body { margin: 0 0 ${px(SPACING.bodyParagraphAfter)} 0; text-align: left; }
p.drop-cap::first-letter {
  float: left;
  font-size: 3.4em;
  line-height: 0.78;
  padding-right: 0.06em;
  padding-top: 0.02em;
}
.block-quote {
  font-style: italic;
  margin: ${px(SPACING.blockQuoteBefore)} ${px(360)} ${px(SPACING.blockQuoteAfter)} ${px(360)};
}
.citation {
  font-style: italic;
  margin: 0 0 ${px(SPACING.bodyParagraphAfter)} ${px(360)};
}
.section-heading {
  font-weight: bold;
  margin: ${px(SPACING.sectionHeadingBefore)} 0 ${px(SPACING.sectionHeadingAfter)} 0;
}
.section-heading .cap { font-size: ${fpx(SIZES.sectionHeadingCap)}; }
.footer {
  position: absolute;
  bottom: ${px(PAGE.footerDistance / 2)};
  left: 0; right: 0;
  text-align: center;
  font-size: ${fpx(SIZES.footer)};
}
`;

function sectionHeadingHTML(text) {
  return text
    .toUpperCase()
    .split(" ")
    .map((w) => `<span class="cap">${w[0]}</span>${w.slice(1)}`)
    .join(" ");
}

const bodyParas = [
  "We like to believe we are strong enough to keep close company with anything and remain unaffected. Paul says do not be deceived. The word he uses for communications is not only speech; it is fellowship, exchange, the daily traffic of two lives. What is exchanged in that traffic either corrupts or builds. Nobody stays neutral.",
];

const html = `<!doctype html>
<html><head><meta charset="utf-8"><style>${css}</style></head>
<body>
  <div class="page">
    <div class="chapter-label">Chapter One</div>
    <div class="chapter-title">Who Are Your Friends?</div>
    <div class="epigraph">&ldquo;Be not deceived: evil communications corrupt good manners.&rdquo;</div>
    <div class="epigraph-citation">1 Corinthians 15:33, KJV</div>
    <p class="body drop-cap">You come under curse or under blessing depending on who you associate with. That single sentence, if you receive it, will reorder your life.</p>
    ${bodyParas.map((t) => `<p class="body">${t}</p>`).join("\n    ")}
    <div class="footer">12</div>
  </div>
</body></html>`;

async function run() {
  const outDir = path.resolve(__dirname, "../sample");
  fs.mkdirSync(outDir, { recursive: true });
  const htmlPath = path.join(outDir, "chapter-one-preview.html");
  fs.writeFileSync(htmlPath, html);

  const browser = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium-1194/chrome-linux/chrome" });
  const page = await browser.newPage({ viewport: { width: PAGE.width * TWIP_TO_PX, height: PAGE.height * TWIP_TO_PX } });
  await page.goto("file://" + htmlPath);
  await page.waitForTimeout(200);
  const pngPath = path.join(outDir, "chapter-one-preview.png");
  await page.locator(".page").screenshot({ path: pngPath });
  await browser.close();
  console.log("Written to", pngPath);
}

run().catch((err) => { console.error(err); process.exit(1); });
