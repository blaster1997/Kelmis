// SVG -> PNG über Chromium. Aufruf: node render_svg.js in.svg out.png [scale=3] [--white]
const { chromium } = require("/opt/node22/lib/node_modules/playwright");
const fs = require("fs");
(async () => {
  const [, , inp, out, sc] = process.argv;
  const white = process.argv.includes("--white");
  const scale = sc && !sc.startsWith("--") ? +sc : 3;
  const svg = fs.readFileSync(inp, "utf8");
  const w = +(svg.match(/<svg[^>]*\swidth="([\d.]+)/) || [])[1], h = +(svg.match(/<svg[^>]*\sheight="([\d.]+)/) || [])[1];
  if (!w || !h) { console.error("svg braucht width/height in px"); process.exit(1); }
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: Math.ceil(w), height: Math.ceil(h) }, deviceScaleFactor: scale });
  await page.setContent(`<html><body style="margin:0;background:${white ? "#fff" : "transparent"}">${svg}</body></html>`);
  await page.waitForTimeout(150);
  await page.screenshot({ path: out, omitBackground: !white, clip: { x: 0, y: 0, width: w, height: h } });
  await browser.close();
  console.log("ok", out, w, "x", h);
})().catch(e => { console.error(e.message); process.exit(1); });
