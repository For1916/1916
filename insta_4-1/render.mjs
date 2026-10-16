import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import path from "node:path";

const require = createRequire(import.meta.url);
const { chromium } = require("playwright");

const root = path.dirname(fileURLToPath(import.meta.url));

const browser = await chromium.launch({
  channel: "chrome",
  args: ["--no-sandbox", "--disable-dev-shm-usage", "--font-render-hinting=none"],
});

async function render(option) {
  const page = await browser.newPage({
    viewport: { width: 1080, height: 1080 },
    deviceScaleFactor: 1,
  });
  const htmlPath = path.join(root, option, "index.html");
  await page.goto("file://" + htmlPath);
  await page.evaluate(async () => {
    await document.fonts.ready;
  });
  const fontInfo = await page.evaluate(() => {
    const faces = [...document.fonts].map((f) => `${f.family} ${f.weight} ${f.status}`);
    return faces;
  });
  const loaded = fontInfo.filter((f) => f.includes("Pretendard") && f.endsWith("loaded"));
  if (loaded.length < 5) {
    throw new Error(`Pretendard did not load for ${option}: ${fontInfo.join(" | ")}`);
  }

  const report = await page.evaluate(() => {
    const problems = [];
    for (const card of document.querySelectorAll(".card")) {
      const c = card.getBoundingClientRect();
      if (Math.abs(c.width - 1080) > 1 || Math.abs(c.height - 1080) > 1) {
        problems.push({ id: card.id, issue: "size", w: c.width, h: c.height });
      }
      for (const el of card.querySelectorAll("*")) {
        const r = el.getBoundingClientRect();
        if (r.width < 1 || r.height < 1) continue;
        const tol = 1.5;
        if (r.left < c.left - tol || r.right > c.right + tol || r.top < c.top - tol || r.bottom > c.bottom + tol) {
          problems.push({
            id: card.id,
            issue: "outside",
            tag: el.tagName,
            text: (el.innerText || "").replace(/\s+/g, " ").slice(0, 60),
            rect: [Math.round(r.left), Math.round(r.top), Math.round(r.right), Math.round(r.bottom)],
          });
        }
        const style = getComputedStyle(el);
        if ((style.overflowX !== "visible" || style.overflowY !== "visible") && (el.scrollWidth > el.clientWidth + 2 || el.scrollHeight > el.clientHeight + 2)) {
          problems.push({
            id: card.id,
            issue: "clipped",
            tag: el.tagName,
            text: (el.innerText || "").replace(/\s+/g, " ").slice(0, 60),
          });
        }
      }
    }
    return problems;
  });

  const cards = await page.$$(".card");
  for (let i = 0; i < cards.length; i++) {
    const name = String(i + 1).padStart(2, "0") + ".png";
    await cards[i].screenshot({ path: path.join(root, option, name) });
  }
  await page.close();
  return report;
}

const all = [];
for (const option of ["option1", "option2"]) {
  const report = await render(option);
  all.push({ option, report });
  console.log(option, report.length ? JSON.stringify(report, null, 2) : "clean");
}
await browser.close();
if (all.some((item) => item.report.length)) process.exit(1);
