const path = require("path");
const { chromium } = require("playwright-core");

const cwd = __dirname;
const htmlPath = path.join(cwd, "enjad-hajj-2026-final-report.html");
const pdfPath = path.join(cwd, "enjad-hajj-1447-final-report.pdf");
const edgePath = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";

(async () => {
  const browser = await chromium.launch({
    executablePath: edgePath,
    headless: true,
  });
  const page = await browser.newPage({ viewport: { width: 1240, height: 1754 } });
  await page.goto(`file:///${htmlPath.replace(/\\/g, "/")}`, { waitUntil: "networkidle" });
  await page.emulateMedia({ media: "print" });
  await page.pdf({
    path: pdfPath,
    format: "A4",
    printBackground: true,
    preferCSSPageSize: true,
    displayHeaderFooter: false,
  });
  await browser.close();
  console.log(pdfPath);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
