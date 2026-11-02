// Start npm.cmd run dev:web, then node --test tests/page-bottom-margin-browser.cjs.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const test = require("node:test");
const projectPath = require("node:path").resolve(__dirname, "..").replace(/\\/g, "/");
const { chromium } = require(process.env.PLAYWRIGHT_MODULE ||
  "C:/Users/Mauro/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright");
const url = process.env.EDITOR_URL || "http://127.0.0.1:5173/";
const executablePath = process.env.BROWSER_PATH || "C:/Program Files/Google/Chrome/Application/chrome.exe";
const paragraphs = Array.from({ length: 55 }, (_, i) =>
  `<p>Parrafo ${i + 1}. ${"Texto para comprobar el margen inferior y la continuidad del documento. ".repeat(4)}</p>`).join("");
const table = '<table class="hwe-word-table"><tbody>' + Array.from({ length: 100 }, (_, i) =>
  `<tr><td><p>Fila ${i + 1}</p></td><td><p>${"Contenido de tabla para comprobar su reparto entre hojas. ".repeat(3)}</p></td></tr>`).join("") + "</tbody></table>";

async function measure(page) {
  return page.evaluate(() => [...document.querySelectorAll(".hwe-page-inner")].map(inner => {
    const css = getComputedStyle(inner);
    const rect = inner.getBoundingClientRect();
    const scale = rect.height / parseFloat(css.height);
    const bottom = Math.max(...[...inner.children].map(child => child.getBoundingClientRect().bottom));
    return { freeMm: (rect.bottom - bottom) / scale / (96 / 25.4),
      paddingMm: parseFloat(css.paddingBottom) / (96 / 25.4) };
  }));
}

test("texto y tablas respetan el margen inferior al paginar y al imprimir", async (t) => {
  const browser = await chromium.launch({ headless: true,
    ...(fs.existsSync(executablePath) ? { executablePath } : { channel: "chrome" }) });
  t.after(() => browser.close());
  for (const zoom of [0.5, 1, 1.18, 2.5]) {
    for (const [kind, html] of [["texto", paragraphs], ["tabla", table]]) {
      await t.test(`${kind}, zoom ${zoom}`, async () => {
        const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
        try {
          const errors = [];
          page.on("pageerror", error => errors.push(error.message));
          await page.goto(url);
          await page.waitForSelector(".hwe-page-inner");
          await page.waitForFunction(() => !document.querySelector(".pcf-html-editor-root").inert);
          await page.addStyleTag({ content: `.pcf-html-editor-root { --hwe-page-view-zoom:${zoom} !important; }` });
          await page.locator("#html-file").setInputFiles({ name: "margen.html", mimeType: "text/html", buffer: Buffer.from(html) });
          await page.waitForFunction(() => document.querySelector("#html-file").value === "" &&
            document.querySelectorAll(".hwe-page-inner").length > 1);
          await page.evaluate(() => document.fonts.ready);
          await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
          const expected = html.replace(/<[^>]+>/g, "").replace(/\s+/g, "");
          const actual = await page.locator(".hwe-page-inner").allTextContents();
          assert.equal(actual.join("").replace(/\s+/g, ""), expected, "conserva todo el texto en orden");
          if (kind === "tabla") assert.equal(await page.locator(".hwe-page-inner tr").count(), 100);
          for (const media of ["screen", "print"]) {
            await page.emulateMedia({ media });
            const measurements = await measure(page);
            assert.ok(measurements.length > 1);
            for (const [index, item] of measurements.entries()) {
              assert.ok(Math.abs(item.paddingMm - 20) < 0.01);
              assert.ok(item.freeMm >= 19.7, `${media}, hoja ${index + 1}: quedan ${item.freeMm} mm`);
            }
          }
          const printHtml = await page.evaluate(async projectPath => {
            const { DocumentSerializer } = await import(`/@fs/${projectPath}/EditorHTML/services/DocumentSerializer.ts`);
            const { DEFAULT_PAGE_SETUP } = await import(`/@fs/${projectPath}/EditorHTML/pagination/PageGeometry.ts`);
            const workspace = document.querySelector(".hwe-workspace");
            return new DocumentSerializer().collectPdfHtml(workspace,
              [...workspace.querySelectorAll(".hwe-page")], DEFAULT_PAGE_SETUP, "");
          }, projectPath);
          const printPage = await browser.newPage();
          try {
            await printPage.emulateMedia({ media: "print" });
            await printPage.setContent(printHtml);
            await printPage.evaluate(() => document.fonts.ready);
            for (const [index, item] of (await measure(printPage)).entries()) {
              assert.ok(item.freeMm >= 19.7, `HTML para PDF, hoja ${index + 1}: quedan ${item.freeMm} mm`);
            }
            const exported = await printPage.locator(".hwe-page-inner").allTextContents();
            assert.equal(exported.join("").replace(/\s+/g, ""), expected);
          } finally {
            await printPage.close();
          }
          assert.deepEqual(errors, []);
        } finally {
          await page.close();
        }
      });
    }
  }
});
