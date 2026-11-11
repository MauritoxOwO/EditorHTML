// Start npm.cmd run dev:web, then node --test tests/paragraph-reflow-browser.cjs.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const test = require("node:test");
const { chromium } = require(process.env.PLAYWRIGHT_MODULE ||
  "C:/Users/Mauro/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright");
const url = process.env.EDITOR_URL || "http://127.0.0.1:5173/";
const channel = process.env.BROWSER_CHANNEL || "chrome";
const installed = channel === "msedge"
  ? "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
  : "C:/Program Files/Google/Chrome/Application/chrome.exe";

const settle = (page) => page.evaluate(() => new Promise((resolve) =>
  requestAnimationFrame(() => requestAnimationFrame(resolve))));

async function state(page) {
  return page.evaluate(() => {
    const inners = [...document.querySelectorAll(".hwe-page-inner")];
    const paragraphs = inners.flatMap((inner) => [...inner.querySelectorAll("p")]);
    const p21 = paragraphs.find((p) => p.textContent.includes("Parrafo 21."));
    return {
      text: inners.map((inner) => inner.textContent).join(""),
      blanks: paragraphs.filter((p) => !p.textContent.trim()).length,
      page: inners.indexOf(p21?.closest(".hwe-page-inner")),
      pageCount: inners.length,
      // Relative measurements exclude scrolling and account for document zoom.
      used: inners.map((inner) => {
        const rect = inner.getBoundingClientRect();
        return (inner.lastElementChild.getBoundingClientRect().bottom - rect.top) / rect.height;
      }),
    };
  });
}

test("párrafo 21: Enter, volver con Retroceso y seguir borrando no deja huecos de paginación", async (t) => {
  const browser = await chromium.launch({ headless: true,
    ...(fs.existsSync(installed) ? { executablePath: installed } : { channel }) });
  t.after(() => browser.close());
  const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto(url);
  await page.locator("#sample-button").click();
  await page.waitForFunction(() => document.querySelectorAll(".hwe-page").length === 3);
  await page.evaluate(() => document.fonts.ready);
  await settle(page);
  const baseline = await state(page);
  await page.evaluate(() => {
    const p = [...document.querySelectorAll("p")].find((p) => p.textContent.startsWith("Parrafo 21."));
    p.scrollIntoView({ block: "center" });
    document.querySelector(".hwe-workspace").focus();
    getSelection().setPosition(p.querySelector("strong").firstChild, 0);
  });

  let inserted = 0;
  // Actual margins/fonts determine the spare lines; eight Enters need not
  // cross a page boundary. Exercise the real boundary instead of a fixed count.
  while (inserted < 80 && (await state(page)).page === baseline.page) {
    await page.keyboard.press("Enter");
    await settle(page);
    inserted++;
  }
  const moved = await state(page);
  assert.ok(moved.page > baseline.page, "Enter must move paragraph 21 to the next page");
  assert.equal(moved.blanks, baseline.blanks + inserted);

  for (let i = 0; i < inserted; i++) {
    await page.keyboard.press("Backspace");
    await settle(page);
    assert.equal((await state(page)).blanks, baseline.blanks + inserted - 1 - i);
  }
  const restored = await state(page);
  assert.equal(restored.text, baseline.text);
  assert.equal(restored.page, baseline.page);
  assert.equal(restored.pageCount, baseline.pageCount);

  // The first Backspace joins paragraph 20 and 21. Subsequent keys delete text
  // through a previously split paragraph, the operation that exposed the gap.
  for (let i = 0; i < 12; i++) {
    await page.keyboard.press("Backspace");
    await settle(page);
    const current = await state(page);
    assert.equal(current.blanks, baseline.blanks);
    assert.equal(current.text.length, baseline.text.length - i);
    assert.equal(current.pageCount, baseline.pageCount);
    assert.equal(current.page, baseline.page, "deleting text must not push paragraph 21 forward");
    assert.ok(current.used[baseline.page] >= baseline.used[baseline.page] - 0.08,
      `page ${baseline.page + 1} unexpectedly became empty: ${current.used[baseline.page]}`);
  }
  assert.deepEqual(errors, []);
});

test("cada Enter delante de un párrafo añade una sola altura de línea", async (t) => {
  const browser = await chromium.launch({ headless: true,
    ...(fs.existsSync(installed) ? { executablePath: installed } : { channel }) });
  t.after(() => browser.close());
  const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
  await page.goto(url);
  await page.locator("#sample-button").click();
  await page.waitForFunction(() => document.querySelectorAll(".hwe-page").length === 3);
  await page.evaluate(() => document.fonts.ready);
  await settle(page);
  const position = () => page.evaluate(() => {
    const p = [...document.querySelectorAll("p")].find((p) => p.textContent.includes("Parrafo 26."));
    const inner = p.closest(".hwe-page-inner");
    const rect = inner.getBoundingClientRect();
    const scale = rect.width / parseFloat(getComputedStyle(inner).width);
    return { top: (p.getBoundingClientRect().top - rect.top) / scale,
      lineHeight: parseFloat(getComputedStyle(p).lineHeight) };
  });
  await page.evaluate(() => {
    const p = [...document.querySelectorAll("p")].find((p) => p.textContent.startsWith("Parrafo 26."));
    p.scrollIntoView({ block: "center" });
    document.querySelector(".hwe-workspace").focus();
    getSelection().setPosition(p.querySelector("strong").firstChild, 0);
  });
  let previous = await position();
  for (let i = 0; i < 4; i++) {
    await page.keyboard.press("Enter");
    await settle(page);
    const current = await position();
    const displacement = current.top - previous.top;
    assert.ok(Math.abs(displacement - current.lineHeight) < 1,
      `Enter ${i + 1}: desplazamiento ${displacement}px, línea ${current.lineHeight}px`);
    previous = current;
  }
});
