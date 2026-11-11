// Uses the real TypeScript modules and CSS in headless Chrome; no local server is required.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const ts = require("typescript");
const { chromium } = require(process.env.PLAYWRIGHT_MODULE ||
  "C:/Users/Mauro/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright");
const project = path.resolve(__dirname, "..");

function buildRuntime() {
  const modules = new Map();
  function collect(filename) {
    filename = path.resolve(filename).replace(/\\/g, "/");
    if (modules.has(filename)) return filename;
    const output = ts.transpileModule(fs.readFileSync(filename, "utf8"), {
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 }, fileName: filename,
    }).outputText;
    const entry = { output, deps: {} };
    modules.set(filename, entry);
    for (const [, dependency] of output.matchAll(/require\("([^"]+)"\)/g)) {
      if (!dependency.startsWith(".")) throw new Error(`Unexpected dependency: ${dependency}`);
      entry.deps[dependency] = collect(path.resolve(path.dirname(filename), dependency + ".ts"));
    }
    return filename;
  }
  const exports = Object.fromEntries([
    ["Paginator", "EditorHTML/pagination/Paginator.ts"],
    ["EditorLayoutService", "EditorHTML/services/EditorLayoutService.ts"],
    ["ParagraphStyleManager", "EditorHTML/services/ParagraphStyleManager.ts"],
    ["DocumentSerializer", "EditorHTML/services/DocumentSerializer.ts"],
    ["getPageContentBounds", "EditorHTML/pagination/PageContentGeometry.ts"],
    ["DEFAULT_PAGE_SETUP", "EditorHTML/pagination/PageGeometry.ts"],
  ].map(([name, file]) => [name, collect(path.join(project, file))]));
  return `(function(){const factories={${[...modules].map(([id, entry]) =>
    `${JSON.stringify(id)}:(module,exports,require)=>{${entry.output}\n}`).join(",")}};
    const deps=${JSON.stringify(Object.fromEntries([...modules].map(([id, entry]) => [id, entry.deps])))};
    const cache={};function load(id){if(cache[id])return cache[id].exports;const m={exports:{}};
    cache[id]=m;factories[id](m,m.exports,n=>load(deps[id][n]));return m.exports;}
    window.titleRuntime={${Object.entries(exports).map(([name, id]) => `${name}:load(${JSON.stringify(id)}).${name}`).join(",")}};})();`;
}

test("títulos y párrafos permanecen unidos al paginar, editar y exportar", async t => {
  const runtime = buildRuntime();
  const css = fs.readFileSync(path.join(project, "EditorHTML/css/editor.css"), "utf8");
  const browser = await chromium.launch({ headless: true,
    executablePath: process.env.BROWSER_PATH || "C:/Program Files/Google/Chrome/Application/chrome.exe" });
  t.after(() => browser.close());
  for (const zoom of [0.5, 1, 1.18, 2.5]) {
    await t.test(`unión y redistribución con zoom ${zoom}`, async () => {
      const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
      try {
        const errors = [];
        page.on("pageerror", error => errors.push(error.message));
        await page.setContent(`<style>${css}</style><main class="pcf-html-editor-root" style="height:1000px;--hwe-page-view-zoom:${zoom}"><div class="hwe-workspace"></div></main>`);
        await page.addScriptTag({ content: runtime });
        const result = await page.evaluate(async () => {
          await document.fonts.ready;
          const { Paginator, EditorLayoutService, ParagraphStyleManager, DocumentSerializer,
            getPageContentBounds, DEFAULT_PAGE_SETUP } = window.titleRuntime;
          const root = document.querySelector("main");
          const workspace = root.firstElementChild;
          const layout = new EditorLayoutService();
          const styles = new ParagraphStyleManager(() => root, { setParagraphStyles() {}, setFontFamilies() {} });
          const catalog = [
            ...["art-TIT", "art-tit", "art_tit", "cap-tit"].map(className => ({
              label: className, className, keepWithNext: true,
              cssText: "font-size:12px;line-height:20px;margin:0",
            })),
            { label: "Cuerpo", className: "cuerpo", cssText: "font-size:12px;line-height:20px;margin:0" },
          ];
          styles.setStyles(catalog);
          let pages = [];
          const factory = () => {
            const p = document.createElement("section");p.className = "hwe-page";
            p.innerHTML = '<div class="hwe-page-inner"></div>';
            return p;
          };
          const paginator = new Paginator(factory, next => {
            pages = next;
            for (const p of next) if (p.parentElement !== workspace) workspace.append(p);
            pages.forEach(p => layout.applyOfficialTableWidths(p));
          }, (p, after) => after.after(p));
          paginator.setKeepWithNextStyles(catalog.filter(style => style.keepWithNext).map(style => style.className));
          function render(html, available = 30) {
            workspace.innerHTML = "";
            const first = factory();workspace.append(first);
            const inner = first.firstElementChild;
            const height = getPageContentBounds(inner).contentHeight;
            inner.innerHTML = `<p id="previous" style="height:${height - available}px;margin:0;line-height:20px">Previo</p>${html}`;
            inner.querySelectorAll("p").forEach(p => {
              if (p.id !== "previous") p.style.cssText += ";margin:0;font-size:12px;line-height:20px";
            });
            layout.applyOfficialTableWidths(first);
            const expected = inner.textContent.replace(/\s+/g, "");
            paginator.setPages([first]);paginator.repaginateAll();
            return expected;
          }
          const pageIndex = selector => pages.findIndex(p => p.querySelector(selector));
          const text = () => pages.map(p => p.querySelector(".hwe-page-inner").textContent).join("").replace(/\s+/g, "");
          const fits = () => pages.every(p => {
            const inner = p.firstElementChild;const bounds = getPageContentBounds(inner);
            return [...inner.children].every(child => child.getBoundingClientRect().bottom <= bounds.bottom + bounds.scale);
          });
          const short = '<p id="title" data-hwe-paragraph-style="art-TIT">Título</p><p id="body">' + "Contenido del párrafo. ".repeat(10) + "</p>";
          const expected = render(short);
          const together = { title: pageIndex("#title"), body: pageIndex("#body"), fits: fits(), text: text() === expected };
          styles.applyToBlock(workspace.querySelector("#title"), "cuerpo");
          paginator.repaginateAll();
          const removed = { groups: workspace.querySelectorAll("[data-hwe-title-group]").length, title: pageIndex("#title"), text: text() === expected };
          styles.applyToBlock(workspace.querySelector("#title"), "art-TIT");
          paginator.repaginateAll();
          const reapplied = pageIndex("#title") === pageIndex("#body") && pageIndex("#title") > 0;
          workspace.querySelector("#previous").style.height = "10px";
          paginator.repaginateAll();
          const compacted = pageIndex("#title") === 0 && pageIndex("#body") === 0;
          const chainExpected = render('<p class="cap-tit">Capítulo</p><p id="title" class="art_tit">Artículo</p><p id="body">' + "Cuerpo del artículo. ".repeat(8) + "</p>");
          const chain = { count: workspace.querySelector("[data-hwe-title-group]")?.querySelectorAll("p").length,
            title: pageIndex("#title"), body: pageIndex("#body"), intact: text() === chainExpected, fits: fits() };
          render('<p id="title" class="art-tit">Título</p><div data-hwe-manual-page-break="true"></div><p id="body">Cuerpo</p>', 200);
          const manual = { groups: workspace.querySelectorAll("[data-hwe-title-group]").length, separate: pageIndex("#body") > pageIndex("#title") };
          render('<p id="title" class="art-tit">Título</p><p data-hwe-user-blank="true"><br></p><p id="body">Cuerpo</p>', 200);
          const blank = workspace.querySelectorAll("[data-hwe-title-group]").length === 0;
          render('<p class="art-tit">Título</p><p id="separator"><br></p><p>Cuerpo</p>', 200);
          const importedBlank = workspace.querySelectorAll("[data-hwe-title-group]").length === 0 &&
            workspace.querySelector("#separator").getAttribute("data-hwe-user-blank") === "true";
          const longExpected = render('<p id="title" class="art-tit">Título</p><p id="body">' + "Contenido extenso para dividir el párrafo sin dejar solo el título. ".repeat(160) + "</p>");
          const firstBody = workspace.querySelector("#body");
          const range = document.createRange();range.selectNodeContents(firstBody);
          const lineTops = [];
          for (const rect of range.getClientRects()) if (rect.width > 0 && !lineTops.some(top => Math.abs(top - rect.top) < 1)) lineTops.push(rect.top);
          const oversized = { pages: pages.length, title: pageIndex("#title"), body: pageIndex("#body"), lines: lineTops.length, fits: fits(), intact: text() === longExpected };
          paginator.repaginateAll();
          oversized.stable = text() === longExpected && fits();
          const serializer = new DocumentSerializer();
          const dirtyBefore = serializer.normalizeHtmlForDirtyCheck(serializer.collectHtml(workspace, pages, DEFAULT_PAGE_SETUP, ""));
          paginator.repaginateAll();
          const saved = serializer.collectHtml(workspace, pages, DEFAULT_PAGE_SETUP, "");
          const printed = serializer.collectPdfHtml(workspace, pages, DEFAULT_PAGE_SETUP, "");
          return { together, removed, reapplied, compacted, chain, manual, blank, importedBlank, oversized,
            dirtyStable: dirtyBefore === serializer.normalizeHtmlForDirtyCheck(saved),
            cleanSaved: !saved.includes("data-hwe-title-group"), cleanPrint: !printed.includes("data-hwe-title-group"), printed };
        });
        assert.equal(result.together.title, result.together.body);
        assert.ok(result.together.title > 0);
        assert.ok(result.together.fits && result.together.text);
        assert.equal(result.removed.groups, 0);
        assert.equal(result.removed.title, 0);
        assert.ok(result.removed.text && result.reapplied && result.compacted);
        assert.equal(result.chain.count, 3);
        assert.equal(result.chain.title, result.chain.body);
        assert.ok(result.chain.title > 0 && result.chain.intact && result.chain.fits);
        assert.equal(result.manual.groups, 0);
        assert.ok(result.manual.separate && result.blank && result.importedBlank);
        assert.ok(result.oversized.pages > 2);
        assert.equal(result.oversized.title, result.oversized.body);
        assert.ok(result.oversized.lines >= 2);
        assert.ok(result.oversized.fits && result.oversized.intact && result.oversized.stable);
        assert.ok(result.dirtyStable && result.cleanSaved && result.cleanPrint);
        const printPage = await browser.newPage();
        try {
          await printPage.emulateMedia({ media: "print" });
          await printPage.setContent(result.printed);
          const margins = await printPage.evaluate(() => [...document.querySelectorAll(".hwe-page-inner")].map(inner => {
            const bounds = inner.getBoundingClientRect();
            return (bounds.bottom - Math.max(...[...inner.children].map(child => child.getBoundingClientRect().bottom))) / (96 / 25.4);
          }));
          assert.ok(margins.length > 1);
          assert.ok(margins.every(mm => mm >= 19.7));
        } finally { await printPage.close(); }
        assert.deepEqual(errors, []);
      } finally { await page.close(); }
    });
  }
});
