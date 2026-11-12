// Ejecuta el componente y el CSS reales en Chrome, sin servidor ni Dataverse.
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
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
      fileName: filename,
    }).outputText;
    const entry = { output, deps: {} };
    modules.set(filename, entry);
    for (const [, dependency] of output.matchAll(/require\("([^"]+)"\)/g)) {
      if (!dependency.startsWith(".")) throw new Error(`Unexpected dependency: ${dependency}`);
      entry.deps[dependency] = collect(path.resolve(path.dirname(filename), dependency + ".ts"));
    }
    return filename;
  }
  const entry = collect(path.join(project, "EditorHTML/app/EditorComponent.ts"));
  return `(function(){const factories={${[...modules].map(([id, value]) =>
    `${JSON.stringify(id)}:(module,exports,require)=>{${value.output}\n}`).join(",")}};
    const deps=${JSON.stringify(Object.fromEntries([...modules].map(([id, value]) => [id, value.deps])))};
    const cache={};function load(id){if(cache[id])return cache[id].exports;
    const module={exports:{}};cache[id]=module;factories[id](module,module.exports,n=>load(deps[id][n]));return module.exports;}
    window.EditorComponent=load(${JSON.stringify(entry)}).EditorComponent;})();`;
}

let browser;
let runtime;
test.before(async () => {
  runtime = buildRuntime();
  browser = await chromium.launch({ headless: true,
    executablePath: process.env.BROWSER_PATH || "C:/Program Files/Google/Chrome/Application/chrome.exe" });
});
test.after(async () => browser?.close());

async function open(t, hidden = false) {
  const page = await browser.newPage({ viewport: { width: 1100, height: 850 } });
  t.after(() => page.close());
  const errors = [];
  page.on("pageerror", error => errors.push(error.message));
  t.after(() => assert.deepEqual(errors, [], "sin errores del navegador"));
  await page.setContent(`<style>html,body{margin:0;height:100%}#surface{height:100%;${hidden ? "display:none;" : ""}}#host{height:100%}</style><div id="surface"><div id="host"></div></div>`);
  await page.addStyleTag({ content: fs.readFileSync(path.join(project, "EditorHTML/css/editor.css"), "utf8") });
  await page.addScriptTag({ content: runtime });
  return page;
}

async function assertBlocked(page) {
  await page.locator(".hwe-loading-overlay").waitFor({ state: "visible" });
  const state = await page.evaluate(() => {
    const root = document.querySelector(".pcf-html-editor-root");
    const button = root.querySelector("button");
    window.blockedClicks = 0;
    button.addEventListener("click", () => window.blockedClicks++);
    button.focus();
    return { inert: root.inert, busy: root.getAttribute("aria-busy"),
      focusedInside: root.contains(document.activeElement) };
  });
  assert.deepEqual(state, { inert: true, busy: "true", focusedInside: false });
  const button = await page.locator(".hwe-toolbar button").first().boundingBox();
  await page.mouse.click(button.x + button.width / 2, button.y + button.height / 2);
  await page.keyboard.press("Tab");
  assert.equal(await page.evaluate(() => window.blockedClicks), 0, "la pantalla intercepta los clics");
  assert.equal(await page.evaluate(() => document.querySelector(".pcf-html-editor-root").contains(document.activeElement)), false,
    "el teclado no entra en el editor mientras carga");
}

async function assertReady(page) {
  await page.locator(".hwe-loading-overlay").waitFor({ state: "hidden" });
  assert.deepEqual(await page.evaluate(() => {
    const root = document.querySelector(".pcf-html-editor-root");
    return { inert: root.inert, busy: root.getAttribute("aria-busy") };
  }), { inert: false, busy: "false" });
  await page.locator(".hwe-workspace").focus();
  assert.equal(await page.evaluate(() => document.activeElement.classList.contains("hwe-workspace")), true);
}

test("bloquea ratón y teclado hasta que termina la descarga y después permite editar", async t => {
  const page = await open(t);
  await page.evaluate(() => {
    const download = new Promise(resolve => { window.finishDownload = resolve; });
    window.editor = new window.EditorComponent(document.getElementById("host"), undefined, {
      loadHtml: () => download, paragraphStyleCatalog: { styles: [], fonts: [] },
    });
    window.ready = window.editor.init();
  });
  await assertBlocked(page);
  await page.evaluate(() => window.finishDownload("<p>Documento cargado</p>"));
  await page.evaluate(() => window.ready);
  await assertReady(page);
  await page.keyboard.press("Control+End");
  await page.keyboard.type(" editable");
  assert.match(await page.locator(".hwe-page-inner").innerText(), /Documento cargado editable/);
});

test("mantiene el bloqueo durante la preparación inicial del HTML de impresión", async t => {
  const page = await open(t);
  await page.evaluate(() => {
    const printing = new Promise(resolve => { window.finishPrinting = resolve; });
    window.fetch = async () => new Response("<p>Contenido del registro</p>");
    window.editor = new window.EditorComponent(document.getElementById("host"), {
      page: { getClientUrl: () => "https://example.invalid", entityId: "test-record" },
    }, { paragraphStyleCatalog: { styles: [], fonts: [] } });
    window.editor.loadDynamicHeader = async () => {};
    window.editor.loadRuntimeHeaderLogo = async () => {};
    window.editor.ensureInitialPrintHtmlSaved = async () => { window.preparingPrint = true; await printing; };
    window.ready = window.editor.init();
  });
  await page.waitForFunction(() => window.preparingPrint);
  await assertBlocked(page);
  await page.evaluate(() => window.finishPrinting());
  await page.evaluate(() => window.ready);
  await assertReady(page);
});

test("espera a las imágenes grandes y al ajuste de sus páginas", async t => {
  const page = await open(t);
  await page.evaluate(() => {
    const canvas = document.createElement("canvas");
    canvas.width = 1600; canvas.height = 1000;
    canvas.getContext("2d").fillRect(0, 0, canvas.width, canvas.height);
    const hydration = new Promise(resolve => { window.finishImages = resolve; });
    window.editor = new window.EditorComponent(document.getElementById("host"), undefined, {
      initialHtml: `<p>Imagen grande</p><img src="${canvas.toDataURL()}" width="1600" height="1000">`,
      paragraphStyleCatalog: { styles: [], fonts: [] },
    });
    const assets = window.editor.assetLayoutManager;
    const original = assets.hydrateDetachedImages.bind(assets);
    assets.hydrateDetachedImages = async (...args) => {
      window.preparingImages = true;
      await hydration;
      return original(...args);
    };
    window.ready = window.editor.init();
  });
  await page.waitForFunction(() => window.preparingImages);
  assert.equal(await page.locator("[data-hwe-large-image-placeholder]").count(), 1);
  await assertBlocked(page);
  await page.evaluate(() => window.finishImages());
  await page.evaluate(() => window.ready);
  await assertReady(page);
  assert.equal(await page.locator("[data-hwe-large-image-placeholder]").count(), 0);
  assert.equal(await page.evaluate(() => window.editor.rebalanceFrame), undefined);
});

test("vuelve a bloquear al cargar otro HTML y se libera si la carga falla", async t => {
  const page = await open(t);
  await page.evaluate(async () => {
    window.editor = new window.EditorComponent(document.getElementById("host"), undefined, {
      initialHtml: "<p>Anterior</p>", paragraphStyleCatalog: { styles: [], fonts: [] },
    });
    await window.editor.init();
    const layout = new Promise(resolve => { window.finishLayout = resolve; });
    const assets = window.editor.assetLayoutManager;
    const original = assets.waitForStableLayout.bind(assets);
    assets.waitForStableLayout = async (...args) => {
      window.preparingLayout = true;
      await layout;
      return original(...args);
    };
    window.ready = window.editor.loadHtml("<p>Nuevo documento</p>");
  });
  await page.waitForFunction(() => window.preparingLayout);
  await assertBlocked(page);
  await page.evaluate(() => window.finishLayout());
  await page.evaluate(() => window.ready);
  await assertReady(page);
  assert.equal(await page.locator(".hwe-page-inner").innerText(), "Nuevo documento");
  await page.evaluate(async () => {
    window.editor.renderAndPaginate = async () => { throw new Error("No se pudo procesar el HTML"); };
    await window.editor.loadHtml("<p>Fallo</p>").catch(() => {});
  });
  await assertReady(page);
  assert.match(await page.locator(".hwe-status-msg.error").innerText(), /No se pudo procesar el HTML/);
});

test("retira la pantalla y muestra el error si falla la descarga inicial", async t => {
  const page = await open(t);
  await page.evaluate(async () => {
    window.editor = new window.EditorComponent(document.getElementById("host"), undefined, {
      loadHtml: async () => { throw new Error("Error de red"); }, paragraphStyleCatalog: { styles: [], fonts: [] },
    });
    await window.editor.init();
  });
  await assertReady(page);
  assert.match(await page.locator(".hwe-status-msg.error").innerText(), /Error de red/);
});

test("no habilita un editor inicialmente oculto hasta paginar al hacerse visible", async t => {
  const page = await open(t, true);
  await page.evaluate(async () => {
    window.editor = new window.EditorComponent(document.getElementById("host"), undefined, {
      initialHtml: "<p>Documento inicialmente oculto. Texto de prueba.</p>".repeat(160),
      paragraphStyleCatalog: { styles: [], fonts: [] },
    });
    await window.editor.init();
  });
  assert.equal(await page.evaluate(() => document.querySelector(".pcf-html-editor-root").inert), true);
  await page.evaluate(() => { document.getElementById("surface").style.display = "block"; });
  await assertReady(page);
  assert.ok(await page.locator(".hwe-page").count() > 1);
});

test("una fila indivisible más alta que una página no deja bloqueada la carga", {timeout:30000}, async t => {
  const page = await open(t);
  await page.evaluate(async () => {
    window.editor = new EditorComponent(document.querySelector('#host'), undefined, {
      initialHtml:'<table><tbody><tr><td><p>Fila indivisible</p><div style="height:1500px">Contenido alto</div></td></tr></tbody></table>',
      paragraphStyleCatalog:{styles:[],fonts:[]}
    });
    await editor.init();
  });
  await assertReady(page);
  assert.equal(await page.locator('tbody tr').count(), 1);
  await page.evaluate(() => editor.destroy());
});

test("cancelar durante una tanda pendiente no continúa paginando", async t => {
  const page = await open(t);
  const result = await page.evaluate(async () => {
    const editor = new EditorComponent(document.querySelector('#host'), undefined, {
      initialHtml:'<p>Documento</p>', paragraphStyleCatalog:{styles:[],fonts:[]}
    });
    await editor.init();
    let resume;
    editor.waitForLoadingPaint = () => new Promise(resolve => { resume = resolve; });
    editor.layoutService.pageOverflows = () => true;
    let pushes = 0;
    editor.paginator.pushOverflowForwardFromPage = () => { pushes++; };
    const pending = editor.finishInitialPagination();
    editor.destroy();
    resume();
    await pending;
    return pushes;
  });
  assert.equal(result,0);
});
