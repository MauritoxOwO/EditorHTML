const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");
const ts = require("typescript");
const { JSDOM } = require("jsdom");

function loadSource(relativePath, window) {
  const filename = path.join(__dirname, relativePath);
  const { outputText } = ts.transpileModule(fs.readFileSync(filename, "utf8"), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
    fileName: filename,
  });
  const exports = {};
  vm.runInNewContext(outputText, {
    exports, require: () => ({}), window, document: window.document,
    Node: window.Node, NodeFilter: window.NodeFilter,
    MutationObserver: window.MutationObserver,
  }, { filename });
  return exports;
}

function setup(t, pages) {
  const dom = new JSDOM("<!doctype html><html><body><main></main></body></html>");
  const { window } = dom;
  // jsdom no calcula geometría; solo la navegación individual la necesita.
  window.Range.prototype.getBoundingClientRect = () => ({ top: 0, left: 0 });
  const { DocumentSearch } = loadSource("../EditorHTML/ui/DocumentSearch.ts", window);
  const { EditorHistoryController } = loadSource("../EditorHTML/controllers/EditorHistoryController.ts", window);
  const { EditorComponent } = loadSource("../EditorHTML/app/EditorComponent.ts", window);
  const workspace = window.document.querySelector("main");
  workspace.innerHTML = pages.map((html, index) =>
    `<section class="hwe-page" id="page-${index}"><div class="hwe-page-inner" contenteditable="true">${html}</div></section>`
  ).join("");
  const original = workspace.innerHTML;
  const calls = { before: 0, dirty: [], rebalance: [] };
  const history = new EditorHistoryController({
    collectSnapshot: () => workspace.innerHTML,
    restoreSnapshot: (html) => { workspace.innerHTML = html; },
    onRestored: () => calls.dirty.push(workspace.innerHTML !== original),
  });
  history.reset();
  const editor = Object.create(EditorComponent.prototype);
  Object.assign(editor, {
    historyController: history,
    updateToolbarSelectionState() {},
    scheduleRebalance: (page, pull, options) => calls.rebalance.push({ id: page.id, pull, ...options }),
    collectHtml: () => workspace.innerHTML,
    updateDirtyState: (html) => calls.dirty.push(html !== original),
  });
  const search = new DocumentSearch(workspace,
    () => { calls.before++; history.recordNow(); },
    (page, force) => editor.markEditedAndRebalance(page, force));
  const bar = search.build();
  window.document.body.prepend(bar);
  const button = (label) => Array.from(bar.querySelectorAll("button")).find((item) => item.textContent === label);
  const setQuery = (value, replacement = "") => {
    bar.querySelector('input[type="search"]').value = value;
    bar.querySelector('input[type="text"]').value = replacement;
    bar.querySelector('input[type="search"]').dispatchEvent(new window.Event("input"));
  };
  t.after(() => { search.destroy(); history.destroy(); window.close(); });
  return { workspace, bar, button, setQuery, calls, history, original };
}

test("reemplaza en varias páginas, registra una edición y permite deshacer y rehacer", async (t) => {
  const app = setup(t, ["<p>Sin cambios</p>", "<p>sol sol</p>", "<p>sol</p>"]);
  app.setQuery("sol", "luna");
  app.button("Reemplazar todo").click();
  assert.deepEqual(Array.from(app.workspace.querySelectorAll("p"), p => p.textContent), ["Sin cambios", "luna luna", "luna"]);
  assert.equal(app.calls.before, 1);
  assert.deepEqual(app.calls.rebalance, [{ id: "page-1", pull: true, includePreviousPage: false, force: true }]);
  assert.deepEqual(app.calls.dirty, [true]);
  assert.equal(app.button("Reemplazar todo").disabled, true);
  const replaced = app.workspace.innerHTML;
  app.history.undo();
  await Promise.resolve();
  assert.equal(app.workspace.innerHTML, app.original);
  assert.equal(app.calls.dirty.at(-1), false);
  assert.equal(app.history.canUndo, false);
  app.history.redo();
  await Promise.resolve();
  assert.equal(app.workspace.innerHTML, replaced);
});

test("no vuelve a reemplazar el texto insertado y conserva el formato exterior", (t) => {
  const app = setup(t, ["<p><strong>aaaa</strong> y <em>aa</em></p>"]);
  app.setQuery("aa", "aaa");
  app.button("Reemplazar todo").click();
  assert.equal(app.workspace.querySelector("strong").textContent, "aaaaaa");
  assert.equal(app.workspace.querySelector("em").textContent, "aaa");
  assert.equal(app.calls.before, 1);
});

test("encuentra palabras entre spans y respeta mayúsculas y minúsculas", (t) => {
  const app = setup(t, ["<p><span>sw</span><span>iss</span> Swiss SWISS</p>"]);
  app.setQuery("swiss", "fuente");
  app.button("Aa").click();
  app.button("Reemplazar todo").click();
  assert.equal(app.workspace.textContent, "fuente Swiss SWISS");
  app.button("Aa").click();
  app.button("Reemplazar todo").click();
  assert.equal(app.workspace.textContent, "fuente fuente fuente");
});

test("admite eliminación y trata la búsqueda y el reemplazo como texto literal", (t) => {
  const app = setup(t, ["<p>a.b a.b axb</p>"]);
  app.setQuery("a.b", "<b>$&</b>");
  app.button("Reemplazar todo").click();
  assert.equal(app.workspace.textContent, "<b>$&</b> <b>$&</b> axb");
  assert.equal(app.workspace.querySelector("b"), null);
  app.setQuery("<b>$&</b>");
  app.button("Reemplazar todo").click();
  assert.equal(app.workspace.textContent, "  axb");
});

test("no cambia cabeceras, contenido protegido ni búsquedas sin coincidencias", (t) => {
  const app = setup(t, ["<div contenteditable='false'>sol</div><div data-hwe-api-header>sol</div><p hidden>sol</p><p>sol</p>"]);
  for (const query of ["", "   ", "ausente"]) {
    app.setQuery(query, "luna");
    assert.equal(app.button("Reemplazar todo").disabled, true);
    app.button("Reemplazar todo").click();
  }
  assert.equal(app.calls.before, 0);
  app.setQuery("sol", "luna");
  app.button("Reemplazar todo").click();
  assert.equal(app.workspace.textContent, "solsolsolluna");
});

test("el reemplazo individual sigue cambiando únicamente la coincidencia seleccionada", (t) => {
  const app = setup(t, ["<p>sol sol</p>"]);
  app.setQuery("sol", "luna");
  assert.equal(app.button("Reemplazar").disabled, true);
  app.button("Siguiente").click();
  app.button("Reemplazar").click();
  assert.equal(app.workspace.textContent, "luna sol");
  assert.equal(app.calls.rebalance[0].force, false);
});
