const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const test = require("node:test");
const ts = require("typescript");
const { JSDOM } = require("jsdom");

function setup(t, pages) {
  const dom = new JSDOM("<!doctype html><html><body><main></main></body></html>", { pretendToBeVisual: true });
  const { window } = dom;
  const cache = new Map();
  function load(relativePath, mockDependencies = false) {
    const filename = path.resolve(__dirname, relativePath);
    if (cache.has(filename)) return cache.get(filename);
    const { outputText } = ts.transpileModule(fs.readFileSync(filename, "utf8"), {
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
    });
    const exports = {};
    cache.set(filename, exports);
    vm.runInNewContext(outputText, {
      exports, window, document: window.document,
      require: (name) => mockDependencies
        ? (name === "../ui/Toolbar" ? { CLEAR_PARAGRAPH_STYLE_VALUE: "__hwe-clear-paragraph-style" } : {})
        : load(path.resolve(path.dirname(filename), name + ".ts")),
      Node: window.Node, NodeFilter: window.NodeFilter, HTMLElement: window.HTMLElement,
      HTMLTableElement: window.HTMLTableElement, HTMLFontElement: window.HTMLFontElement,
    }, { filename });
    return exports;
  }
  const { TableSelectionController } = load("../EditorHTML/controllers/TableSelectionController.ts");
  const { ParagraphStyleManager } = load("../EditorHTML/services/ParagraphStyleManager.ts");
  const { EditorHistoryController } = load("../EditorHTML/controllers/EditorHistoryController.ts");
  const { EditorComponent } = load("../EditorHTML/app/EditorComponent.ts", true);
  const root = window.document.querySelector("main");
  root.className = "pcf-html-editor-root";
  const workspace = window.document.createElement("div");
  root.appendChild(workspace);
  workspace.innerHTML = pages.map((html, index) =>
    `<section class="hwe-page" id="page-${index}"><div class="hwe-page-inner" contenteditable="true">${html}</div></section>`
  ).join("");
  const manager = new ParagraphStyleManager(() => root, { setParagraphStyles() {}, setFontFamilies() {} });
  manager.setStyles([
    { label: "Título", className: "titulo", cssText: "font-family: Georgia; font-size: 16pt; line-height: 1.4;" },
    { label: "Cuerpo", className: "cuerpo", cssText: "font-size: 11pt;" },
  ]);
  const calls = { rebalance: [], dirty: [], fallback: 0 };
  const original = workspace.innerHTML;
  const history = new EditorHistoryController({
    collectSnapshot: () => workspace.innerHTML,
    restoreSnapshot: (html) => { workspace.innerHTML = html; },
  });
  history.reset();
  const editor = Object.create(EditorComponent.prototype);
  Object.assign(editor, {
    root, paragraphStyleManager: manager, historyController: history,
    pages: Array.from(workspace.querySelectorAll(".hwe-page")),
    scheduleRebalance: (page, pull, options) => calls.rebalance.push({ page: page.id, pull, force: options.force }),
    collectHtml: () => workspace.innerHTML,
    updateDirtyState: (html) => calls.dirty.push(html !== original),
    styleSelectionTracker: { restoreTextSelection() { calls.fallback++; }, getSelectedStyleBlocks: () => [] },
    setStatus() {},
  });
  const controller = new TableSelectionController({
    rootProvider: () => root,
    onTableChanged: (table) => editor.markTableTextFormattingChanged(table),
    onBeforeTableDelete() {}, onTableDeleted() {},
  });
  editor.tableSelectionController = controller;
  controller.start();
  const select = (cellId, mode) => {
    const cell = workspace.querySelector("#" + cellId);
    cell.dispatchEvent(new window.MouseEvent("pointerdown", { bubbles: true }));
    controller.selectTable(cell.closest("table"));
    if (mode !== "table") {
      Array.from(root.querySelectorAll("button")).find(button => button.textContent === (mode === "row" ? "Fila" : "Columna")).click();
    }
  };
  t.after(() => { controller.destroy(); manager.destroy(); history.destroy(); window.close(); });
  return { root, workspace, select, controller, calls, history, original,
    apply: (name = "titulo") => editor.applyParagraphStyle(name),
    styled: () => Array.from(workspace.querySelectorAll(".titulo")),
  };
}

test("aplica el estilo a cada párrafo de la fila, incluido texto suelto, y deshace toda la operación", async (t) => {
  const app = setup(t, ['<p id="outside">Fuera</p><table><tr><td id="a">Texto <b>negrita</b></td><td id="b"><p class="cuerpo">Uno</p><p>Dos</p></td></tr><tr><td id="other">Intacto</td><td>Otro</td></tr></table>']);
  app.select("a", "row");
  app.apply();
  assert.deepEqual(app.styled().map(p => p.textContent), ["Texto negrita", "Uno", "Dos"]);
  assert.ok(app.workspace.querySelector("#a > p.titulo > b"));
  assert.equal(app.workspace.querySelector("#other").innerHTML, "Intacto");
  assert.equal(app.workspace.querySelector("#outside").className, "");
  assert.equal(app.workspace.querySelector("td.titulo, tr.titulo, table.titulo, .cuerpo"), null);
  assert.equal(app.calls.fallback, 0);
  assert.deepEqual(app.calls.dirty, [true]);
  assert.deepEqual(app.calls.rebalance, [{ page: "page-0", pull: true, force: true }]);
  const applied = app.workspace.innerHTML;
  app.history.undo();
  await Promise.resolve();
  assert.equal(app.workspace.innerHTML, app.original);
  assert.equal(app.history.canUndo, false);
  app.history.redo();
  await Promise.resolve();
  assert.equal(app.workspace.innerHTML, applied);
});

test("aplica a la columna en todos los fragmentos y respeta celdas combinadas", (t) => {
  const app = setup(t, [
    '<table data-hwe-table-flow-id="t"><tr><td id="span" colspan="2">Combinada</td><td>Fuera</td></tr><tr><td>A</td><td id="b">B</td><td>C</td></tr></table>',
    '<table data-hwe-table-flow-id="t"><tr><td>D</td><td id="e">E</td><td>F</td></tr></table>',
  ]);
  app.select("e", "column");
  app.apply();
  assert.deepEqual(app.styled().map(p => p.textContent), ["Combinada", "B", "E"]);
  assert.equal(app.workspace.querySelector("#span").colSpan, 2);
  assert.equal(app.calls.rebalance[0].page, "page-0");
});

test("Sin estilo quita las clases de los párrafos seleccionados y conserva los demás", (t) => {
  const app = setup(t, ['<table><tr><td id="a"><p>Uno</p></td><td id="b"><p>Dos</p></td></tr><tr><td>Tres</td><td>Cuatro</td></tr></table>']);
  app.select("a", "table");
  app.apply();
  app.select("b", "column");
  app.apply("__hwe-clear-paragraph-style");
  assert.deepEqual(app.styled().map(p => p.textContent), ["Uno", "Tres"]);
  assert.equal(app.workspace.querySelector("#b p").hasAttribute("data-hwe-paragraph-style"), false);
});

test("respeta tablas anidadas y contenido protegido, y prepara celdas vacías", (t) => {
  const app = setup(t, ['<table><tr><td id="a">Antes<div><p>Interior</p><p hidden>Oculto</p></div><table><tr><td>Anidado</td></tr></table>Después</td><td contenteditable="false">Protegido</td><td id="empty">  </td></tr></table>']);
  app.select("a", "row");
  app.apply();
  assert.deepEqual(app.styled().map(p => p.textContent), ["Antes", "Interior", "Después", ""]);
  assert.ok(app.workspace.querySelector("#empty > p.titulo > br"));
  assert.equal(app.workspace.querySelector("td td").innerHTML, "Anidado");
  assert.equal(app.workspace.querySelector("[hidden]").className, "");
});

test("la fila de cabecera repetida recibe el estilo en todos sus fragmentos", (t) => {
  const app = setup(t, [0, 1].map(i => `<table data-hwe-table-flow-id="t"><thead><tr><th id="h${i}">Cabecera</th></tr></thead><tbody><tr><td>Dato</td></tr></tbody></table>`));
  app.select("h1", "row");
  app.apply();
  assert.deepEqual(app.styled().map(p => p.textContent), ["Cabecera", "Cabecera"]);
});

test("los estilos no válidos no modifican el documento y sin selección de tabla se usa el flujo de texto", (t) => {
  const app = setup(t, ['<table><tr><td id="a">Texto</td></tr></table>']);
  app.select("a", "row");
  app.apply("no-existe");
  assert.equal(app.workspace.innerHTML, app.original);
  app.controller.clearSelection();
  app.apply();
  assert.equal(app.calls.fallback, 1);
});
