const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");
const ts = require("typescript");
const { JSDOM } = require("jsdom");

function setup(t, pages) {
  const dom = new JSDOM("<!doctype html><html><body><main></main></body></html>");
  const { window } = dom;
  function load(relativePath) {
    const filename = path.resolve(__dirname, relativePath);
    const { outputText } = ts.transpileModule(fs.readFileSync(filename, "utf8"), {
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
      fileName: filename,
    });
    const exports = {};
    vm.runInNewContext(outputText, {
      exports, window, document: window.document,
      require: (name) => load(path.resolve(path.dirname(filename), name + ".ts")),
      Element: window.Element, NodeFilter: window.NodeFilter,
      HTMLTableElement: window.HTMLTableElement,
    }, { filename });
    return exports;
  }
  const { TableSelectionController } = load("../EditorHTML/controllers/TableSelectionController.ts");
  const { EditorHistoryController } = load("../EditorHTML/controllers/EditorHistoryController.ts");
  const root = window.document.querySelector("main");
  const workspace = window.document.createElement("div");
  root.appendChild(workspace);
  workspace.innerHTML = pages.map((html, index) =>
    `<section class="hwe-page" id="page-${index}"><div class="hwe-page-inner" contenteditable="true" tabindex="0">${html}</div></section>`
  ).join("");
  const original = workspace.innerHTML;
  const calls = { before: 0, pages: [] };
  const history = new EditorHistoryController({
    collectSnapshot: () => workspace.innerHTML,
    restoreSnapshot: (html) => { workspace.innerHTML = html; },
  });
  history.reset();
  const controller = new TableSelectionController({
    rootProvider: () => root,
    onTableChanged() {},
    onBeforeTableDelete: () => { calls.before++; history.recordNow(); },
    onTableDeleted: (page) => { calls.pages.push(page.id); history.recordNow(); },
  });
  controller.start();
  const select = (selector, mode = "table") => {
    const table = workspace.querySelector(selector);
    table.querySelector("td").dispatchEvent(new window.MouseEvent("pointerdown", { bubbles: true }));
    controller.selectTable(table);
    if (mode !== "table") {
      const label = mode === "row" ? "Fila" : "Columna";
      Array.from(root.querySelectorAll("button")).find(button => button.textContent === label).click();
    }
  };
  const remove = () => root.querySelector(".hwe-table-delete").click();
  t.after(() => { controller.destroy(); history.destroy(); window.close(); });
  return { window, root, workspace, controller, history, original, calls, select, remove };
}

const table = (id, content = "contenido", flow = "") =>
  `<table id="${id}" ${flow ? `data-hwe-table-flow-id="${flow}"` : ""}><tbody><tr><td>${content}</td></tr></tbody></table>`;

test("elimina toda la tabla paginada, conserva otras tablas y permite deshacer en un paso", async (t) => {
  const app = setup(t, [
    "<p>Antes</p>" + table("first", "primera parte", "flow-1"),
    table("second", "segunda parte", "flow-1") + "<p>Después</p>" + table("other"),
  ]);
  // Se selecciona desde la continuación, pero se reordena desde la primera página.
  app.select("#second");
  assert.equal(app.root.querySelector(".hwe-table-delete").previousElementSibling.textContent, "Columna");
  app.remove();
  assert.equal(app.workspace.querySelectorAll("table").length, 1);
  assert.equal(app.workspace.querySelector("table").id, "other");
  assert.match(app.workspace.textContent, /AntesDespuéscontenido/);
  assert.equal(app.calls.before, 1);
  assert.deepEqual(app.calls.pages, ["page-0"]);
  assert.equal(app.root.querySelector(".hwe-table-selection-toolbar").style.display, "none");
  assert.equal(app.controller.handleToolbarCommand("bold"), false);
  assert.equal(app.window.getSelection().anchorNode.nodeName, "P");
  assert.equal(app.window.getSelection().anchorNode.isConnected, true);
  const deleted = app.workspace.innerHTML;
  app.history.undo();
  await Promise.resolve();
  assert.equal(app.workspace.innerHTML, app.original);
  assert.equal(app.history.canUndo, false);
  app.history.redo();
  await Promise.resolve();
  assert.equal(app.workspace.innerHTML, deleted);
});

for (const mode of ["row", "column"]) {
  test(`el botón elimina la tabla completa aunque esté seleccionada una ${mode}`, (t) => {
    const app = setup(t, [table("target", "uno</td><td>dos</td></tr><tr><td>tres</td><td>cuatro")]);
    app.select("#target", mode);
    app.remove();
    assert.equal(app.workspace.querySelector("table"), null);
    assert.equal(app.workspace.querySelector(".hwe-page-inner").innerHTML, "<p><br></p>");
    assert.deepEqual(app.calls.pages, ["page-0"]);
  });
}

test("al eliminar una tabla anidada se conserva la tabla exterior", (t) => {
  const app = setup(t, [table("outer", "Texto " + table("nested"))]);
  app.select("#nested");
  app.remove();
  assert.equal(app.workspace.querySelectorAll("table").length, 1);
  assert.equal(app.workspace.querySelector("table").id, "outer");
  assert.equal(app.workspace.querySelector("td").innerHTML, "Texto <p><br></p>");
});

test("no elimina tablas protegidas ni vuelve a actuar sobre una selección eliminada", (t) => {
  const app = setup(t, ["<div contenteditable='false'>" + table("protected") + "</div>" + table("target")]);
  app.select("#protected");
  app.remove();
  assert.ok(app.workspace.querySelector("#protected"));
  assert.equal(app.calls.before, 0);
  app.select("#target");
  app.remove();
  app.remove();
  assert.equal(app.calls.before, 1);
  assert.ok(app.workspace.querySelector("#protected"));
});
