const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");
const ts = require("typescript");
const { JSDOM } = require("jsdom");

function loadTypeScript(filename, window, dependencies = {}) {
  const source = fs.readFileSync(filename, "utf8");
  const { outputText } = ts.transpileModule(source, {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
    fileName: filename,
  });
  const exports = {};
  vm.runInNewContext(outputText, {
    exports,
    require: (specifier) => dependencies[specifier],
    window,
    document: window.document,
    Node: window.Node,
    NodeFilter: window.NodeFilter,
    Range: window.Range,
    HTMLElement: window.HTMLElement,
  }, { filename });
  return exports;
}

function createController(markup) {
  const dom = new JSDOM("<!doctype html><html><body><main class='hwe-workspace'></main></body></html>");
  const { window } = dom;
  const root = window.document.querySelector("main");
  root.innerHTML = markup;
  const domPath = path.resolve(__dirname, "../EditorHTML/dom/EditableDom.ts");
  const editableDom = loadTypeScript(domPath, window);
  const controllerPath = path.resolve(__dirname, "../EditorHTML/controllers/PageBackspaceController.ts");
  const { PageBackspaceController } = loadTypeScript(controllerPath, window, { "../dom/EditableDom": editableDom });
  return { dom, root, controller: new PageBackspaceController() };
}

const page = (contents) => `<section class="hwe-page"><div class="hwe-page-inner" contenteditable="true">${contents}</div></section>`;

function backspaceAtStart(app, selector) {
  const text = app.root.querySelector(selector).firstChild;
  const range = app.root.ownerDocument.createRange();
  range.setStart(text, 0);
  range.collapse(true);
  app.dom.window.getSelection().removeAllRanges();
  app.dom.window.getSelection().addRange(range);
  return app.controller.handleBoundaryDelete("backward", range, {
    pages: Array.from(app.root.querySelectorAll(".hwe-page")),
    onContentChanged: () => undefined,
  });
}

test("Backspace une párrafos envueltos por contenedores generados", (t) => {
  const app = createController(
    page("<div data-hwe-generated-wrapper='true'><p>anterior</p></div>") +
    page("<div data-hwe-generated-wrapper='true'><p>siguiente</p></div>")
  );
  t.after(() => app.dom.window.close());
  assert.ok(backspaceAtStart(app, ".hwe-page:last-child p"));
  assert.equal(app.root.querySelector(".hwe-page-inner p").textContent, "anteriorsiguiente");
  assert.equal(app.root.querySelectorAll(".hwe-page:last-child p").length, 0);
});

test("Backspace no salta un HR directo para unir bloques no adyacentes", (t) => {
  const app = createController(page("<hr><p>actual</p>") + page("<p>previo</p>"));
  t.after(() => app.dom.window.close());
  assert.equal(backspaceAtStart(app, ".hwe-page:first-child p"), null);
  assert.ok(app.root.querySelector("hr"));
  assert.equal(app.root.querySelectorAll("p").length, 2);
});

test("Backspace no une elementos LI de listas distintas", (t) => {
  const app = createController(page("<ul><li>anterior</li></ul>") + page("<ul><li>siguiente</li></ul>"));
  t.after(() => app.dom.window.close());
  assert.ok(backspaceAtStart(app, ".hwe-page:last-child li"));
  assert.equal(app.root.querySelectorAll("li").length, 2);
  assert.equal(app.root.querySelectorAll("ul").length, 2);
});

test("Backspace elimina una línea con formato y conserva el flujo del párrafo al volver de página", (t) => {
  const app = createController(
    page("<div class='hwe-text-flow-column'><p data-hwe-user-blank='true'><strong><br></strong></p></div>") +
    page("<div class='hwe-text-flow-column'><p data-hwe-text-flow-id='texto'><strong>Parrafo 21.</strong> Texto</p></div>")
  );
  t.after(() => app.dom.window.close());
  assert.ok(backspaceAtStart(app, ".hwe-page:last-child strong"));
  const paragraph = app.root.querySelector("p");
  assert.equal(paragraph.textContent, "Parrafo 21. Texto");
  assert.equal(paragraph.getAttribute("data-hwe-text-flow-id"), "texto");
  assert.equal(paragraph.querySelectorAll("br").length, 0);
  assert.equal(app.dom.window.getSelection().anchorOffset, 0);
});

test("Backspace sigue borrando líneas vacías dentro de la misma columna", (t) => {
  const app = createController(page("<div class='hwe-text-flow-column'>" +
    "<p><strong><br></strong></p><p><strong><br></strong></p><p id='texto'>Parrafo 21.</p></div>"));
  t.after(() => app.dom.window.close());
  assert.ok(backspaceAtStart(app, "#texto"));
  assert.equal(app.root.querySelectorAll("p").length, 2);
  assert.ok(backspaceAtStart(app, "#texto"));
  assert.equal(app.root.querySelectorAll("p").length, 1);
  assert.equal(app.root.querySelector("p").textContent, "Parrafo 21.");
});

test("unir párrafos actualiza las continuaciones para que no queden fragmentos separados", (t) => {
  const app = createController(page("<div class='hwe-text-flow-column'><p data-hwe-text-flow-id='a'>Antes</p>" +
    "<p id='texto' data-hwe-text-flow-id='b'> nuevo</p></div>") +
    page("<div class='hwe-text-flow-column'><p data-hwe-text-flow-id='b'> final</p></div>"));
  t.after(() => app.dom.window.close());
  assert.ok(backspaceAtStart(app, "#texto"));
  assert.equal(app.root.querySelector("p").textContent, "Antes nuevo");
  assert.equal(app.root.querySelector(".hwe-page:last-child p").getAttribute("data-hwe-text-flow-id"), "a");
});

test("Backspace en una continuación automática borra un carácter en vez de volver al mismo salto", (t) => {
  const app = createController(
    page("<div class='hwe-text-flow-column'><p data-hwe-text-flow-id='texto'>antes😀</p></div>") +
    page("<div class='hwe-text-flow-column'><p data-hwe-text-flow-id='texto'>después</p></div>")
  );
  t.after(() => app.dom.window.close());
  assert.ok(backspaceAtStart(app, ".hwe-page:last-child p"));
  assert.equal(app.root.querySelector("p").textContent, "antesdespués");
  const selection = app.dom.window.getSelection();
  assert.equal(selection.anchorNode.textContent, "antes");
  assert.equal(selection.anchorOffset, 5);
});
