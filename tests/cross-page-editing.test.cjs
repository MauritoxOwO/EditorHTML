const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");
const ts = require("typescript");
const { JSDOM } = require("jsdom");

function createSelection(markup) {
  const dom = new JSDOM("<!doctype html><html><body><main class='hwe-workspace'></main></body></html>");
  const { window } = dom;
  const filename = path.resolve(__dirname, "../EditorHTML/selection/DocumentSelection.ts");
  const { outputText } = ts.transpileModule(fs.readFileSync(filename, "utf8"), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
    fileName: filename,
  });
  const exports = {};
  vm.runInNewContext(outputText, {
    exports, window, document: window.document, Node: window.Node,
    NodeFilter: window.NodeFilter, Range: window.Range, HTMLElement: window.HTMLElement,
  }, { filename });
  const root = window.document.querySelector("main");
  root.innerHTML = markup;
  return { dom, root, selection: new exports.DocumentSelection(() => root) };
}

const page = (html) => `<section class="hwe-page"><div class="hwe-page-inner" contenteditable="true">${html}</div></section>`;

test("copiar un rango entre fragmentos omite headers, divisores y el encabezado repetido", (t) => {
  const app = createSelection(
    page("<div data-hwe-api-header>HEADER DE PÁGINA</div><p>Inicio seleccionado</p>" +
      "<table data-hwe-repeat-header='true' data-hwe-table-flow-id='tabla'><thead><tr><th><strong>Columna</strong></th><th>Otra</th></tr></thead><tbody><tr><td>Fila 1</td></tr></tbody></table><p>medio</p>") +
    page("<div data-hwe-runtime-page-header>HEADER REPETIDO</div><table data-hwe-repeat-header='true' data-hwe-table-flow-id='tabla'><thead><tr><th><strong>Columna</strong></th><th>Otra</th></tr></thead><tbody><tr><td>Fila 2</td></tr></tbody></table><p>Final seleccionado</p><div data-hwe-page-divider>DIVISOR</div>")
  );
  t.after(() => app.dom.window.close());
  const start = app.root.querySelector(".hwe-page:first-child p").firstChild;
  const end = app.root.querySelector(".hwe-page:last-child p").firstChild;
  const range = app.root.ownerDocument.createRange();
  range.setStart(start, 0);
  range.setEnd(end, end.length);
  const copied = app.selection.getClipboardContent(range);
  assert.ok(copied);
  assert.match(copied.html, /Inicio seleccionado/);
  assert.match(copied.html, /Final seleccionado/);
  assert.equal((copied.html.match(/<thead/gi) ?? []).length, 1);
  assert.doesNotMatch(copied.html, /HEADER DE PÁGINA|HEADER REPETIDO|DIVISOR|data-hwe-page/);
  assert.equal((copied.text.match(/Columna/g) ?? []).length, 1);
  assert.equal((copied.text.match(/Otra/g) ?? []).length, 1);
  assert.match(copied.text, /Fila 1/);
  assert.match(copied.text, /Fila 2/);
});

test("copiar omite todos los textos de cabeceras repetidas y reúne fragmentos del mismo párrafo", (t) => {
  const app = createSelection(page("<div class='hwe-text-flow-column'><p data-hwe-text-flow-id='flujo'>inicio</p></div>") +
    page("<div class='hwe-text-flow-column'><p data-hwe-text-flow-id='flujo'> final</p></div>"));
  t.after(() => app.dom.window.close());
  const start = app.root.querySelector(".hwe-page:first-child p").firstChild;
  const end = app.root.querySelector(".hwe-page:last-child p").firstChild;
  const range = app.root.ownerDocument.createRange();
  range.setStart(start, 0);
  range.setEnd(end, end.length);
  const copied = app.selection.getClipboardContent(range);
  assert.ok(copied);
  assert.equal((copied.html.match(/<p\b/gi) ?? []).length, 1);
  assert.match(copied.text, /inicio final/);
  assert.doesNotMatch(copied.html, /data-hwe-/);
  assert.doesNotMatch(copied.html, /hwe-text-flow-column/);
});

test("el texto copiado conserva saltos de línea y celdas vacías", (t) => {
  const app = createSelection(page("<p>uno<br>dos</p><table><tbody><tr><td>A</td><td></td><td>B</td></tr></tbody></table>"));
  t.after(() => app.dom.window.close());
  const first = app.root.querySelector("p").firstChild;
  const last = app.root.querySelector("td:last-child").firstChild;
  const range = app.root.ownerDocument.createRange();
  range.setStart(first, 0);
  range.setEnd(last, last.length);
  assert.equal(app.selection.getClipboardContent(range).text, "uno\ndos\nA\t\tB");
});

test("el texto copiado conserva párrafos vacíos entre párrafos", (t) => {
  const app = createSelection(page("<p>A</p><p></p><p>B</p>"));
  t.after(() => app.dom.window.close());
  const first = app.root.querySelector("p").firstChild;
  const last = app.root.querySelector("p:last-child").firstChild;
  const range = app.root.ownerDocument.createRange();
  range.setStart(first, 0);
  range.setEnd(last, last.length);
  assert.equal(app.selection.getClipboardContent(range).text, "A\n\nB");
});

test("reemplazar un rango hacia atrás inserta texto literal una vez y conserva ambos bordes", (t) => {
  const app = createSelection(page("<p>antes <strong>rango seleccionado</strong> después</p>") + page("<p>continuación</p>"));
  t.after(() => app.dom.window.close());
  const start = app.root.querySelector("strong").firstChild;
  const end = app.root.querySelector(".hwe-page:last-child p").firstChild;
  const range = app.root.ownerDocument.createRange();
  range.setStart(start, 6);
  range.setEnd(end, 6);
  const result = app.selection.replaceSelection(range, "<i>texto</i>");
  assert.ok(result);
  assert.equal(app.root.querySelector("strong").textContent, "rango <i>texto</i>");
  assert.equal(app.root.querySelectorAll("i").length, 0);
  assert.equal(app.root.textContent.includes("<i>texto</i>"), true);
  assert.equal(app.root.textContent.includes("después"), false);
  assert.equal(app.root.textContent.includes("continuación"), false);
});
