const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");
const ts = require("typescript");
const { JSDOM } = require("jsdom");

function createApp(markup) {
  const dom = new JSDOM("<!doctype html><html><body><main class='hwe-workspace'></main></body></html>");
  const { window } = dom;
  const filename = path.resolve(__dirname, "../EditorHTML/selection/DocumentSelection.ts");
  const { outputText } = ts.transpileModule(fs.readFileSync(filename, "utf8"), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
    fileName: filename,
  });
  const exports = {};
  vm.runInNewContext(outputText, {
    exports,
    window,
    document: window.document,
    Node: window.Node,
    NodeFilter: window.NodeFilter,
    Range: window.Range,
    HTMLElement: window.HTMLElement,
  }, { filename });
  const root = window.document.querySelector("main");
  root.innerHTML = markup;
  const selection = new exports.DocumentSelection(() => root);
  const text = (selector) => root.querySelector(selector).firstChild;
  const range = (start, startOffset, end, endOffset) => {
    const result = window.document.createRange();
    result.setStart(start, startOffset);
    result.setEnd(end, endOffset);
    return result;
  };
  const select = (anchor, anchorOffset, focus, focusOffset) => {
    const nativeSelection = window.getSelection();
    nativeSelection.removeAllRanges();
    nativeSelection.setBaseAndExtent(anchor, anchorOffset, focus, focusOffset);
  };
  return { dom, window, root, selection, text, range, select };
}

const page = (content) => `<section class="hwe-page"><div class="hwe-page-inner" contenteditable="true">${content}</div></section>`;

test("una selección local en 261 páginas no crea rangos ni recorre texto de las páginas ajenas", (t) => {
  const app = createApp(Array.from({ length: 261 }, (_, index) => page(`<p id="p${index}">Texto ${index}</p>`)).join(""));
  t.after(() => app.dom.window.close());
  const selected = app.range(app.text("#p130"), 0, app.text("#p130"), 5);
  const createRange = app.window.document.createRange.bind(app.window.document);
  let ranges = 0;
  app.window.document.createRange = () => { ranges++; return createRange(); };
  const createWalker = app.window.document.createTreeWalker.bind(app.window.document);
  const walkedPages = [];
  app.window.document.createTreeWalker = (root, ...args) => {
    walkedPages.push(root.closest(".hwe-page"));
    return createWalker(root, ...args);
  };
  const paragraphs = app.selection.getSelectedParagraphs(selected);
  const targets = app.selection.getSelectedTextTargets(selected);
  assert.equal(paragraphs.length, 1);
  assert.equal(paragraphs[0].id, "p130");
  assert.equal(targets.length, 1);
  assert.ok(ranges <= 2, `rangos temporales: ${ranges}`);
  assert.equal(walkedPages.length, 1);
  assert.equal(walkedPages[0], app.root.querySelector("#p130").closest(".hwe-page"));
});

test("lee selección hacia delante y hacia atrás y enumera las páginas afectadas", (t) => {
  const app = createApp(page("<p>Alfa</p>") + page("<p>Beta</p>") + page("<p>Gamma</p>"));
  t.after(() => app.dom.window.close());
  const alfa = app.text(".hwe-page:nth-child(1) p");
  const beta = app.text(".hwe-page:nth-child(2) p");
  app.select(alfa, 2, beta, 2);
  let snapshot = app.selection.readSelection();
  assert.equal(snapshot.backward, false);
  assert.equal(snapshot.range.toString(), "faBe");
  assert.equal(snapshot.pages.length, 2);

  app.select(beta, 2, alfa, 2);
  snapshot = app.selection.readSelection();
  assert.equal(snapshot.backward, true);
  assert.equal(snapshot.range.toString(), "faBe");
  assert.equal(JSON.stringify(snapshot.pages.map((item) => [...app.root.children].indexOf(item))), "[0,1]");
});

test("devuelve offsets de texto exactos y excluye cabeceras y contenido protegido", (t) => {
  const app = createApp(page(
    "<p data-hwe-text-flow-id='flow'><strong>Inicio</strong> y <span>fin</span></p>" +
    "<div contenteditable='false'>candado</div><div data-hwe-api-header>cabecera</div>"
  ) + page("<p data-hwe-text-flow-id='flow'><em>continúa</em></p>"));
  t.after(() => app.dom.window.close());
  const start = app.text("strong");
  const end = app.text("em");
  const range = app.range(start, 2, end, 4);
  const targets = app.selection.getSelectedTextTargets(range);
  assert.equal(JSON.stringify(targets.map((target) => [target.node.data, target.start, target.end])), JSON.stringify([
    ["Inicio", 2, 6], [" y ", 0, 3], ["fin", 0, 3], ["continúa", 0, 4],
  ]));
  assert.equal(targets.map((target) => target.node.data.slice(target.start, target.end)).join(""), "icio y fincont");
});

test("identifica objetos y páginas como límites de un único flujo de texto", (t) => {
  const app = createApp(
    page("<p data-hwe-text-flow-id='flow'>Inicio<img src='x'>medio</p>") +
    page("<p data-hwe-text-flow-id='flow'>fin</p>")
  );
  t.after(() => app.dom.window.close());
  const start = app.text(".hwe-page:nth-child(1) p");
  const end = app.text(".hwe-page:nth-child(2) p");
  const range = app.range(start, 1, end, 2);
  assert.equal(app.selection.isSafeSinglePageRange(range), false);
  assert.equal(app.selection.getSelectedTextTargets(range).map((target) => target.node.data.slice(target.start, target.end)).join(""), "niciomediofi");
});

test("reemplaza entre páginas respetando el prefijo, el sufijo y los nodos protegidos", (t) => {
  const app = createApp(
    page("<p>prefijo <strong>primera porción</strong> cola</p>") +
    page("<p><span contenteditable='false'>intacto</span> medio párrafo</p>") +
    page("<p>final inicio y sufijo</p><p>ajeno</p>")
  );
  t.after(() => app.dom.window.close());
  const start = app.text("strong");
  const end = app.root.querySelector(".hwe-page:nth-child(3) p").firstChild;
  const range = app.range(start, 8, end, 12);
  const affected = app.selection.replaceSelection(range, "CAMBIO");
  assert.ok(affected);
  assert.equal(app.root.querySelector("strong").textContent, "primera CAMBIO");
  assert.equal(app.root.querySelector("[contenteditable='false']").textContent, "intacto");
  assert.equal(app.root.querySelectorAll(".hwe-page-inner")[0].textContent, "prefijo primera CAMBIO");
  assert.equal(app.root.querySelectorAll(".hwe-page-inner")[1].textContent, "intacto");
  assert.equal(app.root.querySelectorAll(".hwe-page-inner")[2].textContent, " y sufijoajeno");
});

test("normalizar un caret en el workspace o en el borde de página no selecciona toda la página", (t) => {
  const app = createApp(page("<p>Uno</p><p>Dos</p>") + page("<p>Tres</p>"));
  t.after(() => app.dom.window.close());
  const inner = app.root.querySelector(".hwe-page-inner");
  const outer = app.root.querySelector(".hwe-page");
  for (const [container, offset] of [[app.root, 0], [outer, 0], [inner, 0]]) {
    const collapsed = app.range(container, offset, container, offset);
    const normalized = app.selection.normalizeRange(collapsed);
    assert.ok(normalized);
    assert.equal(normalized.collapsed, true);
    assert.equal(normalized.toString(), "");
  }
});

test("un rango dentro de contenido protegido no se expande a la página ni se vuelve editable", (t) => {
  const app = createApp(page("<p>antes <span contenteditable='false'>Candado</span> después</p><p>Otro párrafo</p>"));
  t.after(() => app.dom.window.close());
  const protectedText = app.text("[contenteditable='false']");
  const source = app.range(protectedText, 1, protectedText, 4);
  const normalized = app.selection.normalizeRange(source);
  assert.ok(!normalized || normalized.toString() !== "antes Candado despuésOtro párrafo");
  assert.equal(normalized ? app.selection.isEditableRange(normalized) : false, false);
});

test("un Ctrl+A del host se normaliza en un rango hacia delante sobre todo el documento", (t) => {
  const app = createApp(page("<p>Uno</p>") + page("<p>Dos</p>"));
  t.after(() => app.dom.window.close());
  const hostRange = app.range(app.root, 0, app.root, app.root.childNodes.length);
  const normalized = app.selection.normalizeRange(hostRange);
  assert.ok(normalized);
  assert.equal(normalized.toString(), "UnoDos");
  app.select(normalized.startContainer, normalized.startOffset, normalized.endContainer, normalized.endOffset);
  const snapshot = app.selection.readSelection();
  assert.equal(snapshot.backward, false);
  assert.equal(snapshot.pages.length, 2);
});

test("reemplazar texto dentro de una celda conserva la forma de la tabla y sus vecinos", (t) => {
  const app = createApp(page("<table><tbody><tr><td>uno</td><td>vecina</td></tr><tr><td>otra fila</td><td>final</td></tr></tbody></table>"));
  t.after(() => app.dom.window.close());
  const first = app.root.querySelector("td").firstChild;
  const last = app.root.querySelectorAll("td")[3].firstChild;
  const range = app.range(first, 1, last, 2);
  assert.ok(app.selection.replaceSelection(range, "X"));
  assert.equal(app.root.querySelectorAll("table").length, 1);
  assert.equal(app.root.querySelectorAll("tr").length, 2);
  assert.deepEqual(Array.from(app.root.querySelectorAll("td"), (cell) => cell.textContent), ["uX", "", "", "nal"]);
  assert.equal(app.root.querySelector("table").textContent.includes("X"), true);
});

test("reemplazar una selección que atraviesa fragmentos de página conserva párrafos no seleccionados", (t) => {
  const app = createApp(
    page("<p>párrafo previo</p><p>antes<em> texto inicial</em></p>") +
    page("<p>fragmento seleccionado</p>") +
    page("<p>texto final<em> posterior</em></p><p>después</p>")
  );
  t.after(() => app.dom.window.close());
  const start = app.root.querySelector(".hwe-page:nth-child(1) em").firstChild;
  const end = app.root.querySelector(".hwe-page:nth-child(3) em").firstChild;
  const range = app.range(start, 2, end, 3);
  const affected = app.selection.replaceSelection(range, "cambio");
  assert.ok(affected);
  assert.equal(app.root.textContent.includes("párrafo previo"), true);
  assert.equal(app.root.textContent.includes("después"), true);
  assert.equal(app.root.textContent.includes("cambio"), true);
  assert.equal(app.root.textContent.includes("fragmento seleccionado"), false);
  assert.equal(app.root.textContent.includes("texto final"), false);
  assert.equal(app.root.textContent.includes("terior"), true);
});

test("Ctrl+A que atraviesa una tabla también borra el texto anterior y posterior", (t) => {
  const app = createApp(page("<p>antes</p><table><tbody><tr><td>celda</td></tr></tbody></table><p>después</p>"));
  t.after(() => app.dom.window.close());
  const all = app.range(app.root, 0, app.root, app.root.childNodes.length);
  assert.ok(app.selection.replaceSelection(all, ""));
  assert.equal(app.root.textContent.includes("antes"), false);
  assert.equal(app.root.textContent.includes("celda"), false);
  assert.equal(app.root.textContent.includes("después"), false);
});

test("al borrar alrededor de una celda protegida conserva su contenido y la tabla", (t) => {
  const app = createApp(page(
    "<table><tbody><tr><td>editable</td><td contenteditable='false'>protegido</td></tr></tbody></table>"
  ));
  t.after(() => app.dom.window.close());
  const all = app.range(app.root, 0, app.root, app.root.childNodes.length);
  assert.ok(app.selection.replaceSelection(all, ""));
  assert.equal(app.root.querySelectorAll("table").length, 1);
  assert.equal(app.root.querySelector("[contenteditable='false']").textContent, "protegido");
  assert.equal(app.root.textContent.includes("editable"), false);
});

test("reemplazar texto de la celda exterior conserva la tabla anidada y el resto de la cuadrícula", (t) => {
  const app = createApp(page(
    "<table id='outer'><tbody><tr><td>antes <table id='nested'><tbody><tr><td>interno</td></tr></tbody></table> después</td><td>vecina</td></tr></tbody></table>"
  ));
  t.after(() => app.dom.window.close());
  const cellText = app.root.querySelector("#outer > tbody > tr > td").firstChild;
  const range = app.range(cellText, 0, cellText, 5);
  assert.ok(app.selection.replaceSelection(range, "cambio"));
  assert.ok(app.root.querySelector("#nested"));
  assert.equal(app.root.querySelectorAll("#outer > tbody > tr").length, 1);
  assert.equal(app.root.querySelectorAll("#outer > tbody > tr > td").length, 2);
  assert.equal(app.root.querySelector("#nested td").textContent, "interno");
  assert.equal(app.root.querySelectorAll("#outer > tbody > tr > td")[1].textContent, "vecina");
});
