const assert = require("node:assert/strict");
const fs = require("node:fs");
const test = require("node:test");
const ts = require("typescript");
const { JSDOM } = require("jsdom");

require.extensions[".ts"] = (module, filename) => {
  const { outputText } = ts.transpileModule(fs.readFileSync(filename, "utf8"), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 }, fileName: filename,
  });
  module._compile(outputText, filename);
};
const { esEstiloTitulo, KeepTogetherController } = require("../EditorHTML/pagination/KeepTogetherController.ts");
const { DocumentSerializer } = require("../EditorHTML/services/DocumentSerializer.ts");
const { DEFAULT_PAGE_SETUP } = require("../EditorHTML/pagination/PageGeometry.ts");

function setup(t, html) {
  const dom = new JSDOM(`<main>${html}</main>`);
  t.after(() => dom.window.close());
  for (const name of ["window", "document", "Node", "HTMLElement", "DOMParser"]) global[name] = dom.window[name];
  global.getComputedStyle = dom.window.getComputedStyle.bind(dom.window);
  const root = document.querySelector("main");
  const controller = new KeepTogetherController();
  const group = () => {
    const nodes = controller.groupFlowNodes([...root.childNodes]);
    root.replaceChildren(...nodes);
    return nodes;
  };
  return { root, controller, group };
}

for (const name of ["tit", "art-tit", "art-Tit", "art-TIT", "art_tit", "tit-general", "art-tit-2", "  art-TIT  "]) {
  test(`reconoce el estilo de título ${JSON.stringify(name)}`, () => assert.equal(esEstiloTitulo(name), true));
}
for (const name of ["", "art-texto", "art-titular", "arttit", "art-tit2"]) {
  test(`no confunde ${JSON.stringify(name)} con un título`, () => assert.equal(esEstiloTitulo(name), false));
}

test("agrupa el título aplicado y el párrafo inmediato sin incluir los demás", t => {
  const app = setup(t, '<p data-hwe-paragraph-style="art-TIT">Título</p><p>Cuerpo</p><p>Después</p>');
  const nodes = app.group();
  assert.equal(nodes.length, 2);
  assert.equal(app.controller.isTitleGroup(nodes[0]), true);
  assert.equal(nodes[0].children.length, 2);
  assert.equal(nodes[0].textContent, "TítuloCuerpo");
  assert.equal(nodes[1].textContent, "Después");
});

test("lee el estilo en el párrafo dentro de su columna de texto", t => {
  const app = setup(t, '<div class="hwe-text-flow-column"><p class="art-tit">Título</p></div><div class="hwe-text-flow-column"><p>Cuerpo</p></div>');
  assert.equal(app.group().length, 1);
  assert.equal(app.root.firstElementChild.children.length, 2);
});

test("prioriza el estilo aplicado frente a una clase de título residual", t => {
  const app = setup(t, '<p class="art-tit" data-hwe-paragraph-style="texto-general">Texto</p><p>Cuerpo</p>');
  assert.equal(app.group().length, 2);
});

test("agrupa títulos consecutivos y el primer párrafo, ignorando espacios del HTML", t => {
  const app = setup(t, '<p class="cap-tit">Capítulo</p>\n<!-- formato --><p class="art_TIT">Artículo</p>\n<p>Cuerpo</p><p>Después</p>');
  const nodes = app.group();
  assert.equal(nodes.length, 2);
  assert.equal(nodes[0].querySelectorAll("p").length, 3);
  assert.equal(nodes[1].textContent, "Después");
});

for (const barrier of [
  '<p><br></p>', '<p data-hwe-user-blank="true"><br></p>',
  '<div class="hwe-text-flow-column"><p><br></p></div>',
  '<div data-hwe-manual-page-break="true"></div>', '<p hidden>Oculto</p>',
  '<p contenteditable="false">Protegido</p>', '<table><tbody><tr><td>Tabla</td></tr></tbody></table>',
]) {
  test(`no atraviesa la barrera ${barrier}`, t => {
    const app = setup(t, `<p class="art-tit">Título</p>${barrier}<p>Cuerpo</p>`);
    app.group();
    assert.equal(app.root.querySelector("[data-hwe-title-group]"), null);
  });
}

test("no limita los títulos a 240 caracteres ni agrupa títulos sin párrafo posterior", t => {
  const app = setup(t, `<p class="art-tit">${"Título ".repeat(60)}</p><p>Cuerpo</p>`);
  assert.equal(app.group().length, 1);
  app.root.innerHTML = '<p class="art-tit">Último título</p>';
  assert.equal(app.group()[0].hasAttribute("data-hwe-title-group"), false);
});

test("conserva las agrupaciones anteriores de imagen y tabla", t => {
  const app = setup(t, '<p>Pie</p><img src="a.png"><p>Tabla siguiente</p><table><tr><td>A</td></tr></table>');
  const nodes = app.group();
  assert.equal(app.controller.isKeepTogetherGroup(nodes[0]), true);
  assert.equal(app.controller.isTitleGroup(nodes[0]), false);
  assert.equal(nodes[1].getAttribute("data-hwe-keep-with-next"), "true");
});

test("preserva una línea vacía importada después del título antes de filtrar el flujo", t => {
  const app = setup(t, '<div class="hwe-text-flow-column"><p class="art-tit">Título</p></div>\n<div class="hwe-text-flow-column"><p id="separator"><br></p></div><p>Cuerpo</p><p id="unrelated"><br></p>');
  app.controller.preserveTitleBlankBarriers(app.root);
  assert.equal(app.root.querySelector("#separator").getAttribute("data-hwe-user-blank"), "true");
  assert.equal(app.root.querySelector("#unrelated").hasAttribute("data-hwe-user-blank"), false);
  app.group();
  assert.equal(app.root.querySelector("[data-hwe-title-group]"), null);
});

test("guardar y preparar HTML para PDF no conserva el contenedor temporal ni cambia dirty", t => {
  const app = setup(t, '<p class="art-tit" data-hwe-paragraph-style="art-tit">Título</p><p>Cuerpo</p>');
  const serializer = new DocumentSerializer();
  const original = serializer.normalizeHtmlForDirtyCheck(app.root.innerHTML);
  app.group();
  assert.equal(serializer.normalizeHtmlForDirtyCheck(app.root.innerHTML), original);
  const page = document.createElement("section");page.className = "hwe-page";
  const inner = document.createElement("div");inner.className = "hwe-page-inner";
  inner.append(...app.root.childNodes);page.append(inner);app.root.append(page);
  const saved = serializer.collectHtml(app.root, [page], DEFAULT_PAGE_SETUP, "");
  const printed = serializer.collectPdfHtml(app.root, [page], DEFAULT_PAGE_SETUP, "");
  assert.ok(!saved.includes("data-hwe-title-group"));
  assert.ok(!printed.includes("data-hwe-title-group"));
  assert.ok(saved.includes('data-hwe-paragraph-style="art-tit"'));
});
