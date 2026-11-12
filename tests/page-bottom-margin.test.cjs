const assert = require("node:assert/strict");
const fs = require("node:fs");
const test = require("node:test");
const ts = require("typescript");
const { JSDOM } = require("jsdom");

require.extensions[".ts"] = (module, filename) => {
  const { outputText } = ts.transpileModule(fs.readFileSync(filename, "utf8"), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
    fileName: filename,
  });
  module._compile(outputText, filename);
};
const { getPageContentBounds } = require("../EditorHTML/pagination/PageContentGeometry.ts");
const { getContentLimitBottom, getContentHeight, pageOverflows } = require("../EditorHTML/pagination/PaginatorDom.ts");
const { EditorLayoutService } = require("../EditorHTML/services/EditorLayoutService.ts");
const { Paginator } = require("../EditorHTML/pagination/Paginator.ts");
const { getHweDebugReport } = require("../EditorHTML/debug/DebugLogger.ts");

const mm = 96 / 25.4;
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 0.001, `${actual} != ${expected}`);

function setup(t, scale, tag = "p") {
  const content = tag === "table" ? "<table><tbody><tr><td>Contenido</td></tr></tbody></table>" : "<p>Contenido</p>";
  const dom = new JSDOM(`<main><section class="hwe-page"><div class="hwe-page-inner">${content}</div></section></main>`);
  t.after(() => dom.window.close());
  for (const name of ["window", "document", "Node", "HTMLElement", "Range"]) global[name] = dom.window[name];
  global.getComputedStyle = dom.window.getComputedStyle.bind(dom.window);
  const root = document.querySelector("main");
  const page = root.firstElementChild;
  const inner = page.firstElementChild;
  inner.style.cssText = `box-sizing:border-box;height:${297 * mm}px;padding-top:${35 * mm}px;padding-bottom:${20 * mm}px`;
  Object.defineProperties(inner, {
    clientHeight: { value: Math.round(297 * mm) },
    offsetHeight: { value: Math.round(297 * mm) },
    scrollHeight: { value: Math.round(297 * mm) },
  });
  const top = 123;
  inner.getBoundingClientRect = () => ({ top, bottom: top + 297 * mm * scale, height: 297 * mm * scale });
  const child = inner.firstElementChild;
  let childBottom = top + 277 * mm * scale;
  child.getBoundingClientRect = () => ({ bottom: childBottom, top: 200, height: 100 });
  return { root, page, inner, child, setBottom: (value) => { childBottom = value; } };
}

for (const scale of [0.5, 1, 1.18, 2.5]) {
  test(`reserva los 20 mm completos y 242 mm útiles con zoom ${scale}`, (t) => {
    const { inner, page } = setup(t, scale);
    const bounds = getPageContentBounds(inner);
    near(bounds.bottom, 123 + 277 * mm * scale);
    near(bounds.top, 123 + 35 * mm * scale);
    near(bounds.scale, scale);
    near(getContentHeight(page), 242 * mm);
    near(getContentLimitBottom(inner), bounds.bottom);
  });

  for (const tag of ["p", "table"]) {
    test(`${tag}: detecta invasión del margen antes del borde de hoja con zoom ${scale}`, (t) => {
      const { root, page, inner, setBottom } = setup(t, scale, tag);
      const bounds = getPageContentBounds(inner);
      const layout = new EditorLayoutService();
      assert.equal(pageOverflows(page), false);
      assert.equal(layout.pageOverflows(page), false);
      setBottom(bounds.bottom + 5 * mm * scale);
      assert.equal(pageOverflows(page), true);
      assert.equal(layout.pageOverflows(page), true);
      const report = getHweDebugReport(root).pages[0];
      near(report.limitBottom, bounds.bottom);
      assert.equal(report.overflows, true);
    });
  }

  test(`evalúa bloques e imágenes en el mismo espacio útil con zoom ${scale}`, (t) => {
    const { inner, page } = setup(t, scale);
    const bounds = getPageContentBounds(inner);
    const paginator = Object.create(Paginator.prototype);
    const block = document.createElement("div");
    block.getBoundingClientRect = () => ({ height: 243 * mm * scale });
    assert.equal(paginator.elementFitsOnFreshPage(block, page), false);
    block.getBoundingClientRect = () => ({ height: 242 * mm * scale });
    assert.equal(paginator.elementFitsOnFreshPage(block, page), true);
    const image = document.createElement("img");
    inner.appendChild(image);
    image.getBoundingClientRect = () => ({ top: bounds.top });
    assert.equal(paginator.fitImagesToAvailableSpace(page), true);
    assert.equal(image.style.maxHeight, `${Math.floor(242 * mm)}px`);
  });
}

test("descuenta bordes y conserva la reserva de un margen inferior menor de 12 px", (t) => {
  const { inner } = setup(t, 1);
  inner.style.cssText = "height:200px;box-sizing:content-box;padding:10px 0 5px;border-top:2px solid;border-bottom:3px solid";
  inner.getBoundingClientRect = () => ({ top: 10, bottom: 450, height: 440 });
  const bounds = getPageContentBounds(inner);
  near(bounds.scale, 2);
  near(bounds.top, 34);
  near(bounds.bottom, 434);
  near(bounds.contentHeight, 200);
});

test("una tabla larga conserva la densidad al fragmentarse y la recalcula al borrar filas", (t) => {
  const { root, inner } = setup(t, 1);
  root.className = "pcf-html-editor-root";
  const rows = count => "<tr><td><p>Fila</p></td></tr>".repeat(count);
  inner.innerHTML = `<table class="hwe-word-table" data-hwe-table-flow-id="a"><tbody>${rows(30)}</tbody></table>`;
  root.insertAdjacentHTML("beforeend", `<section class="hwe-page"><div class="hwe-page-inner"><table class="hwe-word-table" data-hwe-table-flow-id="a"><tbody>${rows(30)}</tbody></table></div></section>`);
  const layout = new EditorLayoutService();
  layout.applyOfficialTableWidths(root);
  const tables = [...root.querySelectorAll("table")];
  assert.ok(tables.every(table => table.classList.contains("hwe-table-compact")));
  tables[1].querySelector("tr").remove();
  layout.applyOfficialTableWidths(root);
  assert.ok(tables.every(table => !table.classList.contains("hwe-table-compact")));
  // Another editor's table must not contribute rows, even with the same flow id.
  root.insertAdjacentHTML("afterend", `<main class="pcf-html-editor-root"><table data-hwe-table-flow-id="a"><tbody>${rows(100)}</tbody></table></main>`);
  layout.applyOfficialTableWidths(root);
  assert.ok(tables.every(table => !table.classList.contains("hwe-table-compact")));
});

test("el cálculo por páginas conserva la densidad global y se actualiza tras borrar filas", t => {
  const { root, inner } = setup(t, 1);
  root.className = "pcf-html-editor-root";
  const table = count => `<table class="hwe-word-table" data-hwe-table-flow-id="a"><tbody>${'<tr><td>Fila</td></tr>'.repeat(count)}</tbody></table>`;
  inner.innerHTML = table(30);
  root.insertAdjacentHTML('beforeend', `<section class="hwe-page"><div class="hwe-page-inner">${table(30)}</div></section>`);
  const pages = [...root.querySelectorAll('.hwe-page')];
  const tables = [...root.querySelectorAll('table')];
  const layout = new EditorLayoutService();
  layout.applyOfficialTableWidthsToPages(pages);
  assert.ok(tables.every(table => table.classList.contains('hwe-table-compact')));
  tables[1].querySelector('tr').remove();
  layout.applyOfficialTableWidthsToPages(pages);
  assert.ok(tables.every(table => !table.classList.contains('hwe-table-compact')));
});

test("reaplicar el layout no modifica párrafos intactos y retira el layout al moverlos a una celda", t => {
  const { inner } = setup(t, 1);
  const layout = new EditorLayoutService();
  layout.applyTextFlowLayout(inner);
  const paragraph = inner.querySelector('p');
  const column = paragraph.parentElement;
  const observer = new window.MutationObserver(() => {});
  observer.observe(inner, {subtree:true,attributes:true,childList:true});
  layout.applyTextFlowLayout(inner);
  assert.deepEqual(observer.takeRecords(), [], 'sin invalidaciones de estilo innecesarias');
  assert.equal(paragraph.parentElement,column);
  inner.insertAdjacentHTML('beforeend','<table><tbody><tr><td></td></tr></tbody></table>');
  inner.querySelector('td').appendChild(paragraph);
  layout.applyTextFlowLayout(inner);
  assert.equal(paragraph.classList.contains('hwe-text-flow-block'),false);
  assert.equal(inner.querySelector('.hwe-text-flow-column'),null);
  observer.disconnect();
});
