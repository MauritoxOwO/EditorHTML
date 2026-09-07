const assert = require("node:assert/strict");
const fs = require("node:fs");
const test = require("node:test");
const ts = require("typescript");
const { JSDOM } = require("jsdom");

// Load the actual TypeScript sources without generating build files or a PCF host.
require.extensions[".ts"] = (module, filename) => {
  const { outputText } = ts.transpileModule(fs.readFileSync(filename, "utf8"), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
    fileName: filename,
  });
  module._compile(outputText, filename);
};
const dom = new JSDOM("<!doctype html><html><body></body></html>");
for (const name of ["window", "document", "Node", "HTMLElement", "DOMParser"]) {
  global[name] = dom.window[name];
}
global.getComputedStyle = dom.window.getComputedStyle.bind(dom.window);

const { DocumentSerializer } = require("../EditorHTML/services/DocumentSerializer.ts");
const { TextBlockSplitter } = require("../EditorHTML/pagination/TextBlockSplitter.ts");
const { TablePaginator } = require("../EditorHTML/pagination/TablePaginator.ts");
const { Paginator } = require("../EditorHTML/pagination/Paginator.ts");
const { DEFAULT_PAGE_SETUP } = require("../EditorHTML/pagination/PageGeometry.ts");
const serializer = new DocumentSerializer();
const normalize = (html) => serializer.normalizeHtmlForDirtyCheck(html);
const doc = (...pages) => '<div data-hwe-document="true" data-hwe-page-width="210mm">' +
  pages.map((html, index) => index === 0 ? html :
    `<div data-hwe-page-break="before" style="page-break-before:always">${html}</div>`
  ).join("\n") + "</div>";
const element = (html) => {
  const root = document.createElement("div");
  root.innerHTML = html;
  return root.firstElementChild;
};
const textFragment = (text, id = "text1", attrs = "") =>
  `<p data-hwe-text-flow-id="${id}" data-hwe-text-fragment="true" ${attrs}>${text}</p>`;
const rows = (...values) => values.map((value) => `<tr><td>${value}</td></tr>`).join("");
const table = (content, attrs = "", shell = "") =>
  `<table ${attrs}>${shell}<tbody>${content}</tbody></table>`;
const tableId = 'data-hwe-table-flow-id="table1" data-hwe-table-fragment="true"';
const header = "<thead><tr><th>Cabecera</th></tr></thead>";
const columns = '<colgroup><col style="width: 40%"><col style="width: 60%"></colgroup>';

function same(name, original, paginated) {
  test(name, () => assert.equal(normalize(paginated), normalize(original)));
}
function different(name, original, edited) {
  test(name, () => assert.notEqual(normalize(edited), normalize(original)));
}

same("automatic page boundaries do not change whole paragraphs", doc("<p>A</p><p>B</p>"), doc("<p>A</p>", "<p>B</p>"));
same("three automatic pages", doc("<p>A</p><p>B</p><p>C</p>"), doc("<p>A</p>", "<p>B</p>", "<p>C</p>"));
same("API and runtime headers do not affect content", doc("<p>A</p>"), doc(
  '<div data-hwe-api-header="true">Fecha nueva</div><div data-hwe-runtime-page-header="true">Logo</div><p>A</p>'
));
same("paragraph fragments are joined without adding spaces", doc("<p>uno dos tres</p>"),
  doc(textFragment("uno"), textFragment(" dos"), textFragment(" tres")));
same("technical identities can change", doc(textFragment("A", "old")), doc(textFragment("A", "new")));
same("empty selection class and runtime state are ignored", doc('<p>A</p>'),
  doc('<p class="hwe-text-flow-block hwe-image-selected" contenteditable="true" spellcheck="false">A</p>'));
same("attribute/class ordering is stable", doc('<p title="X" class="b a">A</p>'),
  doc('<p class="a b" title="X">A</p>'));
same("user blank metadata does not change the actual blank paragraph", doc("<p><br></p>"),
  doc('<p data-hwe-user-blank="true"><br></p>'));
same("keep-together wrapper is ignored", doc("<p>Título</p><img src='a.png'>"),
  doc('<div data-hwe-generated-wrapper="true" data-hwe-keep-together="true"><p>Título</p><img src="a.png"></div>'));
same("auto image max-height can vary", doc('<img src="a.png" style="width: 70%; max-height: 800px" data-hwe-auto-max-height="800px">'),
  doc('<img src="a.png" style="width: 70%; max-height: 600px" data-hwe-auto-max-height="600px">'));
same("table fragments and repeated headers join", doc(table(rows("A", "B"), 'data-hwe-repeat-header="true"', columns + header)),
  doc(table(rows("A"), tableId + ' data-hwe-repeat-header="true"', columns + header),
    table(rows("B"), tableId + ' data-hwe-repeat-header="true"', columns + header)));
same("table header without repeat is preserved once", doc(table(rows("A", "B"), "", header)),
  doc(table(rows("A"), tableId, header), table(rows("B"), tableId)));
same("nested table rows stay inside their own cell", doc(table(rows(table(rows("inside")), "B"))),
  doc(table(rows(table(rows("inside"))), tableId), table(rows("B"), tableId)));
same("automatic table density is not a user edit", doc(table(rows("A", "B"))),
  doc(table(rows("A"), tableId + ' class="hwe-table-compact hwe-long-word-table"'), table(rows("B"), tableId)));
same("nested styled containers join by identity", doc('<section style="color: red"><p>uno dos</p></section>'),
  doc('<section style="color: red" data-hwe-container-flow-id="c">' + textFragment("uno") + '</section>',
    '<section style="color: red" data-hwe-container-flow-id="c">' + textFragment(" dos") + '</section>'));
same("cloned containers can be compacted back onto one page", doc('<section style="color: red"><p>uno dos</p></section>'),
  doc('<section style="color: red" data-hwe-container-flow-id="c">' + textFragment("uno") + '</section>' +
    '<section style="color: red" data-hwe-container-flow-id="c">' + textFragment(" dos") + '</section>'));

different("plain text edits remain dirty", doc("<p>A</p>"), doc("<p>B</p>"));
different("text deletion remains dirty", doc("<p>AB</p>"), doc("<p>A</p>"));
different("formatting remains dirty", doc("<p>A</p>"), doc("<p><strong>A</strong></p>"));
different("fragment-specific formatting is not discarded", doc('<p style="color: red">AB</p>'),
  doc(textFragment("A", "x", 'style="color: red"'), textFragment("B", "x", 'style="color: blue"')));
different("independent paragraphs are not merged", doc("<p>AB</p>"), doc("<p>A</p>", "<p>B</p>"));
different("Enter inside a page is not hidden by copied flow IDs", doc(textFragment("AB")),
  doc(textFragment("A") + textFragment("B")));
different("different flow IDs cannot join", doc("<p>AB</p>"), doc(textFragment("A", "x"), textFragment("B", "y")));
different("identical-looking independent containers stay separate", doc('<section class="x"><p>A</p><p>B</p></section>'),
  doc('<section class="x"><p>A</p></section>', '<section class="x"><p>B</p></section>'));
different("blank paragraph addition remains dirty", doc("<p>A</p>"), doc("<p>A</p><p><br></p>"));
different("manual page break remains dirty", doc("<p>A</p><p>B</p>"),
  doc('<p>A</p><div data-hwe-manual-page-break="true"></div>', '<p>B</p>'));
different("manual break interrupts text flow", doc("<p>AB</p>"),
  doc(textFragment("A") + '<div data-hwe-manual-page-break="true"></div>', textFragment("B")));
different("preformatted whitespace remains significant", doc("<pre>A B</pre>"), doc("<pre>A  B</pre>"));
different("real whitespace at a page boundary is preserved", doc("<p>A</p><p>B</p>"), doc("<p>A</p>\n", "<p>B</p>"));
different("image source remains significant", doc('<img src="a.png">'), doc('<img src="b.png">'));
different("image width remains significant", doc('<img src="a.png" style="width: 50%">'), doc('<img src="a.png" style="width: 70%">'));
different("image alignment remains significant", doc('<img src="a.png" style="margin-left: auto">'), doc('<img src="a.png" style="margin-left: 0">'));
different("non-automatic max-height remains significant", doc('<img src="a.png" style="max-height: 800px">'),
  doc('<img src="a.png" style="max-height: 600px">'));
different("a changed repeated table header is not silently discarded", doc(table(rows("A", "B"), 'data-hwe-repeat-header="true"', header)),
  doc(table(rows("A"), tableId + ' data-hwe-repeat-header="true"', header),
    table(rows("B"), tableId + ' data-hwe-repeat-header="true"', header.replace("Cabecera", "Editada"))));
different("missing repeated table header remains dirty", doc(table(rows("A", "B"), 'data-hwe-repeat-header="true"', header)),
  doc(table(rows("A"), tableId + ' data-hwe-repeat-header="true"', header), table(rows("B"), tableId + ' data-hwe-repeat-header="true"')));
different("column resize remains dirty", doc(table(rows("A", "B"), "", columns)),
  doc(table(rows("A"), tableId, columns), table(rows("B"), tableId, columns.replace("40%", "30%"))));
different("cell edits remain dirty", doc(table(rows("A", "B"))), doc(table(rows("A"), tableId), table(rows("Changed"), tableId)));
different("cell styles remain dirty", doc(table(rows("A", "B"))),
  doc(table(rows("A"), tableId), table('<tr><td style="color: red">B</td></tr>', tableId)));
different("independent tables are not merged", doc(table(rows("A", "B"))), doc(table(rows("A")), table(rows("B"))));
different("page geometry is preserved", doc("<p>A</p>"), doc("<p>A</p>").replace("210mm", "297mm"));
different("document CSS is preserved", doc('<style>p{color:red}</style><p>A</p>'), doc('<style>p{color:blue}</style><p>A</p>'));

test("actual inline splitting restores nested styled runs", () => {
  const original = '<p><strong style="color: red"><em>uno dos tres cuatro</em></strong></p>';
  const source = element(original);
  const overflow = element('<p></p>');
  const splitter = new TextBlockSplitter();
  // Exercise the actual clone/move algorithm, without pretending jsdom has layout.
  source.setAttribute("data-hwe-text-flow-id", "real-split");
  overflow.setAttribute("data-hwe-text-flow-id", "real-split");
  splitter.moveLastInlinePiece(source, overflow);
  splitter.moveLastInlinePiece(source, overflow);
  assert.equal(normalize(doc(source.outerHTML, overflow.outerHTML)), normalize(doc(original)));
});

test("actual table splitter round-trip", () => {
  const original = table(rows("A", "B", "C"), 'data-hwe-repeat-header="true"', columns + header);
  const source = element(original);
  const target = document.createElement("div");
  const page = document.createElement("div");
  source.querySelectorAll("tbody > tr").forEach((row, index) => {
    row.getBoundingClientRect = () => ({ bottom: (index + 1) * 50 });
  });
  const split = new TablePaginator().splitTable(source, target, page, {
    getInner: () => target,
    getContentLimitBottom: () => 75,
  });
  assert.equal(split, true);
  assert.equal(normalize(doc(source.outerHTML, target.innerHTML)), normalize(doc(original)));
});

for (const [name, original] of [
  ["styled tbody", '<table><tbody style="color: red">' + rows("A", "B", "C") + '</tbody></table>'],
  ["tfoot", '<table><tbody>' + rows("A", "B") + '</tbody><tfoot style="font-weight: bold">' + rows("Total") + '</tfoot></table>'],
  ["separate tbody groups", '<table><tbody>' + rows("A", "B") + '</tbody><tbody>' + rows("C", "D") + '</tbody></table>'],
]) {
  test(`actual splitting preserves ${name}`, () => {
    const source = element(original);
    const target = document.createElement("div");
    Array.from(source.rows).forEach((row, index) => {
      row.getBoundingClientRect = () => ({ bottom: (index + 1) * 50 });
    });
    new TablePaginator().splitTable(source, target, document.createElement("div"), {
      getInner: () => target, getContentLimitBottom: () => 75,
    });
    assert.equal(normalize(doc(source.outerHTML, target.innerHTML)), normalize(doc(original)));
    const paginator = new Paginator(() => document.createElement("div"), () => {});
    paginator.mergeTableRows(source, target.firstElementChild);
    assert.equal(normalize(doc(source.outerHTML)), normalize(doc(original)), "after merging rows again");
    Array.from(source.rows).forEach((row, index) => {
      row.getBoundingClientRect = () => ({ bottom: (index + 1) * 50 });
    });
    new TablePaginator().splitTable(source, target, document.createElement("div"), {
      getInner: () => target, getContentLimitBottom: () => 75,
    });
    assert.equal(normalize(doc(source.outerHTML, target.innerHTML)), normalize(doc(original)), "after splitting again");
  });
}

test("untrusted row-group metadata cannot crash comparison", () => {
  const original = doc('<table><tbody data-hwe-generated-row-group="true"><tr data-hwe-row-group-origin="not JSON"><td>A</td></tr></tbody></table>');
  assert.doesNotThrow(() => normalize(original));
  assert.match(normalize(original), /<td>A<\/td>/);
});

test("paginator identifies cloned styled containers", () => {
  const original = '<section style="color: red"><p>A</p><p>B</p></section>';
  const source = element(original);
  const root = document.createElement("div");
  root.appendChild(source);
  const page = element('<div class="hwe-page"><div class="hwe-page-inner"></div></div>');
  const inner = page.firstElementChild;
  Object.defineProperty(inner, "scrollHeight", { get: () => source.children.length > 1 ? 200 : 0 });
  Object.defineProperty(inner, "clientHeight", { value: 100 });
  const target = document.createElement("div");
  const paginator = new Paginator(() => page, () => {});
  assert.equal(paginator.splitContainerPreservingShell(source, target, page), true);
  assert.equal(source.getAttribute("data-hwe-container-flow-id"), target.firstElementChild.getAttribute("data-hwe-container-flow-id"));
  assert.equal(normalize(doc(source.outerHTML, target.innerHTML)), normalize(doc(original)));
});

test("normalization leaves live editor and actual saved HTML untouched", () => {
  const root = document.createElement("div");
  const page = element('<div class="hwe-page"><div class="hwe-page-inner"><div data-hwe-api-header="true">Header</div><p>A</p></div></div>');
  root.appendChild(page);
  const liveBefore = root.innerHTML;
  const saved = serializer.collectHtml(root, [page], DEFAULT_PAGE_SETUP);
  normalize(saved);
  assert.equal(root.innerHTML, liveBefore);
  assert.equal(serializer.collectHtml(root, [page], DEFAULT_PAGE_SETUP), saved);
  assert.match(saved, /Header/);
});

test("normalization is idempotent", () => {
  const html = doc(textFragment("uno", "x"), textFragment(" dos", "x"));
  assert.equal(normalize(normalize(html)), normalize(html));
});
