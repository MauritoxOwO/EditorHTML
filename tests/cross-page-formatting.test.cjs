const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const test = require("node:test");
const ts = require("typescript");
const { JSDOM } = require("jsdom");

function setup(t, pageHtml) {
  const dom = new JSDOM("<!doctype html><main class='pcf-html-editor-root' contenteditable='true'></main>", { pretendToBeVisual: true });
  const { window } = dom;
  if (!window.Range.prototype.getClientRects) window.Range.prototype.getClientRects = () => [];
  const root = window.document.querySelector("main");
  root.innerHTML = pageHtml.map((html, index) => `<section id='page-${index}' class='hwe-page'><div class='hwe-page-inner'>${html}</div></section>`).join("");
  const cache = new Map();
  function load(relativePath) {
    const filename = path.resolve(__dirname, relativePath);
    if (cache.has(filename)) return cache.get(filename);
    const { outputText } = ts.transpileModule(fs.readFileSync(filename, "utf8"), {
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
    });
    const exports = {};
    cache.set(filename, exports);
    vm.runInNewContext(outputText, { exports, window, document: window.document, Node: window.Node,
      NodeFilter: window.NodeFilter, Range: window.Range, HTMLElement: window.HTMLElement,
      Element: window.Element,
      require: (name) => name === "../dom/EditorCssScope"
        ? { EDITOR_ROOT_CLASS: "pcf-html-editor-root" }
        : load(path.relative(__dirname, path.resolve(path.dirname(filename), name + ".ts"))),
    }, { filename });
    return exports;
  }
  const exports = load("../EditorHTML/controllers/TextSelectionFormatter.ts");
  const resolver = load("../EditorHTML/controllers/SelectionFormattingResolver.ts");
  const selectionExports = load("../EditorHTML/selection/DocumentSelection.ts");
  const documentSelection = new selectionExports.DocumentSelection(() => root);
  t.after(() => window.close());
  return { window, root, documentSelection,
    formatter: new exports.TextSelectionFormatter(() => root, documentSelection),
    SelectionFormattingResolver: resolver.SelectionFormattingResolver };
}

function select(window, start, startOffset, end, endOffset, backwards = false) {
  const selection = window.getSelection();
  selection.removeAllRanges();
  if (backwards) selection.setBaseAndExtent(end, endOffset, start, startOffset);
  else {
    const range = window.document.createRange();
    range.setStart(start, startOffset);
    range.setEnd(end, endOffset);
    selection.addRange(range);
  }
}

test("formats partial backward selection across pages, excludes protected text, and toggles mixed bold on", (t) => {
  const { window, root, formatter } = setup(t, [
    "<p><span id='a'>alpha</span></p>",
    "<p><span id='b'>bravo</span><span contenteditable='false'>LOCK</span></p>",
    "<p><span id='c'>charlie</span></p>",
  ]);
  const a = root.querySelector("#a").firstChild;
  const b = root.querySelector("#b").firstChild;
  const c = root.querySelector("#c").firstChild;
  select(window, a, 2, c, 3, true);
  formatter.applyCommand("bold");
  assert.equal(root.querySelector("#a").textContent, "alpha");
  assert.equal(root.querySelector("#b").textContent, "bravo");
  assert.equal(root.querySelector("[contenteditable=false]").textContent, "LOCK");
  const spans = Array.from(root.querySelectorAll("span[style*='font-weight']"));
  assert.equal(spans.length, 3);
  assert.equal(spans.map((span) => span.textContent).join(""), "phabravoch a".replaceAll(" ", ""));
  assert.equal(window.getSelection().isCollapsed, false);
  assert.equal(window.getSelection().toString(), "phabravoLOCKcha");
  assert.equal(window.getSelection().anchorNode.parentElement.closest(".hwe-page").id, "page-2");
  assert.equal(window.getSelection().focusNode.parentElement.closest(".hwe-page").id, "page-0");
});

test("host-level select-all is normalized to page contents before formatting", (t) => {
  const { window, root, formatter } = setup(t, [
    "<p>first</p><div contenteditable='false'>locked</div>",
    "<p>last</p>",
  ]);
  const selection = window.getSelection();
  const range = window.document.createRange();
  range.selectNodeContents(root);
  selection.removeAllRanges();
  selection.addRange(range);
  formatter.applyCommand("bold");
  assert.equal(root.querySelectorAll("span[style*='font-weight: bold']").length, 2);
  assert.equal(root.querySelector("[contenteditable='false'] span"), null);
  assert.equal(selection.toString(), "firstlockedlast");
});

test("mixed bold toggle removes formatting only when every selected fragment is already bold", (t) => {
  const { window, root, formatter } = setup(t, ["<p><b id='a'>bold</b><span id='b'>plain</span></p>"]);
  const start = root.querySelector("#a").firstChild;
  const end = root.querySelector("#b").firstChild;
  select(window, start, 0, end, end.length);
  formatter.applyCommand("bold");
  assert.equal(root.querySelectorAll("span[style*='font-weight: bold']").length, 2);
  const selectedBold = root.querySelector("span[style*='font-weight: bold']").firstChild;
  const lastBold = root.querySelectorAll("span[style*='font-weight: bold']")[1].firstChild;
  select(window, selectedBold, 0, lastBold, lastBold.length);
  formatter.applyCommand("bold");
  assert.equal(root.querySelectorAll("span[style*='font-weight: normal']").length, 2);
});

test("font and highlight apply across three page fragments and alignment includes linked flow continuations", (t) => {
  const { window, root, formatter } = setup(t, [
    "<p data-hwe-text-flow-id='flow'>one</p>",
    "<p data-hwe-text-flow-id='flow'>two</p>",
    "<p data-hwe-text-flow-id='flow'>three</p>",
  ]);
  const first = root.querySelector("p").firstChild;
  const last = root.querySelectorAll("p")[2].firstChild;
  select(window, first, 1, last, 2);
  formatter.applyCommand("fontName", "Georgia");
  assert.equal(root.querySelectorAll("span[style*='font-family']").length, 3);
  formatter.applyCommand("hiliteColor", "#ffff00");
  assert.equal(root.querySelectorAll("span[style*='background-color']").length, 3);
  formatter.applyCommand("justifyCenter");
  assert.deepEqual(Array.from(root.querySelectorAll("p")).map((p) => p.style.textAlign), ["center", "center", "center"]);
});

test("case conversion handles length-changing Unicode while retaining a live selection", (t) => {
  const { window, root, formatter } = setup(t, ["<p><span id='s'>aßz</span></p>"]);
  const text = root.querySelector("#s").firstChild;
  select(window, text, 1, text, 2, true);
  formatter.transformSelectedTextCase("uppercase");
  assert.equal(root.querySelector("#s").textContent, "aSSz");
  assert.equal(window.getSelection().toString(), "SS");
  assert.equal(window.getSelection().anchorOffset, 3);
  assert.equal(window.getSelection().focusOffset, 1);
});

test("removing underline splits only the selected part from semantic underline while keeping sibling formatting", (t) => {
  const { window, root, formatter } = setup(t, ["<p><u><b id='b'>abcdef</b></u></p>"]);
  const text = root.querySelector("#b").firstChild;
  select(window, text, 2, text, 4);
  formatter.applyCommand("underline");
  const paragraph = root.querySelector("p");
  assert.equal(paragraph.textContent, "abcdef");
  assert.equal(paragraph.querySelectorAll("u").length, 2);
  const selected = Array.from(paragraph.querySelectorAll("span")).find((span) => span.textContent === "cd");
  assert.ok(selected);
  assert.equal(selected.querySelector("b")?.textContent, "cd");
});

test("underline removal preserves mixed inline sibling order, nested underline, and block structure", (t) => {
  const { window, root, formatter } = setup(t, [
    "<p><u><i>A</i><b>B</b><em>CD</em>E</u></p>",
    "<p><u><u>abcdef</u></u></p>",
    "<style>.underlined { text-decoration: underline; }</style><p class='underlined'><i>inline</i></p>",
  ]);
  const first = root.querySelector("em").firstChild;
  select(window, first, 1, first, 2);
  formatter.applyCommand("underline");
  assert.equal(root.querySelectorAll(".hwe-page-inner p")[0].textContent, "ABCDE");
  assert.equal(root.querySelectorAll(".hwe-page-inner p")[0].children[1].textContent, "D");
  assert.equal(window.getSelection().toString(), "D");

  const nestedText = root.querySelectorAll(".hwe-page-inner")[1].querySelector("u u").firstChild;
  select(window, nestedText, 2, nestedText, 4);
  formatter.applyCommand("underline");
  assert.equal(root.querySelectorAll(".hwe-page-inner")[1].textContent, "abcdef");
  const middle = Array.from(root.querySelectorAll(".hwe-page-inner")[1].querySelectorAll("span"))
    .find((span) => span.textContent === "cd");
  assert.ok(middle);
  assert.equal(middle.closest("u"), null);

  const styledText = root.querySelectorAll(".hwe-page-inner")[2].querySelector("i").firstChild;
  select(window, styledText, 1, styledText, 4);
  formatter.applyCommand("underline");
  const paragraphs = Array.from(root.querySelectorAll(".hwe-page-inner")[2].querySelectorAll("p"));
  assert.equal(paragraphs.map((p) => p.textContent).join(""), "inline");
  assert.equal(paragraphs.length, 1);
  assert.equal(paragraphs[0].style.textDecorationLine, "none");
  assert.equal(window.getSelection().toString(), "nli");
});

test("alignment across a selected logical flow never mutates protected continuation blocks", (t) => {
  const { window, root, formatter } = setup(t, [
    "<p data-hwe-text-flow-id='f'>one</p>",
    "<div contenteditable='false'><p data-hwe-text-flow-id='f'>locked</p></div>",
    "<p data-hwe-text-flow-id='f'>three</p>",
  ]);
  const first = root.querySelector("p").firstChild;
  select(window, first, 0, first, 2);
  formatter.applyCommand("justifyRight");
  assert.equal(root.querySelector("p[data-hwe-text-flow-id='f']").style.textAlign, "right");
  assert.equal(root.querySelector("[contenteditable='false'] p").style.textAlign, "");
  assert.equal(root.querySelectorAll("p[data-hwe-text-flow-id='f']")[2].style.textAlign, "right");
});

test("alignment excludes a paragraph that only touches the final offset-zero boundary and styles an empty caret paragraph", (t) => {
  const { window, root, formatter } = setup(t, ["<p id='first'>text</p><p id='last'>next</p><p id='empty'><br></p>"]);
  const firstText = root.querySelector("#first").firstChild;
  const last = root.querySelector("#last");
  select(window, firstText, 0, last, 0);
  formatter.applyCommand("justifyRight");
  assert.equal(root.querySelector("#first").style.textAlign, "right");
  assert.equal(root.querySelector("#last").style.textAlign, "");
  const emptyBr = root.querySelector("#empty br");
  const caret = window.document.createRange();
  caret.setStartBefore(emptyBr);
  caret.collapse(true);
  window.getSelection().removeAllRanges();
  window.getSelection().addRange(caret);
  formatter.applyCommand("justifyCenter");
  assert.equal(root.querySelector("#empty").style.textAlign, "center");
});

test("font size toolbar dispatch preserves the selected text and reports command plus value", (t) => {
  const { window, root } = setup(t, ["<p><span id='s'>selected</span></p>"]);
  const filename = path.resolve(__dirname, "../EditorHTML/ui/Toolbar.ts");
  const { outputText } = ts.transpileModule(fs.readFileSync(filename, "utf8"), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
  });
  const exports = {};
  vm.runInNewContext(outputText, { exports, window, document: window.document,
    require: () => ({}) }, { filename });
  const calls = [];
  const toolbar = new exports.Toolbar({
    onCommand: (command, value) => { calls.push([command, value]); return false; },
    onApplyFontSize: (value) => calls.push(["legacy-size", value]),
  });
  root.appendChild(toolbar.build());
  const text = root.querySelector("#s").firstChild;
  select(window, text, 1, text, 6, true);
  const size = root.querySelector("select[title='Tamaño']");
  size.value = "14";
  size.dispatchEvent(new window.Event("change", { bubbles: true }));
  assert.deepEqual(calls, [["fontSize", "14pt"], ["legacy-size", "14pt"]]);
  assert.equal(window.getSelection().anchorNode, text);
  assert.equal(window.getSelection().focusNode, text);
  assert.ok(window.getSelection().anchorOffset > window.getSelection().focusOffset);
  toolbar.destroy();
});

test("selection resolver reports mixed inline states without choosing the start fragment", (t) => {
  const { window, root, documentSelection, SelectionFormattingResolver } = setup(t, ["<p><b id='bold' style='font-size:12pt'>bold</b><span id='plain' style='font-size:16pt'>plain</span></p>"]);
  const bold = root.querySelector("#bold").firstChild;
  const plain = root.querySelector("#plain").firstChild;
  select(window, bold, 0, plain, plain.length);
  const resolver = new SelectionFormattingResolver({ rootProvider: () => root, isParagraphStyleClass: () => false, documentSelection });
  const state = resolver.resolve();
  assert.equal(state.commandStates.bold, "mixed");
  assert.equal(state.fontSize, undefined);
});
