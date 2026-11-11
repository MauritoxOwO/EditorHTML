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
const { TextBlockSplitter } = require("../EditorHTML/pagination/TextBlockSplitter.ts");

for (const pieces of [["start", " one", " two"], ["start", " ", "one", " ", "two"]]) {
  test(`repaginar nodos de texto fragmentados produce una sola continuación: ${JSON.stringify(pieces)}`, (t) => {
    const dom = new JSDOM('<section class="hwe-page"><div class="hwe-page-inner"><p></p></div></section><div id="next"></div>');
    t.after(() => dom.window.close());
    for (const name of ["window", "document", "Node", "HTMLElement"]) global[name] = dom.window[name];
    global.getComputedStyle = dom.window.getComputedStyle.bind(dom.window);
    const page = document.querySelector(".hwe-page");
    const inner = page.firstElementChild;
    const block = inner.firstElementChild;
    const target = document.querySelector("#next");
    pieces.forEach((piece) => block.appendChild(document.createTextNode(piece)));
    // A deterministic page with room for five characters; JSDOM has no layout.
    inner.getBoundingClientRect = () => ({ bottom: 50 });
    block.getBoundingClientRect = () => ({ bottom: block.textContent.length * 10 });

    assert.equal(new TextBlockSplitter().splitTextBlock(block, target, page), true);
    assert.equal(block.textContent, "start", "the source must fit before pagination advances");
    assert.equal(target.children.length, 1, "words must share the same overflow paragraph");
    assert.equal(target.textContent, " one two");
    assert.equal(block.textContent + target.textContent, pieces.join(""));
  });
}
