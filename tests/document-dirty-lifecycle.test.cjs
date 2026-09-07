const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");
const ts = require("typescript");

// Ejecutar init, loadContent, save y el adaptador PCF reales. Se sustituyen
// únicamente el DOM, los servicios remotos y la construcción de la interfaz.
function loadSource(relativePath, dependencies) {
  const filename = path.join(__dirname, relativePath);
  const { outputText } = ts.transpileModule(fs.readFileSync(filename, "utf8"), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
    fileName: filename,
  });
  const exports = {};
  vm.runInNewContext(outputText, {
    exports,
    require: (name) => dependencies[name] || {},
    console: { error() {}, warn() {} },
    window: { setTimeout() {} },
  }, { filename });
  return exports;
}

function deferred() {
  let resolve, reject;
  const promise = new Promise((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

function setup(initialValue = true, options = {}) {
  let value = initialValue;
  let html = "";
  let editor, ready;
  const outputs = [];
  const download = deferred();
  const { EditorComponent } = loadSource("../EditorHTML/app/EditorComponent.ts", {
    "../pagination/Paginator": { Paginator: class {} },
    "../services/dataverse/fileApi": {
      fetchHtmlFromFileField: () => download.promise,
      saveHtmlToFileField: options.upload || (async () => {}),
    },
  });
  class HeadlessEditor {
    constructor(_container, _context, callbacks) {
      editor = Object.create(EditorComponent.prototype);
      Object.assign(editor, {
        options: callbacks,
        lastSavedHtml: null,
        baseUrl: "https://example.invalid",
        entityId: "test-record",
        root: { inert: false },
        buildShell() {},
        loadParagraphStyles: async () => {},
        loadDynamicHeader: async () => {},
        loadRuntimeHeaderLogo: async () => {},
        renderAndPaginate: async (content) => { html = content; },
        historyController: { reset() {} },
        ensureInitialPrintHtmlSaved: async () => {},
        setStatus() {},
        collectHtml: () => html,
        documentSerializer: { normalizeHtmlForDirtyCheck: (content) => content },
        toolbar: { getSaveButton: () => ({ disabled: false }) },
        refreshDynamicHeaderBeforeSave: async () => {},
        collectPrintHtml: () => html,
        savePrintHtml: async () => {},
      });
    }
    init() { ready = editor.init(); return ready; }
    resize() {}
    destroy() {}
  }
  const { EditorHTML2 } = loadSource("../EditorHTML/index.ts", {
    "./app/EditorComponent": { EditorComponent: HeadlessEditor },
  });
  const control = new EditorHTML2();
  const context = () => ({
    parameters: { documentDirty: { raw: value } },
    mode: { trackContainerResize() {}, allocatedWidth: 800, allocatedHeight: 600 },
  });
  control.init(context(), () => {
    const output = control.getOutputs();
    if (Object.hasOwn(output, "documentDirty")) value = output.documentDirty;
    outputs.push(value);
    control.updateView(context());
  }, {}, {});
  return {
    control, editor, outputs,
    value: () => value,
    ready: () => ready,
    finish: (content = "<p>Guardado</p>") => download.resolve({ html: content, fileName: "test.html" }),
    fail: () => download.reject(new Error("Error de red")),
    edit: (content) => { html = content; editor.updateDirtyState(html); },
    refresh: (incoming) => { value = incoming; control.updateView(context()); },
  };
}

for (const initialValue of [true, false, null]) {
  test(`successful load publishes No even when the bound value was ${initialValue}`, async () => {
    const app = setup(initialValue);
    assert.equal(Object.hasOwn(app.control.getOutputs(), "documentDirty"), false);
    assert.equal(app.outputs.length, 0);
    app.finish();
    await app.ready();
    assert.equal(app.value(), false);
    assert.deepEqual(app.outputs, [false]);
    assert.equal(app.editor.root.inert, false);
  });
}

test("discarding and opening a new instance resets an old Yes", async () => {
  const first = setup(false);
  first.finish();
  await first.ready();
  first.edit("<p>Sin guardar</p>");
  assert.equal(first.value(), true);
  first.control.destroy();
  const reopened = setup(first.value());
  reopened.finish();
  await reopened.ready();
  assert.equal(reopened.value(), false);
  assert.equal(reopened.editor.collectHtml(), "<p>Guardado</p>");
});

test("edits and undo publish changes without repeated notifications", async () => {
  const app = setup();
  app.finish();
  await app.ready();
  app.edit("<p>Cambio</p>");
  app.edit("<p>Otro cambio</p>");
  app.edit("<p>Guardado</p>");
  assert.deepEqual(app.outputs, [false, true, false]);
});

test("updateView never replaces pending document edits with an incoming No", async () => {
  const app = setup();
  app.finish();
  await app.ready();
  app.edit("<p>Cambio</p>");
  app.refresh(false);
  assert.equal(app.control.getOutputs().documentDirty, true);
  assert.deepEqual(app.outputs, [false, true]);
});

test("a failed download does not mark the blank error document as saved", async () => {
  const app = setup(false);
  app.fail();
  await app.ready();
  assert.equal(app.value(), true);
  assert.equal(app.editor.lastSavedHtml, null);
  app.edit("<p><br></p>");
  assert.equal(app.value(), true);
});

test("successful save clears dirty", async () => {
  const app = setup();
  app.finish();
  await app.ready();
  app.edit("<p>Cambio</p>");
  await app.editor.save();
  assert.equal(app.value(), false);
});

test("failed save leaves dirty set", async () => {
  const app = setup(false, { upload: async () => { throw new Error("Error de red"); } });
  app.finish();
  await app.ready();
  app.edit("<p>Cambio</p>");
  await app.editor.save();
  assert.equal(app.value(), true);
});

test("edits during upload remain dirty after the upload completes", async () => {
  const upload = deferred();
  const started = deferred();
  const app = setup(false, { upload: () => { started.resolve(); return upload.promise; } });
  app.finish();
  await app.ready();
  app.edit("<p>Enviado</p>");
  const saving = app.editor.save();
  await started.promise;
  app.edit("<p>Edición posterior</p>");
  upload.resolve();
  await saving;
  assert.equal(app.value(), true);
  assert.equal(app.editor.lastSavedHtml, "<p>Enviado</p>");
});

test("a late load completion cannot notify a destroyed PCF", async () => {
  const app = setup();
  app.control.destroy();
  app.finish();
  await app.ready();
  assert.deepEqual(app.outputs, []);
});

test("a late save completion cannot clear dirty after leaving the form", async () => {
  const upload = deferred();
  const started = deferred();
  const app = setup(false, { upload: () => { started.resolve(); return upload.promise; } });
  app.finish();
  await app.ready();
  app.edit("<p>Cambio</p>");
  const saving = app.editor.save();
  await started.promise;
  app.control.destroy();
  upload.resolve();
  await saving;
  assert.deepEqual(app.outputs, [false, true]);
});
