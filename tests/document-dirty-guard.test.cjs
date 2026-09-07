const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const script = fs.readFileSync(path.join(__dirname, "../webresources/ays_documentdirty_guard.js"), "utf8");
const COLUMN = "ays_documentdirty_sn";
const NOTICE = "ays_documentdirty_guard";

function setup(options = {}) {
  let value = options.value ?? false;
  const changes = new Set();
  const stages = new Set();
  const notifications = new Map();
  const calls = [];
  const dialogs = [];
  const errors = [];
  const attribute = {
    getValue: () => value,
    addOnChange: (handler) => changes.add(handler),
    removeOnChange: (handler) => changes.delete(handler),
    setValue: () => assert.fail("El script no debe escribir el indicador"),
    setSubmitMode: () => assert.fail("El script no debe alterar el envío de la columna"),
  };
  const form = {
    getAttribute: (name) => {
      assert.equal(name, COLUMN);
      return options.missingColumn ? null : attribute;
    },
    data: {
      entity: { getEntityName: () => options.table || "ays_versionanuncio" },
      process: options.noProcess ? undefined : {
        addOnPreStageChange: (handler) => stages.add(handler),
        removeOnPreStageChange: (handler) => stages.delete(handler),
      },
      save: () => assert.fail("El script no debe iniciar un guardado"),
    },
    ui: {
      setFormNotification: (message, level, id) => {
        notifications.set(id, { message, level });
        calls.push("notice");
      },
      clearFormNotification: (id) => notifications.delete(id),
    },
  };
  const host = {
    AYS: { existingFeature: true },
    console: { error: (...args) => errors.push(args) },
    Xrm: { Navigation: {
      openAlertDialog: (message) => {
        calls.push("dialog");
        if (options.dialogThrows) throw new Error("dialog failed");
        const entry = { message };
        dialogs.push(entry);
        return new Promise((resolve, reject) => {
          entry.resolve = resolve;
          entry.reject = reject;
        });
      },
    } },
  };
  vm.runInNewContext(script, { window: host, WeakSet }, { filename: "ays_documentdirty_guard.js" });
  const context = { getFormContext: () => form };
  function save(mode = 1) {
    let prevented = false;
    const result = host.AYS.DocumentDirtyGuard.onSave({
      ...context,
      getEventArgs: () => ({
        getSaveMode: () => mode,
        preventDefault: () => { prevented = true; calls.push("prevent"); },
      }),
    });
    assert.equal(result, undefined, "OnSave debe ser síncrono, no devolver una promesa");
    return prevented;
  }
  return {
    host, form, attribute, changes, stages, notifications, calls, dialogs, errors, context, save,
    load: () => host.AYS.DocumentDirtyGuard.onLoad(context),
    setValue: (next, fire = true) => {
      value = next;
      if (fire) changes.forEach((handler) => handler(context));
    },
  };
}

test("onLoad registers handlers once and preserves the existing namespace", () => {
  const app = setup();
  app.load();
  app.load();
  assert.equal(app.changes.size, 1);
  assert.equal(app.stages.size, 1);
  assert.equal(app.notifications.size, 0);
  assert.equal(app.host.AYS.existingFeature, true);
});

test("PCF changes show and clear only this script's notification", () => {
  const app = setup();
  app.notifications.set("another-script", { message: "Other notice" });
  app.load();
  app.setValue(true);
  assert.equal(app.notifications.get(NOTICE).level, "WARNING");
  assert.equal(app.dialogs.length, 0);
  app.setValue(false);
  assert.equal(app.notifications.has(NOTICE), false);
  assert.equal(app.notifications.has("another-script"), true);
});

for (const mode of [1, 2, 5, 6, 47, 59, 70]) {
  test(`dirty document synchronously blocks save mode ${mode}`, () => {
    const app = setup({ value: true });
    assert.equal(app.save(mode), true);
    assert.equal(app.calls[0], "prevent");
    assert.equal(app.dialogs.length, mode === 70 ? 0 : 1);
  });
  test(`clean document does not block save mode ${mode}`, () => {
    const app = setup();
    assert.equal(app.save(mode), false);
    assert.equal(app.dialogs.length, 0);
  });
}

test("onSave reads the current field even before an OnChange notification", () => {
  const app = setup();
  app.load();
  app.setValue(true, false);
  assert.equal(app.save(), true);
  app.setValue(false, false);
  assert.equal(app.save(), false);
  assert.equal(app.notifications.has(NOTICE), false);
});

test("an already true indicator is never reset on load", () => {
  const app = setup({ value: true });
  app.load();
  assert.equal(app.attribute.getValue(), true);
  assert.equal(app.notifications.get(NOTICE).level, "WARNING");
});

test("repeated autosaves remain silent apart from the form warning", () => {
  const app = setup({ value: true });
  for (let index = 0; index < 4; index++) assert.equal(app.save(70), true);
  assert.equal(app.dialogs.length, 0);
  assert.equal(app.notifications.size, 1);
});

test("only one alert is open at a time and accepting it does not save", async () => {
  const app = setup({ value: true });
  assert.equal(app.save(), true);
  assert.equal(app.save(2), true);
  assert.equal(app.dialogs.length, 1);
  app.dialogs[0].resolve();
  await Promise.resolve();
  assert.equal(app.attribute.getValue(), true);
  assert.equal(app.save(), true);
  assert.equal(app.dialogs.length, 2);
});

test("dialog rejection never releases the cancelled save and allows another alert", async () => {
  const app = setup({ value: true });
  assert.equal(app.save(), true);
  app.dialogs[0].reject(new Error("test rejection"));
  await Promise.resolve();
  assert.equal(app.errors.length, 1);
  assert.equal(app.save(), true);
  assert.equal(app.dialogs.length, 2);
});

test("synchronous dialog error does not bypass the guard", () => {
  const app = setup({ value: true, dialogThrows: true });
  assert.equal(app.save(), true);
  assert.equal(app.errors.length, 1);
  assert.equal(app.notifications.get(NOTICE).level, "WARNING");
});

test("missing column exposes a configuration error and blocks saving", () => {
  const app = setup({ missingColumn: true });
  app.load();
  assert.equal(app.changes.size, 0);
  assert.equal(app.notifications.get(NOTICE).level, "ERROR");
  assert.equal(app.save(), true);
});

test("registration on a different table has no effect", () => {
  const app = setup({ table: "account", value: true });
  app.load();
  assert.equal(app.changes.size, 0);
  assert.equal(app.stages.size, 0);
  assert.equal(app.notifications.size, 0);
  assert.equal(app.save(), false);
});

test("forms without a business process flow are supported", () => {
  const app = setup({ noProcess: true, value: true });
  assert.doesNotThrow(() => app.load());
  assert.equal(app.save(), true);
});

test("business-process stage changes are cancelled only while dirty", () => {
  const app = setup({ value: true });
  app.load();
  let blocked = 0;
  const context = { ...app.context, getEventArgs: () => ({ preventDefault: () => blocked++ }) };
  app.stages.forEach((handler) => handler(context));
  assert.equal(blocked, 1);
  app.setValue(false);
  app.stages.forEach((handler) => handler(context));
  assert.equal(blocked, 1);
});
