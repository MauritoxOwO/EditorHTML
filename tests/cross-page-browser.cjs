// Run with Vite serving the local harness: npm.cmd run dev:web
// Playwright is supplied by the Codex runtime and can be overridden with PLAYWRIGHT_MODULE.
const assert = require("node:assert/strict");
const path = require("node:path");
const test = require("node:test");

const playwrightPath = process.env.PLAYWRIGHT_MODULE ||
  "C:/Users/Mauro/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright";
const { chromium } = require(playwrightPath);
const baseUrl = process.env.EDITOR_URL || "http://127.0.0.1:5173/?fixture=cross-page-selection";

async function openEditor(t) {
  const browser = await chromium.launch({ channel: "chrome", headless: true });
  t.after(async () => browser.close());
  const page = await browser.newPage({ viewport: { width: 1800, height: 9000 } });
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto(baseUrl);
  await page.waitForFunction(() => document.querySelectorAll(".hwe-page").length >= 4);
  await page.waitForTimeout(500);
  t.after(() => assert.deepEqual(errors, [], "la aplicación no debe lanzar errores en el navegador"));
  return page;
}

async function textPoints(page, firstTerm, lastTerm) {
  return page.evaluate(({ firstTerm, lastTerm }) => {
    const root = document.querySelector(".hwe-workspace");
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    const nodes = [];
    let content = "";
    let node;
    while ((node = walker.nextNode())) {
      if (!node.parentElement.closest(".hwe-page-inner")) continue;
      nodes.push({ node, start: content.length, end: content.length + node.length });
      content += node.data;
    }
    const characterRect = (offset) => {
      const entry = nodes.find((item) => offset >= item.start && offset < item.end);
      const range = document.createRange();
      range.setStart(entry.node, offset - entry.start);
      range.setEnd(entry.node, offset - entry.start + 1);
      return range.getBoundingClientRect();
    };
    const find = (term, chooseLast) => {
      const at = chooseLast ? content.lastIndexOf(term) : content.indexOf(term);
      if (at < 0) throw new Error(`No se encontró el texto de referencia: ${term}`);
      const start = characterRect(at);
      const end = characterRect(at + term.length - 1);
      return {
        start: { x: start.left + 1, y: start.top + start.height / 2 },
        end: { x: end.right - 1, y: end.top + end.height / 2 },
      };
    };
    return { first: find(firstTerm, false), last: find(lastTerm, true) };
  }, { firstTerm, lastTerm });
}

async function documentContent(page) {
  const sourceRoot = `/@fs/${path.resolve(__dirname, "../EditorHTML").replace(/\\/g, "/")}`;
  return page.evaluate(async (sourceRoot) => {
    const { DocumentSerializer } = await import(`${sourceRoot}/services/DocumentSerializer.ts`);
    const { DEFAULT_PAGE_SETUP } = await import(`${sourceRoot}/pagination/PageGeometry.ts`);
    const serializer = new DocumentSerializer();
    const root = document.querySelector(".pcf-html-editor-root");
    const html = serializer.collectHtml(root, [...root.querySelectorAll(".hwe-page")], DEFAULT_PAGE_SETUP);
    // Compare the whole document, including formatting, without runtime page
    // labels whose positions legitimately change when text is repaginated.
    return serializer.normalizeHtmlForDirtyCheck(html);
  }, sourceRoot);
}

async function dragAcrossFlow(page, backwards = false) {
  const points = await textPoints(page, "Inicio del flujo", "flujo.");
  const start = backwards ? points.last.end : points.first.start;
  const end = backwards ? points.first.start : points.last.end;
  await page.mouse.move(start.x, start.y);
  await page.mouse.down();
  await page.mouse.move(end.x, end.y, { steps: 60 });
  await page.mouse.up();
  await page.waitForTimeout(100);
  return page.evaluate(() => {
    const selection = window.getSelection();
    const range = selection.getRangeAt(0);
    return {
      text: selection.toString(),
      anchorPage: selection.anchorNode.parentElement.closest(".hwe-page")?.dataset.hwePageIndex ??
        Array.from(document.querySelectorAll(".hwe-page")).findIndex((item) => item.contains(selection.anchorNode)),
      focusPage: Array.from(document.querySelectorAll(".hwe-page")).findIndex((item) => item.contains(selection.focusNode)),
      affectedPages: Array.from(document.querySelectorAll(".hwe-page")).filter((item) => range.intersectsNode(item)).length,
    };
  });
}

async function clickCaret(page, term) {
  const points = await textPoints(page, term, "flujo.");
  await page.mouse.click(points.first.start.x, points.first.start.y);
}

async function selectionInfo(page) {
  return page.evaluate(() => {
    const selection = window.getSelection();
    if (!selection?.rangeCount) return { text: "", collapsed: true, anchorPage: -1, focusPage: -1, pages: 0 };
    const range = selection.getRangeAt(0);
    const allPages = Array.from(document.querySelectorAll(".hwe-page"));
    const content = range.cloneContents();
    content.querySelectorAll(".hwe-page-divider, [data-hwe-runtime-page-header]").forEach(node => node.remove());
    const seenHeaders = new Set();
    content.querySelectorAll("thead").forEach(header => {
      const id = header.closest("table[data-hwe-repeat-header='true']")?.getAttribute("data-hwe-table-flow-id");
      if (id && seenHeaders.has(id)) header.remove();
      else if (id) seenHeaders.add(id);
    });
    return {
      text: selection.toString(),
      logicalText: content.textContent.replace(/\s+/g, ""),
      collapsed: selection.isCollapsed,
      anchorPage: allPages.findIndex((item) => item.contains(selection.anchorNode)),
      focusPage: allPages.findIndex((item) => item.contains(selection.focusNode)),
      pages: allPages.filter((item) => range.intersectsNode(item)).length,
    };
  });
}

async function placeCaret(page, term, offset = 0) {
  await page.evaluate(({ term, offset }) => {
    const root = document.querySelector(".hwe-workspace");
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    let node;
    while ((node = walker.nextNode())) {
      const at = node.data.indexOf(term);
      if (at < 0) continue;
      const range = document.createRange();
      range.setStart(node, at + offset);
      range.collapse(true);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      return;
    }
    throw new Error(`No se encontró el caret para ${term}`);
  }, { term, offset });
}

async function selectText(page, term) {
  await page.evaluate((term) => {
    const root = document.querySelector(".hwe-workspace");
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    const nodes = [];
    let content = "";
    let node;
    while ((node = walker.nextNode())) {
      nodes.push({ node, start: content.length, end: content.length + node.length });
      content += node.data;
    }
    const at = content.indexOf(term);
    if (at < 0) throw new Error(`No se encontró el texto para arrastrar: ${term}`);
    const start = nodes.find((entry) => at >= entry.start && at < entry.end);
    const finish = nodes.find((entry) => at + term.length > entry.start && at + term.length <= entry.end);
    if (!start || !finish) throw new Error(`No se encontraron los extremos del texto: ${term}`);
    const range = document.createRange();
    range.setStart(start.node, at - start.start);
    range.setEnd(finish.node, at + term.length - finish.start);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
  }, term);
}

async function dispatchInternalMove(page, source, target, before = false) {
  await selectText(page, source);
  const points = await textPoints(page, target, target);
  const point = before ? points.first.start : points.last.end;
  return page.evaluate(({ x, y }) => {
    const workspace = document.querySelector(".hwe-workspace");
    const data = new DataTransfer();
    Object.defineProperty(data, "dropEffect", { configurable: true, get: () => "move" });
    workspace.dispatchEvent(new DragEvent("dragstart", { bubbles: true, dataTransfer: data }));
    const event = new DragEvent("drop", {
      bubbles: true, cancelable: true, clientX: x, clientY: y, dataTransfer: data,
    });
    workspace.dispatchEvent(event);
    return { prevented: event.defaultPrevented, effect: data.dropEffect, text: workspace.textContent };
  }, point);
}

async function dispatchMoveAt(page, target, before = false) {
  const points = await textPoints(page, target, target);
  const point = before ? points.first.start : points.last.end;
  return page.evaluate(({ x, y }) => {
    const workspace = document.querySelector(".hwe-workspace");
    const data = new DataTransfer();
    Object.defineProperty(data, "dropEffect", { configurable: true, get: () => "move" });
    workspace.dispatchEvent(new DragEvent("dragstart", { bubbles: true, dataTransfer: data }));
    const event = new DragEvent("drop", {
      bubbles: true, cancelable: true, clientX: x, clientY: y, dataTransfer: data,
    });
    workspace.dispatchEvent(event);
    return { prevented: event.defaultPrevented, text: workspace.textContent };
  }, point);
}

test("arrastrar entre páginas y borrar conserva el contenido exterior y permite deshacer", async (t) => {
  const page = await openEditor(t);
  const selected = await dragAcrossFlow(page);
  assert.ok(selected.text.startsWith("Inicio del flujo:"));
  assert.ok(selected.text.endsWith("Final del flujo."));
  assert.ok(selected.affectedPages >= 2, "el gesto de ratón debe seleccionar más de una página");

  await page.keyboard.press("Delete");
  await page.waitForTimeout(1000);
  const afterDelete = await page.evaluate(() => {
    const text = Array.from(document.querySelectorAll(".hwe-page-inner"), (inner) => inner.textContent).join("");
    return {
      text,
      protected: document.querySelector("p[contenteditable='false']")?.textContent,
      manualBreak: !!document.querySelector("[data-hwe-manual-page-break='true']"),
      tableCells: document.querySelectorAll(".hwe-page-inner td").length,
    };
  });
  assert.ok(!afterDelete.text.includes("Inicio del flujo:"));
  assert.ok(!afterDelete.text.includes("Final del flujo."));
  assert.ok(afterDelete.text.includes("Inicio de la prueba."));
  assert.ok(afterDelete.text.includes("Contenido posterior al salto manual."));
  assert.match(afterDelete.protected, /REGIÓN PROTEGIDA/);
  assert.equal(afterDelete.manualBreak, true);
  assert.ok(afterDelete.tableCells > 0);

  await page.keyboard.press("Control+z");
  await page.waitForTimeout(400);
  const restored = await page.evaluate(() => Array.from(
    document.querySelectorAll(".hwe-page-inner"), (inner) => inner.textContent
  ).join(""));
  assert.ok(restored.includes("Inicio del flujo:"));
  assert.ok(restored.includes("Final del flujo."));
});

test("el arrastre hacia atrás mantiene el ancla final y el foco inicial", async (t) => {
  const page = await openEditor(t);
  const selected = await dragAcrossFlow(page, true);
  assert.ok(selected.text.startsWith("Inicio del flujo:"));
  assert.ok(selected.text.endsWith("Final del flujo."));
  assert.ok(selected.affectedPages >= 2);
  assert.ok(selected.anchorPage > selected.focusPage, "el navegador debe conservar la dirección inversa del arrastre");
});

test("Shift+clic extiende el caret real desde una página a otra", async (t) => {
  const page = await openEditor(t);
  const points = await textPoints(page, "Inicio del flujo", "flujo.");
  await page.mouse.click(points.first.start.x, points.first.start.y);
  await page.keyboard.down("Shift");
  await page.mouse.click(points.last.end.x, points.last.end.y);
  await page.keyboard.up("Shift");
  const selected = await selectionInfo(page);
  assert.ok(selected.text.startsWith("Inicio del flujo:"));
  assert.ok(selected.text.includes("Final del flujo"));
  assert.ok(selected.pages >= 2);
  assert.ok(selected.anchorPage < selected.focusPage);
});

test("Shift+flecha abajo cruza el final visual de una página", async (t) => {
  const page = await openEditor(t);
  await clickCaret(page, "Inicio del flujo");
  await page.keyboard.down("Shift");
  for (let index = 0; index < 80; index++) await page.keyboard.press("ArrowDown");
  await page.keyboard.up("Shift");
  const selected = await selectionInfo(page);
  assert.ok(selected.text.startsWith("Inicio del flujo:"));
  assert.ok(selected.pages >= 2);
  assert.ok(selected.anchorPage < selected.focusPage);
});

test("Ctrl+A del editor abarca el documento y Negrita conserva selección y extremos", async (t) => {
  const page = await openEditor(t);
  await clickCaret(page, "Inicio de la prueba");
  await page.keyboard.press("Control+a");
  const before = await selectionInfo(page);
  assert.equal(before.pages, await page.locator(".hwe-page").count());
  assert.ok(before.text.includes("REGIÓN PROTEGIDA"));
  await page.locator('button[title="Negrita (Ctrl+B)"]').click();
  await page.waitForTimeout(1000);
  const after = await selectionInfo(page);
  // Formatting can move a repeated table header to another row/page boundary.
  // Compare every selected document character, including protected text, once.
  assert.equal(after.logicalText, before.logicalText);
  assert.equal(after.collapsed, false);
  assert.equal(after.anchorPage, before.anchorPage);
  assert.equal(after.focusPage, before.focusPage);
});

test("copy de Ctrl+A no exporta wrappers ni duplica cabeceras de tabla", async (t) => {
  const page = await openEditor(t);
  await clickCaret(page, "Inicio de la prueba");
  await page.keyboard.press("Control+a");
  const copied = await page.evaluate(() => {
    const data = new DataTransfer();
    const event = new ClipboardEvent("copy", { bubbles: true, cancelable: true, clipboardData: data });
    document.querySelector(".hwe-workspace").dispatchEvent(event);
    return {
      prevented: event.defaultPrevented,
      text: data.getData("text/plain"),
      html: data.getData("text/html"),
    };
  });
  assert.equal(copied.prevented, true);
  assert.ok(copied.text.includes("Inicio de la prueba."));
  assert.ok(!copied.text.includes("REGIÓN PROTEGIDA"));
  assert.ok(!copied.html.includes(".hwe-page"));
  assert.ok(!copied.html.includes("data-hwe-page-divider"));
  assert.equal((copied.html.match(/<thead/gi) ?? []).length, 1);
  assert.equal((copied.text.match(/Sección/g) ?? []).length, 1);
});

test("pegar texto plano reemplaza una selección que cruza páginas", async (t) => {
  const page = await openEditor(t);
  await dragAcrossFlow(page);
  const pasted = await page.evaluate(() => {
    const data = new DataTransfer();
    data.setData("text/plain", "PEGADO PLANO");
    const event = new ClipboardEvent("paste", { bubbles: true, cancelable: true, clipboardData: data });
    document.querySelector(".hwe-workspace").dispatchEvent(event);
    return { prevented: event.defaultPrevented, text: document.querySelector(".hwe-workspace").textContent };
  });
  await page.waitForTimeout(1000);
  assert.equal(pasted.prevented, true);
  assert.ok(pasted.text.includes("PEGADO PLANO"));
  assert.ok(!pasted.text.includes("Inicio del flujo:"));
  assert.ok(!pasted.text.includes("Final del flujo."));
  assert.ok(pasted.text.includes("Contenido posterior al salto manual."));
});

test("pegar HTML reemplaza la selección y elimina etiquetas activas", async (t) => {
  const page = await openEditor(t);
  await dragAcrossFlow(page);
  const pasted = await page.evaluate(() => {
    const data = new DataTransfer();
    data.setData("text/html", "<p><strong>PEGADO HTML</strong><script>window.pasteInjected = true</script></p>");
    const event = new ClipboardEvent("paste", { bubbles: true, cancelable: true, clipboardData: data });
    document.querySelector(".hwe-workspace").dispatchEvent(event);
    return {
      prevented: event.defaultPrevented,
      text: document.querySelector(".hwe-workspace").textContent,
      html: document.querySelector(".hwe-workspace").innerHTML,
    };
  });
  await page.waitForTimeout(1000);
  assert.equal(pasted.prevented, true);
  assert.ok(pasted.text.includes("PEGADO HTML"));
  assert.ok(!pasted.text.includes("Inicio del flujo:"));
  assert.ok(!pasted.html.includes("<script"));
  assert.ok(pasted.text.includes("Contenido posterior al salto manual."));
});

test("Negrita desde caret permite escribir texto con formato activo", async (t) => {
  const page = await openEditor(t);
  await placeCaret(page, "Inicio de la prueba", 7);
  await page.locator('button[title="Negrita (Ctrl+B)"]').click();
  await page.keyboard.type("X ");
  await page.waitForTimeout(100);
  const result = await page.evaluate(() => {
    const inserted = Array.from(document.querySelectorAll(".hwe-page-inner b, .hwe-page-inner strong"))
      .find((element) => element.textContent?.replace(/\u00a0/g, " ").includes("X "));
    const selection = window.getSelection();
    return { bold: !!inserted, text: inserted?.textContent };
  });
  assert.equal(result.bold, true);
  assert.match(result.text, /X/);
});

test("alinear desde caret aplica el estilo a todo el flujo de párrafo paginado", async (t) => {
  const page = await openEditor(t);
  const flowId = await page.evaluate(() => {
    const walker = document.createTreeWalker(document.querySelector(".hwe-workspace"), NodeFilter.SHOW_TEXT);
    let node;
    while ((node = walker.nextNode())) {
      if (!node.data.includes("Inicio del flujo")) continue;
      return node.parentElement.closest("p")?.getAttribute("data-hwe-text-flow-id");
    }
    return null;
  });
  assert.ok(flowId);
  await placeCaret(page, "Inicio del flujo", 4);
  await page.locator('button[title="Centrar"]').click();
  await page.waitForTimeout(500);
  const aligned = await page.evaluate((flowId) => Array.from(
    document.querySelectorAll(`.hwe-page-inner p[data-hwe-text-flow-id="${flowId}"]`)
  ).map((paragraph) => paragraph.style.textAlign), flowId);
  assert.ok(aligned.length >= 2);
  assert.ok(aligned.every((value) => value === "center"));
});

test("Enter sobre selección multipágina se deshace en un paso", async (t) => {
  const page = await openEditor(t);
  await dragAcrossFlow(page);
  await page.keyboard.press("Enter");
  await page.waitForTimeout(500);
  const afterEnter = await page.locator(".hwe-workspace").textContent();
  assert.ok(!afterEnter.includes("Inicio del flujo:"));
  assert.ok(!afterEnter.includes("Final del flujo."));
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(500);
  const restored = await page.locator(".hwe-workspace").textContent();
  assert.ok(restored.includes("Inicio del flujo:"));
  assert.ok(restored.includes("Final del flujo."));
});

test("cortar selección multipágina y deshacer restaura todo el texto", async (t) => {
  const page = await openEditor(t);
  await dragAcrossFlow(page);
  const before = await documentContent(page);
  const cut = await page.evaluate(() => {
    const data = new DataTransfer();
    const event = new ClipboardEvent("cut", { bubbles: true, cancelable: true, clipboardData: data });
    document.querySelector(".hwe-workspace").dispatchEvent(event);
    return { prevented: event.defaultPrevented, copied: data.getData("text/plain") };
  });
  assert.equal(cut.prevented, true);
  assert.ok(cut.copied.startsWith("Inicio del flujo:"));
  const afterCut = await page.locator(".hwe-workspace").textContent();
  assert.ok(!afterCut.includes("Inicio del flujo:"));
  assert.ok(!afterCut.includes("Final del flujo."));
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(500);
  assert.equal(await documentContent(page), before);
});

test("la composición simulada aplaza la repaginación hasta compositionend", async (t) => {
  const page = await openEditor(t);
  await dragAcrossFlow(page);
  const beforePages = await page.locator(".hwe-page").count();
  const composing = await page.evaluate(() => {
    const workspace = document.querySelector(".hwe-workspace");
    workspace.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true, data: "" }));
    const selection = window.getSelection();
    const range = selection.getRangeAt(0);
    const inserted = document.createTextNode("ñ");
    range.insertNode(inserted);
    range.setStartAfter(inserted);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
    const input = new InputEvent("input", { bubbles: true, inputType: "insertText", data: "ñ", isComposing: true });
    workspace.dispatchEvent(input);
    return { pages: document.querySelectorAll(".hwe-page").length, composing: input.isComposing };
  });
  assert.equal(composing.pages, beforePages);
  assert.equal(composing.composing, true);
  await page.evaluate(() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))));
  assert.equal(await page.locator(".hwe-page").count(), beforePages);
  assert.ok((await page.locator(".hwe-workspace").textContent()).includes("ñ"));
  await page.evaluate(() => document.querySelector(".hwe-workspace")
    .dispatchEvent(new CompositionEvent("compositionend", { bubbles: true, data: "ñ" })));
  await page.waitForTimeout(500);
  assert.ok((await page.locator(".hwe-workspace").textContent()).includes("ñ"));
});

test("el drop interno mueve texto hacia delante y mantiene un único destino", async (t) => {
  const page = await openEditor(t);
  const moved = await dispatchInternalMove(page, "Inicio de la prueba.", "flujo.");
  await page.waitForTimeout(500);
  assert.equal(moved.prevented, true);
  assert.equal((moved.text.match(/Inicio de la prueba\./g) ?? []).length, 1);
  assert.ok(moved.text.indexOf("Inicio de la prueba.") > moved.text.indexOf("Final del flujo."));
});

test("el drop interno mueve texto hacia atrás y mantiene un único destino", async (t) => {
  const page = await openEditor(t);
  const moved = await dispatchInternalMove(page, "Final del flujo.", "Inicio de la prueba", true);
  await page.waitForTimeout(500);
  assert.equal(moved.prevented, true);
  assert.equal((moved.text.match(/Final del flujo\./g) ?? []).length, 1);
  assert.ok(moved.text.indexOf("Final del flujo.") < moved.text.indexOf("Inicio de la prueba."));
});

test("el drop mantiene el destino tras eliminar una selección de varias páginas", async (t) => {
  const page = await openEditor(t);
  await dragAcrossFlow(page);
  const moved = await dispatchMoveAt(page, "Contenido posterior al salto manual.");
  await page.waitForTimeout(500);
  assert.equal(moved.prevented, true);
  assert.equal((moved.text.match(/Inicio del flujo:/g) ?? []).length, 1);
  assert.equal((moved.text.match(/Final del flujo\./g) ?? []).length, 1);
  assert.ok(moved.text.indexOf("Contenido posterior al salto manual.") < moved.text.indexOf("Inicio del flujo:"));
});
