// Real Chrome layout and component, without Dataverse or a development server.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const ts = require('typescript');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE ||
  'C:/Users/Mauro/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright');
const source = path.resolve(process.env.EDITOR_SOURCE_DIR || path.join(__dirname, '../EditorHTML'));

function runtime() {
  const modules = new Map();
  function collect(filename) {
    filename = path.resolve(filename);
    if (modules.has(filename)) return filename;
    const output = ts.transpileModule(fs.readFileSync(filename, 'utf8'), {
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 }, fileName: filename,
    }).outputText;
    const entry = { output, deps: {} }; modules.set(filename, entry);
    for (const [, dep] of output.matchAll(/require\("([^"]+)"\)/g)) {
      if (!dep.startsWith('.')) throw new Error(`Unexpected dependency: ${dep}`);
      entry.deps[dep] = collect(path.resolve(path.dirname(filename), dep + '.ts'));
    }
    return filename;
  }
  const entry = collect(path.join(source, 'app/EditorComponent.ts'));
  const paginator = collect(path.join(source, 'pagination/PaginatorDom.ts'));
  return `(()=>{const factories={${[...modules].map(([id, m]) => `${JSON.stringify(id)}:(module,exports,require)=>{${m.output}\n}`).join(',')}};
    const deps=${JSON.stringify(Object.fromEntries([...modules].map(([id, m]) => [id, m.deps])))};
    const cache={};function load(id){if(cache[id])return cache[id].exports;const m={exports:{}};cache[id]=m;factories[id](m,m.exports,n=>load(deps[id][n]));return m.exports;}
    window.EditorComponent=load(${JSON.stringify(entry)}).EditorComponent;
    window.pageOverflows=load(${JSON.stringify(paginator)}).pageOverflows;})();`;
}

test('261 páginas: formato local, selección, scroll, historial y aislamiento del formulario', { timeout: 180000 }, async t => {
  const browser = await chromium.launch({ headless: true,
    executablePath: process.env.BROWSER_PATH || 'C:/Program Files/Google/Chrome/Application/chrome.exe' });
  t.after(() => browser.close());
  const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
  const errors = []; page.on('pageerror', error => errors.push(error.message));
  await page.setContent('<style>html,body{height:100%;margin:0}#host{height:900px}</style><p id="form">Otro campo del formulario</p><div id="host"></div>');
  await page.addStyleTag({ content: fs.readFileSync(path.join(source, 'css/editor.css'), 'utf8') });
  await page.addScriptTag({ content: runtime() });
  const result = await page.evaluate(async historical => {
    const host = document.querySelector('#host');
    const editor = new window.EditorComponent(host, {}, { initialHtml: '<p>Inicio</p>',
      paragraphStyleCatalog: { styles: [
        {label:'Cuerpo',className:'cuerpo',cssText:'font-size:14px;line-height:20px;margin:0'},
        {label:'Grande',className:'grande',cssText:'font-size:20px;line-height:26px;margin:0'},
      ], fonts: [] } });
    await editor.init();
    const settle = () => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const flush = async () => {
      await settle(); let frames = 0;
      while (editor.rebalanceFrame !== undefined || editor.pendingRebalance || editor.historyController.isRestoring || editor.loadingOperations > 0) {
        if (++frames > 100) throw new Error('La paginación no converge');
        await settle();
      }
    };
    const html = Array.from({length:3915}, (_, i) => `<p class="cuerpo">Párrafo ${i+1}. ${'Texto de prueba para comprobar el rendimiento del editor al cambiar el formato. '.repeat(3)}</p>`).join('');
    await editor.loadHtml(html); await flush();
    const text = () => [...host.querySelectorAll('.hwe-page-inner')].map(inner => inner.textContent).join('');
    const original = text();
    const initialPages = editor.pages.length;
    const selectedText = () => getSelection().toString().replace(/\s+/g,'');
    const selectThree = () => {
      const ps=host.querySelectorAll('.hwe-page-inner p');
      editor.workspace.focus({preventScroll:true});
      getSelection().setBaseAndExtent(ps[0].firstChild,0,ps[2].lastChild,ps[2].lastChild.length);
      editor.styleSelectionTracker.rememberTextSelection();
      return selectedText();
    };
    let toolbarUpdates=0; const update=editor.updateToolbarSelectionState.bind(editor);
    editor.updateToolbarSelectionState=()=>{toolbarUpdates++;update();};
    const p=host.querySelector('.hwe-page-inner p');
    const selectionStart=performance.now();
    for(let i=0;i<20;i++){getSelection().setBaseAndExtent(p.firstChild,0,p.firstChild,i+1);editor.handleSelectionChange();}
    await flush();
    const selectionMs=performance.now()-selectionStart;
    const localUpdates=toolbarUpdates;
    const field=document.querySelector('#form');getSelection().setBaseAndExtent(field.firstChild,0,field.firstChild,4);
    await flush();const foreignUpdates=toolbarUpdates-localUpdates;
    selectThree();await flush();
    const expectedSelection=selectedText();
    const styleStart=performance.now();editor.applyParagraphStyle('grande');await flush();
    const styleMs=performance.now()-styleStart;
    const styleSelection=selectedText()===expectedSelection;
    const styleOverflow=editor.pages.filter(window.pageOverflows).length;
    const styleText=text()===original;
    const styled=[...host.querySelectorAll('p.grande')].map(p=>p.textContent).join('').replace(/\s+/g,'');
    const undoStart=performance.now();editor.historyController.undo();await flush();
    const undoText=text()===original;
    const undoStyles=host.querySelectorAll('p.grande').length;
    const undoMs=performance.now()-undoStart;
    selectThree();await flush();
    const sizeStart=performance.now();editor.applyFontSize('16pt');await flush();
    const sizeMs=performance.now()-sizeStart;
    const sizeText=text()===original;
    const sizeOverflow=editor.pages.filter(window.pageOverflows).length;
    const scrollStart=performance.now();
    for(let i=0;i<20;i++){editor.workspace.scrollTop=i*1000;await new Promise(requestAnimationFrame);}
    const scrollMs=performance.now()-scrollStart;
    const pending=editor.rebalanceFrame!==undefined||!!editor.pendingRebalance;
    const result={historical,initialPages,selectionMs,localUpdates,foreignUpdates,styleMs,sizeMs,scrollMs,styleOverflow,sizeOverflow,styleText,sizeText,styleSelection,styledCorrectly:styled===expectedSelection,undoText,undoStyles,undoMs,pending};
    editor.destroy();return result;
  }, !!process.env.EDITOR_SOURCE_DIR);
  console.log(JSON.stringify(result));
  assert.deepEqual(errors, []);
  assert.equal(result.initialPages, 261);
  assert.equal(result.styleText, true);assert.equal(result.sizeText, true);
  assert.equal(result.styleOverflow, 0);assert.equal(result.sizeOverflow, 0);
  assert.equal(result.styleSelection, true);assert.equal(result.styledCorrectly, true);
  assert.equal(result.undoText, true);assert.equal(result.undoStyles, 0);
  assert.equal(result.pending, false);
  if (!process.env.EDITOR_SOURCE_DIR) {
    assert.ok(result.localUpdates <= 2); assert.equal(result.foreignUpdates, 0);
    assert.ok(result.selectionMs < 100);
    assert.ok(result.styleMs < Number(process.env.FORMAT_BUDGET_MS || 1000), `Estilo: ${result.styleMs}ms`);
  }
});
