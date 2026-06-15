import "../../EditorHTML/css/editor.css";
import { EditorComponent } from "../../EditorHTML/app/EditorComponent";
import type { ParagraphStyleCatalog } from "../../EditorHTML/services/dataverse/styleApi";
import "./styles.css";

const LOCAL_STORAGE_KEY = "editorhtml.local.currentHtml";
const CODEX_SPIKE_FONT_FACE =
  '@font-face{font-family:"CodexSpikeTest";src:url("data:font/ttf;base64,AAEAAAAKAIAAAwAgT1MvMkUoRMEAAAEoAAAAYGNtYXACYAMRAAABlAAAAQhnbHlmEyMtpgAAAqQAAABAaGVhZC4phbwAAACsAAAANmhoZWEFogJXAAAA5AAAACRobXR4BpAAUAAAAYgAAAAMbG9jYQANAC0AAAKcAAAACG1heHAABQAKAAABCAAAACBuYW1laWAwPgAAAuQAAAHIcG9zdHDgcWsAAASsAAAALgABAAAAAQAAxuVV3V8PPPUAAQPoAAAAAOY5IPIAAAAA5jkg8gA3AAACbAL4AAAAAwACAAAAAAAAAAEAAAM0/0wAAAK8ACgASwJEAAEAAAAAAAAAAAAAAAAAAAADAAEAAAADAAgAAQAAAAAAAgAAAAAAAAAAAAAAAAAAAAAAAwIwAZAABQAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAPz8/PwAAACAAfQM0/0wAAAOEANwAAAAAAAAAAAAAAAAAAAAgAAACvAAoAWgAAAJsACgAAAACAAAAAwAAABQAAwABAAAAFAAEAPQAAAAOAAgAAgAGACMAWwBdAF8AewB9//8AAAAgACUAXQBfAGEAff//AAAAAP+l/6MAAP+FAAEADgAUAAAAAAB8AAAAAAABAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAgACAAIAAAANAA0AIAABAFAAAAJsArwAAwAAMyERIVACHP3kArwAAAEANwAAAjAC+AAHAAAzMzczAQMjF1XXQcP+7UGlm+ECF/784QAAAAAMAJYAAQAAAAAAAQAOAAAAAQAAAAAAAgAHAA4AAQAAAAAAAwAaABUAAQAAAAAABAAWAC8AAQAAAAAABQALAEUAAQAAAAAABgAWAFAAAwABBAkAAQAcAGYAAwABBAkAAgAOAIIAAwABBAkAAwA0AJAAAwABBAkABAAsAMQAAwABBAkABQAWAPAAAwABBAkABgAsAQZDb2RleFNwaWtlVGVzdFJlZ3VsYXJDb2RleFNwaWtlVGVzdCBSZWd1bGFyIDEuMENvZGV4U3Bpa2VUZXN0IFJlZ3VsYXJWZXJzaW9uIDEuMENvZGV4U3Bpa2VUZXN0LVJlZ3VsYXIAQwBvAGQAZQB4AFMAcABpAGsAZQBUAGUAcwB0AFIAZQBnAHUAbABhAHIAQwBvAGQAZQB4AFMAcABpAGsAZQBUAGUAcwB0ACAAUgBlAGcAdQBsAGEAcgAgADEALgAwAEMAbwBkAGUAeABTAHAAaQBrAGUAVABlAHMAdAAgAFIAZQBnAHUAbABhAHIAVgBlAHIAcwBpAG8AbgAgADEALgAwAEMAbwBkAGUAeABTAHAAaQBrAGUAVABlAHMAdAAtAFIAZQBnAHUAbABhAHIAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADAAAAAwECBXNwaWtlAAA=") format("truetype");font-weight:400;font-style:normal;}';
const SANDBOX_PARAGRAPH_STYLE_CATALOG: ParagraphStyleCatalog = {
  styles: [
    {
      label: "Texto general",
      className: "texto-general",
      cssText: `.texto-general {
  display: inline-block;
  text-indent: 20pt;
  margin: 0;
  text-align: justify;
  hyphens: auto;
  font-family: Calibri, Arial, sans-serif;
  font-size: 11pt;
}`,
    },
    {
      label: "Texto Spike",
      className: "texto-spike",
      cssText: `.texto-spike {
  font-family: "CodexSpikeTest";
  font-size: 16pt;
  line-height: 1.2;
}`,
    },
  ],
  fonts: [
    {
      label: "CodexSpikeTest",
      cssText: CODEX_SPIKE_FONT_FACE,
    },
  ],
};

const editorHost = document.querySelector<HTMLDivElement>("#editor-host");
const fileInput = document.querySelector<HTMLInputElement>("#html-file");
const sourceName = document.querySelector<HTMLSpanElement>("#source-name");
const sampleButton = document.querySelector<HTMLButtonElement>("#sample-button");

if (!editorHost || !fileInput || !sourceName || !sampleButton) {
  throw new Error("No se pudo inicializar el harness local del editor.");
}

const query = new URLSearchParams(window.location.search);
const fixtureName = query.get("fixture");

const getInitialHtml = (): string => {
  return localStorage.getItem(LOCAL_STORAGE_KEY) ?? makeSampleHtml();
};

const editor = new EditorComponent(editorHost, undefined, {
  initialHtml: getInitialHtml(),
  paragraphStyleCatalog: SANDBOX_PARAGRAPH_STYLE_CATALOG,
  saveHtml: (html) => {
    localStorage.setItem(LOCAL_STORAGE_KEY, html);
  },
});

void editor.init().then(() => {
  if (fixtureName) {
    void loadFixture(fixtureName);
  }
});

async function loadFixture(name: string): Promise<void> {
  const response = await fetch(`/fixtures/${encodeURIComponent(name)}.html`, {
    cache: "no-store",
  });

  if (!response.ok) {
    throw new Error(`No se pudo cargar el fixture local: ${name}`);
  }

  const html = await response.text();
  sourceName!.textContent = `Fixture: ${name}`;
  await editor.loadHtml(html);
}

fileInput.addEventListener("change", async () => {
  const file = fileInput.files?.[0];
  if (!file) return;

  const html = await file.text();
  sourceName.textContent = file.name;
  localStorage.setItem(LOCAL_STORAGE_KEY, html);
  await editor.loadHtml(html);
  fileInput.value = "";
});

sampleButton.addEventListener("click", async () => {
  sourceName.textContent = "Documento de ejemplo";
  const html = makeSampleHtml();
  localStorage.setItem(LOCAL_STORAGE_KEY, html);
  await editor.loadHtml(html);
});

function makeSampleHtml(): string {
  const paragraphs = Array.from({ length: 26 }, (_, index) => {
    const n = index + 1;
    return `<p><strong>Parrafo ${n}.</strong> Este texto fuerza el flujo entre paginas y permite comprobar que el caret vuelve a una posicion razonable despues del rebalanceo. La frase incluye palabras largas como internacionalizacion, responsabilidades y documentacion para probar saltos de linea.</p>`;
  }).join("");

  return `
    <h1>Documento de prueba</h1>
    <p>Este contenido se carga localmente, pero usa el mismo componente, paginador, estilos A4 y toolbar que el PCF.</p>
    <p class="texto-spike">Prueba visible de fuente CodexSpikeTest: ABC xyz 123.</p>
    <table>
      <thead>
        <tr><th>Concepto</th><th>Detalle</th><th>Estado</th></tr>
      </thead>
      <tbody>
        <tr><td>Tabla</td><td>Celdas con texto largo para validar cortes y ancho fijo.</td><td>Editable</td></tr>
        <tr><td>Imagenes</td><td>Las imagenes embebidas del HTML se miden antes de paginar.</td><td>Visible</td></tr>
        <tr><td>Caret</td><td>El marcador se conserva durante el rebalanceo de paginas.</td><td>En pruebas</td></tr>
      </tbody>
    </table>
    ${paragraphs}
  `;
}
