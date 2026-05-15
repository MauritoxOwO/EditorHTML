import "../../EditorHTML/css/editor.css";
import { EditorComponent } from "../../EditorHTML/lifecycle/EditorComponent";
import "./styles.css";

const LOCAL_STORAGE_KEY = "editorhtml.local.currentHtml";

const editorHost = document.querySelector<HTMLDivElement>("#editor-host");
const fileInput = document.querySelector<HTMLInputElement>("#html-file");
const sourceName = document.querySelector<HTMLSpanElement>("#source-name");
const sampleButton = document.querySelector<HTMLButtonElement>("#sample-button");
const downloadButton = document.querySelector<HTMLButtonElement>("#download-button");

if (!editorHost || !fileInput || !sourceName || !sampleButton || !downloadButton) {
  throw new Error("No se pudo inicializar el harness local del editor.");
}

let currentFileName = "documento-ejemplo.html";
const query = new URLSearchParams(window.location.search);
const fixtureName = query.get("fixture");

const getInitialHtml = (): string => {
  return localStorage.getItem(LOCAL_STORAGE_KEY) ?? makeSampleHtml();
};

const editor = new EditorComponent(editorHost, undefined, {
  initialHtml: getInitialHtml(),
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
  currentFileName = `${name}.html`;
  sourceName!.textContent = `Fixture: ${name}`;
  await editor.loadHtml(html);
}

fileInput.addEventListener("change", async () => {
  const file = fileInput.files?.[0];
  if (!file) return;

  const html = await file.text();
  currentFileName = file.name;
  sourceName.textContent = file.name;
  localStorage.setItem(LOCAL_STORAGE_KEY, html);
  await editor.loadHtml(html);
  fileInput.value = "";
});

sampleButton.addEventListener("click", async () => {
  currentFileName = "documento-ejemplo.html";
  sourceName.textContent = "Documento de ejemplo";
  const html = makeSampleHtml();
  localStorage.setItem(LOCAL_STORAGE_KEY, html);
  await editor.loadHtml(html);
});

downloadButton.addEventListener("click", () => {
  const blob = new Blob([editor.getHtml()], { type: "text/html;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = currentFileName.replace(/\.(htm|html)$/i, "") + ".html";
  anchor.click();
  URL.revokeObjectURL(url);
});

function makeSampleHtml(): string {
  const paragraphs = Array.from({ length: 26 }, (_, index) => {
    const n = index + 1;
    return `<p><strong>Parrafo ${n}.</strong> Este texto fuerza el flujo entre paginas y permite comprobar que el caret vuelve a una posicion razonable despues del rebalanceo. La frase incluye palabras largas como internacionalizacion, responsabilidades y documentacion para probar saltos de linea.</p>`;
  }).join("");

  return `
    <h1>Documento de prueba</h1>
    <p>Este contenido se carga localmente, pero usa el mismo componente, paginador, estilos A4 y toolbar que el PCF.</p>
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
