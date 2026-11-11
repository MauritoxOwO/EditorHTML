# Selección y edición entre páginas

La selección nativa vive en el workspace editable común. Las páginas y sus divisores siguen siendo contenedores visuales; antes de editar, `DocumentSelection` normaliza los extremos al contenido de `.hwe-page-inner`, conserva anchor/focus y calcula la dirección original. Así se aceptan selecciones de ratón, teclado y Ctrl+A con endpoints de texto o elemento, y se aíslan selecciones de otras instancias del editor.

`DocumentSelection.replaceSelection` calcula rangos locales para cada página afectada y los modifica desde la última hacia la primera. Nunca llama `deleteContents` con un rango que cruce páginas. Excluye cabeceras, nodos ocultos y contenido `contenteditable=false`; para tablas parciales, limita las operaciones a las celdas sin quitar filas ni celdas, y mantiene cualquier celda protegida. Una tabla completa se elimina sólo cuando toda ella es seleccionable. Los extremos de párrafo compatibles se unen usando el formato del párrafo inicial. Los reemplazos de texto plano conservan saltos de línea como `<br>`.

`EditorComponent` intercepta la edición de selecciones que cruzan páginas o estructuras, incluido beforeinput, corte, pegado, Enter y saltos de línea. El pegado HTML va al inner que contiene el caret; el importador de Word sigue atendiendo su formato y el HTML normal pasa por limpieza de scripts, atributos de evento y atributos internos. El copiado y corte multipágina generan HTML sin wrappers de página y texto plano que conserva saltos de línea, párrafos vacíos y celdas vacías. Los headers repetidos de una misma tabla sólo se copian una vez. Los fragmentos del mismo párrafo se reúnen al copiar. Las operaciones de formato reciben la selección normalizada y comparten el bookmark de `HistoryViewState` con historial y repaginación.

`HistoryViewState` guarda anchor y focus como posiciones de contenido, no como offsets de nodos que el paginador puede reemplazar. Al restaurar, usa `setBaseAndExtent`, conserva la dirección, reconoce contenido atómico y ajusta el scroll al caret. La repaginación pendiente se aplaza durante composición IME y arrastre de selección. Backspace al inicio y Delete al final de página unen bloques compatibles con la página vecina, atraviesan wrappers generados y se detienen ante objetos o saltos manuales. Los elementos `li` de listas distintas no se fusionan.

## Comprobación

La regresión del párrafo 21 del documento de ejemplo se comprueba con
`node --test tests/paragraph-reflow-browser.cjs` y el sandbox en ejecución.
La prueba pulsa ocho veces Enter, vuelve con ocho Retrocesos y sigue borrando:
comprueba el texto, los párrafos vacíos y la ocupación de la página. Se puede
repetir en Edge con `$env:BROWSER_CHANNEL='msedge'`.

El fallo provenía de detener el corte al encontrar un nodo de texto corto cuando
la continuación ya tenía contenido. Al repaginar texto dividido anteriormente,
esto creaba un párrafo temporal por palabra: sus márgenes falseaban la altura
antes de reunirlos. El corte ahora acumula esos nodos en una sola continuación.

Desde la raíz del proyecto:

```powershell
node --test tests/*.test.cjs
npx.cmd tsc --noEmit --skipLibCheck --target ES2022 --module ESNext --moduleResolution bundler --lib DOM,ES2022 EditorHTML/selection/DocumentSelection.ts EditorHTML/controllers/HistoryViewState.ts EditorHTML/controllers/PasteController.ts EditorHTML/controllers/PageBackspaceController.ts EditorHTML/app/EditorComponent.ts EditorHTML/controllers/StyleSelectionTracker.ts EditorHTML/controllers/TextSelectionFormatter.ts EditorHTML/ui/Toolbar.ts
npm.cmd run build
```

Para la comprobación de navegador, inicia `npm.cmd run dev:web` en una terminal y usa en otra el puerto que Vite indique:

```powershell
$env:EDITOR_URL='http://127.0.0.1:5174/?fixture=cross-page-selection'
node --test tests/cross-page-browser.cjs
Remove-Item Env:EDITOR_URL
```

Las pruebas DOM cubren dirección, Ctrl+A del workspace, extremos protegidos, reemplazo entre páginas, tablas y celdas protegidas, fragmentos de párrafo, copia HTML/texto y fronteras de Backspace. Chromium verifica ratón/teclado, borrar, copiar, pegar, cortar y Enter con undo, Negrita y alineación desde caret, y drag-and-drop interno hacia ambos lados. La prueba Ctrl+A compara el texto después de normalizar espacios porque la repaginación puede cambiar los separadores de bloque. La prueba IME simula eventos de composición e inserción DOM; no representa una IME del sistema operativo.

## Límites comprobados

Las pruebas de drop usan eventos `DragEvent` sintéticos; las selecciones de origen y los gestos de selección entre páginas usan Chromium. Chromium y Edge se comprobaron en headless: en Edge, arrastres forward/backward, zoom 1.18→1.28, Delete y Ctrl+Z mantuvieron texto, dirección, offsets y las 116 celdas; un borrado parcial dentro de celda conservó filas y columnas. La geometría de caret depende de `caretPositionFromPoint` o `caretRangeFromPoint`; si el navegador no expone ninguno o no resuelve el punto, el drop se cancela y conserva el documento. Las operaciones de tabla que se originan en comandos de fila/columna siguen pasando por sus controladores de tabla y sus reglas de integridad.
