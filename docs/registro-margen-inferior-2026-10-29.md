# Registro de cambios: margen inferior del editor

Fecha: 29 de octubre de 2026.

Proyecto modificado: `C:\Users\Mauro\Documents\New project 3\EditorHTML-document-dirty-push`.

## Petición y alcance

Se ha aplicado la primera modificación autorizada: corregir el cálculo del margen inferior y compartirlo entre las comprobaciones del editor. Se documentan también los ajustes necesarios para que esa corrección funcione con el zoom y con las tablas paginadas.

Se conserva la configuración de página existente. Su valor inferior predeterminado es 20 mm. Esta entrega respeta el margen inferior que tenga la configuración activa: todavía no convierte los valores predeterminados en una política obligatoria ni ignora los márgenes de los documentos recibidos. Esa era la segunda propuesta y queda fuera de esta modificación.

No se han cambiado el CSS de márgenes, el tamaño A4, la columna de texto de 140 mm, los márgenes entre párrafos ni el importador de documentos.

## Problema comprobado

El cálculo anterior descontaba `Math.min(paddingBottom, 12)`. Aunque el padding inferior fuese de 20 mm, se descontaban como máximo 12 píxeles. Sin zoom, 20 mm equivalen aproximadamente a 75,59 píxeles CSS, mientras que 12 píxeles equivalen a unos 3,18 mm.

Había versiones de ese cálculo en la paginación, el servicio de disposición y el diagnóstico. Algunas tenían en cuenta el zoom y otras no. Además, dos comprobaciones de espacio disponible combinaban coordenadas de pantalla, que incluyen el zoom, con medidas CSS del documento, que no lo incluyen.

La prueba de navegador detectó un segundo problema directamente relacionado: al dividir una tabla larga, cada fragmento podía perder el formato compacto porque contenía menos de 60 filas. Las filas crecían después de la paginación y podían invadir el margen inferior. Corregir únicamente la resta de 12 píxeles no resolvía ese caso.

## Cambios de implementación y justificación de cada uno

### Cómo leer los snippets

Los bloques `diff` muestran el código anterior con `-` y el código añadido con `+`; las líneas sin estos signos aportan contexto. Se incluyen todos los cambios de los cuatro archivos de implementación existentes. El nuevo cálculo de geometría se reproduce completo. Los snippets de pruebas muestran las comprobaciones añadidas y se acompañan de una explicación del entorno que necesitan.

Los fragmentos se han extraído del código aplicado y del parche de la corrección. Esta ampliación solo modifica el registro; los resultados de pruebas que figuran más abajo corresponden a la validación del código ya realizada.

Las rutas de esta sección son relativas a la carpeta del proyecto indicada al principio.

### 1. Nuevo archivo `EditorHTML/pagination/PageContentGeometry.ts`

1. Se añade la interfaz `PageContentBounds`, con `top`, `bottom`, `contentHeight` y `scale`. **Justificación:** identificar claramente los límites del contenido y distinguir la altura interna del documento de las coordenadas de pantalla.

2. Se añade `getPageContentBounds(inner)` como cálculo compartido. **Justificación:** evitar que paginación, comprobación de desbordamientos y diagnóstico apliquen fórmulas diferentes al mismo documento.

3. Se leen completos los valores superior e inferior del padding. **Justificación:** la reserva declarada por el CSS debe ser la misma que se descuenta al calcular el área disponible; desaparece el máximo artificial de 12 píxeles.

4. Se descuentan también los bordes superior e inferior, si existen. **Justificación:** el rectángulo obtenido del navegador incluye los bordes y estos tampoco forman parte del espacio donde debe caber el contenido.

5. Se calcula la escala comparando la altura visible con la altura de la caja antes del zoom. Se prefiere la altura calculada por CSS y se distingue `border-box` de `content-box`. **Justificación:** `getBoundingClientRect()` incluye la transformación visual, mientras que el padding se obtiene en píxeles CSS. Además, `offsetHeight` redondea a enteros y puede introducir pequeñas diferencias con medidas en milímetros.

6. Se conserva una alternativa basada en `offsetHeight` cuando no hay una altura CSS calculada válida, y una escala de 1 cuando no puede medirse. **Justificación:** mantener un resultado definido en elementos sin geometría medible, sin dividir por cero.

7. Se devuelve la altura útil sin zoom y se limita a un mínimo de cero. **Justificación:** las dimensiones asignadas mediante estilos necesitan medidas del documento, y un conjunto de reservas mayor que la caja no debe producir una altura negativa.

8. Se añaden comentarios sobre las unidades y el redondeo. **Justificación:** dejar documentada una distinción que, de perderse, volvería a introducir errores con el zoom.

**Snippet añadido completo — puntos 1 a 8:**

```ts
export interface PageContentBounds {
  top: number;
  bottom: number;
  /** Content height in document CSS pixels, before the view zoom is applied. */
  contentHeight: number;
  scale: number;
}

/** Viewport bounds exclude the complete page padding, including at non-default zoom. */
export function getPageContentBounds(inner: HTMLElement): PageContentBounds {
  const styles = getComputedStyle(inner);
  const rect = inner.getBoundingClientRect();
  const paddingTop = parseFloat(styles.paddingTop) || 0;
  const paddingBottom = parseFloat(styles.paddingBottom) || 0;
  const borderTop = parseFloat(styles.borderTopWidth) || 0;
  const borderBottom = parseFloat(styles.borderBottomWidth) || 0;
  const verticalInsets = paddingTop + paddingBottom + borderTop + borderBottom;
  const cssHeight = parseFloat(styles.height);
  // Computed height preserves fractional CSS pixels; offsetHeight rounds them.
  const layoutHeight = Number.isFinite(cssHeight) && cssHeight > 0
    ? cssHeight + (styles.boxSizing === "border-box" ? 0 : verticalInsets)
    : inner.offsetHeight;
  const scale = layoutHeight > 0 && rect.height > 0 ? rect.height / layoutHeight : 1;

  return {
    top: rect.top + (borderTop + paddingTop) * scale,
    bottom: rect.bottom - (borderBottom + paddingBottom) * scale,
    contentHeight: Math.max(0, layoutHeight > 0
      ? layoutHeight - verticalInsets
      : inner.clientHeight - paddingTop - paddingBottom),
    scale,
  };
}
```

La línea `bottom: rect.bottom - (borderBottom + paddingBottom) * scale` reserva todo el padding inferior en coordenadas de pantalla. `contentHeight` se conserva en unidades del documento para asignar alturas CSS sin multiplicarlas por el zoom.

### 2. Archivo `EditorHTML/pagination/PaginatorDom.ts`

9. Se importa el cálculo compartido. `getContentLimitBottom()` pasa a devolver su límite inferior. **Justificación:** las divisiones de tablas y los otros consumidores de esta función reciben ahora el límite completo y corregido, conservando la API existente.

10. `pageOverflows()` compara el final del contenido con ese mismo límite. Se expresa su tolerancia como un píxel CSS convertido a la escala visible. **Justificación:** mantener la tolerancia de redondeo existente de forma coherente en todos los niveles de zoom. No se ha añadido una nueva reserva ni un nuevo margen.

11. `getContentHeight()` utiliza la altura útil compartida. **Justificación:** la altura disponible para imágenes y otros elementos debe proceder de la misma geometría que utiliza la paginación.

**Snippets antes/después — puntos 9 a 11:**

```diff
diff --git a/EditorHTML/pagination/PaginatorDom.ts b/EditorHTML/pagination/PaginatorDom.ts
index 0285f15..f384cb3 100644
--- a/EditorHTML/pagination/PaginatorDom.ts
+++ b/EditorHTML/pagination/PaginatorDom.ts
@@ -1,3 +1,5 @@
+import { getPageContentBounds } from "./PageContentGeometry";
+
 const SPLITTABLE_CONTAINER_TAGS = new Set([
   "DIV",
   "SECTION",
@@ -65,23 +67,19 @@ export function pageOverflows(page: HTMLElement): boolean {
   const contentBottom = getContentBottom(inner);
   if (contentBottom === null) return scrollOverflows;
 
-  return contentBottom > getContentLimitBottom(inner) + 1 || scrollOverflows;
+  const bounds = getPageContentBounds(inner);
+  return contentBottom > bounds.bottom + bounds.scale || scrollOverflows;
 }
 
 export function getContentLimitBottom(inner: HTMLElement): number {
-  const styles = getComputedStyle(inner);
-  const paddingBottom = parseFloat(styles.paddingBottom) || 0;
-  return inner.getBoundingClientRect().bottom - Math.min(paddingBottom, 12);
+  return getPageContentBounds(inner).bottom;
 }
 
 export function getContentHeight(page: HTMLElement): number {
   const inner = getInner(page);
   if (!inner) return page.clientHeight;
 
-  const styles = getComputedStyle(inner);
-  const paddingTop = parseFloat(styles.paddingTop) || 0;
-  const paddingBottom = parseFloat(styles.paddingBottom) || 0;
-  return Math.max(0, inner.clientHeight - paddingTop - paddingBottom);
+  return getPageContentBounds(inner).contentHeight;
 }
 
 export function getContentBottom(inner: HTMLElement): number | null {
```

La firma de `getContentLimitBottom()` se conserva. Por ello, las divisiones de tablas que ya llamaban a esta función reciben el nuevo límite sin cambiar su contrato ni duplicar la fórmula en `TablePaginator`.

### 3. Archivo `EditorHTML/services/EditorLayoutService.ts`

12. La comprobación `pageOverflows()` utiliza el cálculo compartido y su escala. Se elimina su función privada duplicada para el límite inferior. **Justificación:** la comprobación que utiliza el editor al editar debe coincidir con la que utiliza el paginador; eliminar la copia evita mantener dos fórmulas.

13. Se importa el helper existente `getTableFlowFragments()` y se calcula el número de filas de la tabla completa, sumando sus fragmentos. **Justificación:** el formato compacto debe depender de la tabla lógica completa, no del número de filas que casualmente caben en una hoja. Esto evita que sus filas crezcan al terminar de paginarla y vuelvan a desbordar.

14. La búsqueda de fragmentos se limita al workspace o a la raíz del editor más cercana. **Justificación:** dos controles PCF distintos pueden contener identificadores de flujo iguales; las filas de otro editor no deben influir en la densidad de esta tabla.

15. Se mantiene el recálculo a partir de las filas actuales. **Justificación:** si se eliminan filas y el total baja del umbral, el formato compacto debe poder retirarse. No se fija para siempre por el hecho de que la tabla haya sido larga una vez.

**Snippets antes/después — puntos 12 a 15:**

```diff
diff --git a/EditorHTML/services/EditorLayoutService.ts b/EditorHTML/services/EditorLayoutService.ts
index 6f306a6..0c44d6c 100644
--- a/EditorHTML/services/EditorLayoutService.ts
+++ b/EditorHTML/services/EditorLayoutService.ts
@@ -1,4 +1,6 @@
 import { getMeaningfulChildren, unwrapElement } from "../dom/EditableDom";
+import { getTableFlowFragments } from "../dom/TableFlow";
+import { getPageContentBounds } from "../pagination/PageContentGeometry";
 
 const TEXT_FLOW_SELECTOR = "p, h1, h2, h3, h4, h5, h6, blockquote, pre, ul, ol";
 const LONG_TABLE_MIN_ROWS = 60;
@@ -50,8 +52,8 @@ export class EditorLayoutService {
     const contentBottom = this.getContentBottom(inner);
     if (contentBottom === null) return scrollOverflows;
 
-    const scale = inner.getBoundingClientRect().height / (inner.offsetHeight || 1);
-    return contentBottom > this.getContentLimitBottom(inner, scale) + scale || scrollOverflows;
+    const bounds = getPageContentBounds(inner);
+    return contentBottom > bounds.bottom + bounds.scale || scrollOverflows;
   }
 
   private getDirectFlowTable(element: HTMLElement): HTMLTableElement | null {
@@ -100,7 +102,13 @@ export class EditorLayoutService {
   }
 
   private normalizeTableDensity(table: HTMLTableElement): void {
-    const rowCount = table.querySelectorAll("tbody tr, tfoot tr").length;
+    // Density belongs to the complete table, not to each page fragment.
+    // Otherwise splitting a long table removes its compact style and increases row heights.
+    const scope = table.closest(".hwe-workspace, .pcf-html-editor-root");
+    const rowCount = getTableFlowFragments(scope, table).reduce(
+      (count, fragment) => count + fragment.querySelectorAll("tbody tr, tfoot tr").length,
+      0
+    );
     const columnCount = this.getColumnCount(table);
     const minColumnPercent = this.getMinimumColumnPercent(table);
     const density = this.getTableDensity(rowCount, columnCount, minColumnPercent);
@@ -231,12 +239,6 @@ export class EditorLayoutService {
     });
   }
 
-  private getContentLimitBottom(inner: HTMLElement, scale: number): number {
-    const styles = getComputedStyle(inner);
-    const paddingBottom = parseFloat(styles.paddingBottom) || 0;
-    return inner.getBoundingClientRect().bottom - Math.min(paddingBottom, 12) * scale;
-  }
-
   private getContentBottom(inner: HTMLElement): number | null {
     const children = getMeaningfulChildren(inner);
     if (children.length === 0) return null;
```

El primer bloque elimina la discrepancia entre el servicio y el paginador. El recuento con `getTableFlowFragments(scope, table)` mantiene el formato compacto según el total de filas de la tabla; el último bloque elimina la función privada que contenía la fórmula anterior.

### 4. Archivo `EditorHTML/pagination/Paginator.ts`

16. `elementFitsOnFreshPage()` obtiene tanto el límite superior como el inferior del cálculo compartido. También convierte la tolerancia de un píxel CSS a la escala visible. **Justificación:** anteriormente el padding superior se restaba sin zoom dentro de una operación con coordenadas que sí tenían zoom. Esto podía considerar que cabían bloques demasiado altos.

17. `fitImagesToAvailableSpace()` convierte el espacio visible disponible a píxeles CSS antes de asignar `max-height`. **Justificación:** una medida de pantalla al 250 % no puede asignarse directamente como altura interna del documento; hacerlo ampliaría incorrectamente la imagen y podría hacerla invadir el margen inferior.

18. Se conserva el redondeo hacia abajo al asignar la altura máxima. **Justificación:** no asignar una altura superior al espacio disponible por una fracción de píxel.

**Snippets antes/después — puntos 16 a 18:**

```diff
diff --git a/EditorHTML/pagination/Paginator.ts b/EditorHTML/pagination/Paginator.ts
index 8a0497d..be7bfc6 100644
--- a/EditorHTML/pagination/Paginator.ts
+++ b/EditorHTML/pagination/Paginator.ts
@@ -1,6 +1,7 @@
 import { KeepTogetherController } from "./KeepTogetherController";
 import { TablePaginator } from "./TablePaginator";
 import { TextBlockSplitter } from "./TextBlockSplitter";
+import { getPageContentBounds } from "./PageContentGeometry";
 import { hweDebugLog, hweDebugStart } from "../debug/DebugLogger";
 import {
   AUTO_IMAGE_MAX_HEIGHT,
@@ -461,12 +462,10 @@ export class Paginator {
     const inner = getInner(page);
     if (!inner) return false;
 
-    const styles = getComputedStyle(inner);
-    const paddingTop = parseFloat(styles.paddingTop) || 0;
-    const availableHeight =
-      getContentLimitBottom(inner) - (inner.getBoundingClientRect().top + paddingTop);
+    const bounds = getPageContentBounds(inner);
+    const availableHeight = bounds.bottom - bounds.top;
 
-    return element.getBoundingClientRect().height <= availableHeight + 1;
+    return element.getBoundingClientRect().height <= availableHeight + bounds.scale;
   }
 
   private resolveOverflow(index: number): boolean {
@@ -968,12 +967,12 @@ export class Paginator {
     const inner = getInner(page);
     if (!inner) return false;
 
-    const contentBottom = getContentLimitBottom(inner);
+    const bounds = getPageContentBounds(inner);
     let changed = false;
 
     inner.querySelectorAll<HTMLImageElement>("img").forEach((image) => {
       const imageTop = image.getBoundingClientRect().top;
-      const availableHeight = Math.floor(contentBottom - imageTop);
+      const availableHeight = Math.floor((bounds.bottom - imageTop) / bounds.scale);
       if (availableHeight < 120) return;
 
       const currentMaxHeight = parseFloat(image.style.maxHeight || "0");
```

En la comprobación de bloques se comparan alturas visibles entre sí. En la de imágenes, la división por `bounds.scale` vuelve a convertir el espacio visible a la unidad que necesita `image.style.maxHeight`.

### 5. Archivo `EditorHTML/debug/DebugLogger.ts`

19. El informe utiliza el límite inferior compartido y la tolerancia ajustada al zoom. Se elimina su cálculo local antiguo. **Justificación:** el diagnóstico debe informar del mismo límite que aplica el editor. Dejar la fórmula anterior mostraría resultados engañosos incluso después de corregir la paginación.

**Snippets antes/después — punto 19:**

```diff
diff --git a/EditorHTML/debug/DebugLogger.ts b/EditorHTML/debug/DebugLogger.ts
index 390278f..c508ffc 100644
--- a/EditorHTML/debug/DebugLogger.ts
+++ b/EditorHTML/debug/DebugLogger.ts
@@ -1,3 +1,5 @@
+import { getPageContentBounds } from "../pagination/PageContentGeometry";
+
 export type HweDebugEvent = {
   data?: unknown;
   durationMs?: number;
@@ -144,7 +146,8 @@ function getPageReport(page: HTMLElement, index: number): unknown {
     const rect = child.getBoundingClientRect();
     return bottom === null ? rect.bottom : Math.max(bottom, rect.bottom);
   }, null);
-  const limitBottom = inner ? getContentLimitBottom(inner) : null;
+  const bounds = inner ? getPageContentBounds(inner) : null;
+  const limitBottom = bounds?.bottom ?? null;
   const scrollOverflows = inner ? inner.scrollHeight > inner.clientHeight + 1 : false;
 
   return {
@@ -163,7 +166,7 @@ function getPageReport(page: HTMLElement, index: number): unknown {
       : null,
     overflows:
       contentBottom !== null && limitBottom !== null
-        ? contentBottom > limitBottom + 1 || scrollOverflows
+        ? contentBottom > limitBottom + (bounds?.scale ?? 1) || scrollOverflows
         : scrollOverflows,
     page: getRect(page),
     userBlankCount: inner?.querySelectorAll("[data-hwe-user-blank='true']").length ?? 0,
@@ -210,12 +213,6 @@ function getTableReport(table: HTMLTableElement, index: number): unknown {
   };
 }
 
-function getContentLimitBottom(inner: HTMLElement): number {
-  const styles = getComputedStyle(inner);
-  const paddingBottom = parseFloat(styles.paddingBottom) || 0;
-  return inner.getBoundingClientRect().bottom - Math.min(paddingBottom, 12);
-}
-
 function getMeaningfulChildElements(inner: HTMLElement): HTMLElement[] {
   return Array.from(inner.children).filter((child) => {
     if (child.hasAttribute("data-hwe-user-blank")) return true;
```

El informe reutiliza el límite que aplica el resto del editor y mantiene la tolerancia expresada en la escala visible. El bloque eliminado ya no puede producir un límite diferente en el diagnóstico.

## Pruebas añadidas

### 6. Nuevo archivo `tests/page-bottom-margin.test.cjs`

Se añaden 18 casos de prueba con geometría controlada:

- Reserva completa de 20 mm y altura útil de 242 mm en A4 con zoom 0,5; 1; 1,18 y 2,5.
- Detección de párrafos y tablas que invaden 5 mm del margen, aunque todavía no hayan sobrepasado el borde físico de la hoja.
- Coincidencia entre paginador, servicio de disposición e informe de diagnóstico.
- Comprobación de bloques de 242 y 243 mm para distinguir los que caben de los que no.
- Conversión de la altura máxima de imágenes a unidades sin zoom.
- Descuento de bordes y funcionamiento con un margen inferior menor de 12 píxeles.
- Conservación de la densidad de una tabla dividida entre hojas, recálculo después de borrar filas y aislamiento respecto a otro editor.

**Justificación:** son condiciones observables del comportamiento esperado. Las comprobaciones de tablas y texto dentro del margen detectan el problema original antes de que actúe la comprobación de desbordamiento exterior.

JSDOM no calcula posiciones reales: por eso estas pruebas controlan explícitamente las dimensiones y se complementan con Chrome.

**Snippets de las comprobaciones añadidas:**

El helper `setup()` crea una hoja A4 en JSDOM, asigna padding superior de 35 mm e inferior de 20 mm y proporciona rectángulos controlados. `setBottom()` permite mover el final del contenido sin activar artificialmente el desbordamiento exterior. Los siguientes casos utilizan ese helper y las importaciones del archivo de pruebas.

```js
const mm = 96 / 25.4;
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 0.001, `${actual} != ${expected}`);

for (const scale of [0.5, 1, 1.18, 2.5]) {
  test(`reserva los 20 mm completos y 242 mm útiles con zoom ${scale}`, (t) => {
    const { inner, page } = setup(t, scale);
    const bounds = getPageContentBounds(inner);
    near(bounds.bottom, 123 + 277 * mm * scale);
    near(bounds.top, 123 + 35 * mm * scale);
    near(bounds.scale, scale);
    near(getContentHeight(page), 242 * mm);
    near(getContentLimitBottom(inner), bounds.bottom);
  });

  for (const tag of ["p", "table"]) {
    test(`${tag}: detecta invasión del margen antes del borde de hoja con zoom ${scale}`, (t) => {
      const { root, page, inner, setBottom } = setup(t, scale, tag);
      const bounds = getPageContentBounds(inner);
      const layout = new EditorLayoutService();
      assert.equal(pageOverflows(page), false);
      assert.equal(layout.pageOverflows(page), false);
      setBottom(bounds.bottom + 5 * mm * scale);
      assert.equal(pageOverflows(page), true);
      assert.equal(layout.pageOverflows(page), true);
      const report = getHweDebugReport(root).pages[0];
      near(report.limitBottom, bounds.bottom);
      assert.equal(report.overflows, true);
    });
  }

  test(`evalúa bloques e imágenes en el mismo espacio útil con zoom ${scale}`, (t) => {
    const { inner, page } = setup(t, scale);
    const bounds = getPageContentBounds(inner);
    const paginator = Object.create(Paginator.prototype);
    const block = document.createElement("div");
    block.getBoundingClientRect = () => ({ height: 243 * mm * scale });
    assert.equal(paginator.elementFitsOnFreshPage(block, page), false);
    block.getBoundingClientRect = () => ({ height: 242 * mm * scale });
    assert.equal(paginator.elementFitsOnFreshPage(block, page), true);
    const image = document.createElement("img");
    inner.appendChild(image);
    image.getBoundingClientRect = () => ({ top: bounds.top });
    assert.equal(paginator.fitImagesToAvailableSpace(page), true);
    assert.equal(image.style.maxHeight, `${Math.floor(242 * mm)}px`);
  });
}

test("descuenta bordes y conserva la reserva de un margen inferior menor de 12 px", (t) => {
  const { inner } = setup(t, 1);
  inner.style.cssText = "height:200px;box-sizing:content-box;padding:10px 0 5px;border-top:2px solid;border-bottom:3px solid";
  inner.getBoundingClientRect = () => ({ top: 10, bottom: 450, height: 440 });
  const bounds = getPageContentBounds(inner);
  near(bounds.scale, 2);
  near(bounds.top, 34);
  near(bounds.bottom, 434);
  near(bounds.contentHeight, 200);
});

test("una tabla larga conserva la densidad al fragmentarse y la recalcula al borrar filas", (t) => {
  const { root, inner } = setup(t, 1);
  root.className = "pcf-html-editor-root";
  const rows = count => "<tr><td><p>Fila</p></td></tr>".repeat(count);
  inner.innerHTML = `<table class="hwe-word-table" data-hwe-table-flow-id="a"><tbody>${rows(30)}</tbody></table>`;
  root.insertAdjacentHTML("beforeend", `<section class="hwe-page"><div class="hwe-page-inner"><table class="hwe-word-table" data-hwe-table-flow-id="a"><tbody>${rows(30)}</tbody></table></div></section>`);
  const layout = new EditorLayoutService();
  layout.applyOfficialTableWidths(root);
  const tables = [...root.querySelectorAll("table")];
  assert.ok(tables.every(table => table.classList.contains("hwe-table-compact")));
  tables[1].querySelector("tr").remove();
  layout.applyOfficialTableWidths(root);
  assert.ok(tables.every(table => !table.classList.contains("hwe-table-compact")));
  // Another editor's table must not contribute rows, even with the same flow id.
  root.insertAdjacentHTML("afterend", `<main class="pcf-html-editor-root"><table data-hwe-table-flow-id="a"><tbody>${rows(100)}</tbody></table></main>`);
  layout.applyOfficialTableWidths(root);
  assert.ok(tables.every(table => !table.classList.contains("hwe-table-compact")));
});
```

### 7. Nuevo archivo `tests/page-bottom-margin-browser.cjs`

Se verifican ocho escenarios en Chrome: un documento de 55 párrafos y otro de 100 filas de tabla, cada uno con zoom 50 %, 100 %, 118 % y 250 %.

En cada escenario se comprueba:

- Que el documento ocupa varias hojas.
- Que conserva todo el texto, en el mismo orden.
- Que la tabla conserva sus 100 filas.
- Que el padding inferior sigue siendo de 20 mm.
- Que el contenido respeta el límite inferior en la vista de pantalla y al aplicar los estilos de impresión.
- Que el HTML producido por `collectPdfHtml()` conserva el texto y respeta el margen al renderizarse en modo impresión.
- Que no se han producido errores JavaScript en la página de edición.

**Justificación:** probar medidas en un DOM simulado no demuestra que el navegador distribuya correctamente el contenido. La comprobación del HTML preparado para PDF evita confundir la vista del editor con el documento que realmente se exporta.

La prueba espera a que termine la inicialización del editor antes de cargar el HTML. **Justificación:** evita superponer la carga de la muestra inicial con la del documento de prueba.

Se utiliza una tolerancia de 0,3 mm en las mediciones del navegador. **Justificación:** cubre el píxel CSS de tolerancia ya existente, aproximadamente 0,265 mm, y las pequeñas diferencias de redondeo. No se presenta la prueba como una garantía de precisión física de una impresora.

**Snippet de medición en el navegador:**

```js
async function measure(page) {
  return page.evaluate(() => [...document.querySelectorAll(".hwe-page-inner")].map(inner => {
    const css = getComputedStyle(inner);
    const rect = inner.getBoundingClientRect();
    const scale = rect.height / parseFloat(css.height);
    const bottom = Math.max(...[...inner.children].map(child => child.getBoundingClientRect().bottom));
    return { freeMm: (rect.bottom - bottom) / scale / (96 / 25.4),
      paddingMm: parseFloat(css.paddingBottom) / (96 / 25.4) };
  }));
}
```

La distancia libre se divide por la escala y después por `96 / 25.4` para expresarla en milímetros del documento. Así se puede comparar el resultado de los cuatro niveles de zoom con la misma reserva de 20 mm.

**Snippet de carga, conservación del contenido e impresión:**

Este fragmento se ejecuta dentro de cada caso `texto`/`tabla` y cada valor de `zoom` de `[0.5, 1, 1.18, 2.5]`. `page` y `browser` son instancias de Playwright; `html` contiene el documento de prueba y `projectPath` identifica los módulos reales del editor servidos por Vite.

```js
await page.waitForSelector(".hwe-page-inner");
await page.waitForFunction(() => !document.querySelector(".pcf-html-editor-root").inert);
await page.addStyleTag({ content: `.pcf-html-editor-root { --hwe-page-view-zoom:${zoom} !important; }` });
await page.locator("#html-file").setInputFiles({ name: "margen.html", mimeType: "text/html", buffer: Buffer.from(html) });
await page.waitForFunction(() => document.querySelector("#html-file").value === "" &&
  document.querySelectorAll(".hwe-page-inner").length > 1);
await page.evaluate(() => document.fonts.ready);
await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
const expected = html.replace(/<[^>]+>/g, "").replace(/\s+/g, "");
const actual = await page.locator(".hwe-page-inner").allTextContents();
assert.equal(actual.join("").replace(/\s+/g, ""), expected, "conserva todo el texto en orden");
if (kind === "tabla") assert.equal(await page.locator(".hwe-page-inner tr").count(), 100);
for (const media of ["screen", "print"]) {
  await page.emulateMedia({ media });
  const measurements = await measure(page);
  assert.ok(measurements.length > 1);
  for (const [index, item] of measurements.entries()) {
    assert.ok(Math.abs(item.paddingMm - 20) < 0.01);
    assert.ok(item.freeMm >= 19.7, `${media}, hoja ${index + 1}: quedan ${item.freeMm} mm`);
  }
}
const printHtml = await page.evaluate(async projectPath => {
  const { DocumentSerializer } = await import(`/@fs/${projectPath}/EditorHTML/services/DocumentSerializer.ts`);
  const { DEFAULT_PAGE_SETUP } = await import(`/@fs/${projectPath}/EditorHTML/pagination/PageGeometry.ts`);
  const workspace = document.querySelector(".hwe-workspace");
  return new DocumentSerializer().collectPdfHtml(workspace,
    [...workspace.querySelectorAll(".hwe-page")], DEFAULT_PAGE_SETUP, "");
}, projectPath);
const printPage = await browser.newPage();
try {
  await printPage.emulateMedia({ media: "print" });
  await printPage.setContent(printHtml);
  await printPage.evaluate(() => document.fonts.ready);
  for (const [index, item] of (await measure(printPage)).entries()) {
    assert.ok(item.freeMm >= 19.7, `HTML para PDF, hoja ${index + 1}: quedan ${item.freeMm} mm`);
  }
  const exported = await printPage.locator(".hwe-page-inner").allTextContents();
  assert.equal(exported.join("").replace(/\s+/g, ""), expected);
} finally {
  await printPage.close();
}
```

El primer recorrido mide la vista con los medios `screen` y `print`. El segundo renderiza en una página aparte el HTML que devuelve el serializador real. En ambos se comprueba la reserva inferior; también se compara el contenido para detectar pérdidas o cambios de orden.

## Resultado de las verificaciones

1. `node --test tests/*.test.cjs`: **169 pruebas correctas, 0 fallidas**. Incluye las 18 nuevas y 151 casos que ya estaban en el proyecto.
2. `node --test tests/page-bottom-margin-browser.cjs`, con el servidor local en el puerto 5179: **8 escenarios correctos**. Node informa 9 pruebas al incluir también la prueba agrupadora. Verificados pantalla, estilos de impresión y HTML generado para PDF.
3. `node node_modules/typescript/bin/tsc --noEmit --skipLibCheck`: **correcto**.
4. `npm.cmd run lint`: **correcto**.
5. `npm.cmd run build`: **correcto**. Se validó el manifiesto, se generaron tipos y se compiló/empaquetó el control con Webpack.

La comprobación directa `tsc --noEmit`, sin `skipLibCheck`, detectó un error de resolución de declaraciones de Vite/Rollup (`rollup/parseAst`) con la configuración actual del proyecto. No se modificaron dependencias ni la configuración de TypeScript para este trabajo. La comprobación con `skipLibCheck` y la compilación real del PCF terminaron correctamente. La compilación también mostró un aviso de antigüedad de Browserslist; no impidió generar el control.

## Incidencias durante la verificación y cómo se resolvieron

- La primera versión del caso unitario de tabla utilizaba HTML inválido con texto directamente dentro de `table`. El parser movía ese texto fuera de la tabla. Se corrigió la prueba para utilizar `tbody`, `tr` y `td`. No se modificó el código de producción para adaptarlo a ese HTML inválido.
- El primer servidor y el navegador estaban en contextos de red distintos por el aislamiento de ejecución. Se inició un servidor accesible al proceso de Chrome y se repitieron las comprobaciones. No se modificaron los permisos o la configuración de red del proyecto.
- La primera pasada de Chrome confirmó que los párrafos respetaban el margen, pero las tablas compactas crecían tras dividirse. Se inspeccionó el diagnóstico existente del editor y se corrigió el recuento de filas por tabla completa. Después pasaron los ocho escenarios.
- Durante el diagnóstico se probó esperar una estabilización adicional antes de medir. Esa espera no resolvía la tabla y se retiró de la prueba final. El caso final comprueba la geometría tras terminar la carga y dos frames, sin esperar a que la condición de margen se vuelva verdadera.

## Registro documental y archivos generados

### 8. Nuevo archivo `docs/registro-margen-inferior-2026-10-29.md`

Contiene una copia de este registro dentro del proyecto. **Justificación:** mantener las razones, las comprobaciones y el alcance junto al código para futuras sesiones o mantenimiento.

Se entrega también una copia legible en la carpeta de resultados y un parche `correccion-margen-inferior.patch` con las diferencias exactas de esta sesión. **Justificación:** permitir revisar el cambio completo sin mezclarlo con modificaciones anteriores del proyecto. El parche se genera contra las copias de los cuatro archivos existentes guardadas antes de editarlos; incluye los archivos nuevos y este registro.

`npm.cmd run build` regeneró las salidas normales de compilación y los tipos del proyecto. Son productos derivados del código, no decisiones adicionales sobre el formato del documento. No se ha publicado ni instalado el PCF en Power Apps y no se ha creado un commit.

### Ampliación documental: snippets de código

A petición del usuario se han añadido los bloques de código anteriores junto a las justificaciones originales. Se ha incluido el helper de geometría completo, los cuatro diffs de implementación y los snippets de las pruebas de geometría, densidad, navegador e impresión. Se sincronizan la copia de resultados y la copia de `docs`, y se regenera el parche para incluir el documento actualizado. No se modifica el código de producción ni las pruebas durante esta ampliación.

## Cambios previos conservados y límites de esta entrega

El proyecto ya contenía modificaciones o archivos sin seguimiento, entre ellos pruebas de selección entre páginas y cambios en `tests/document-dirty.test.cjs`. No forman parte de esta corrección y se han conservado. El parche de esta entrega no los incluye.

La corrección puede cambiar el reparto en páginas: reservar el margen inferior completo reduce el espacio que antes se consideraba disponible por error; conservar el estilo compacto de una tabla larga también puede modificar su reparto respecto al comportamiento anterior.

Las comprobaciones descritas son locales, con el código real del editor y Chrome. No sustituyen una validación del control desplegado en el formulario de Power Apps ni de un servicio externo que convierta HTML a PDF. La política de imponer siempre los márgenes predeterminados a documentos importados sigue pendiente de una modificación separada.
