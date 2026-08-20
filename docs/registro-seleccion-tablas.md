# Registro de cambios: seleccion de tablas, filas y columnas

## Objetivo

Esta iteracion amplia el selector de tablas existente con dos capacidades:

1. Evitar que el icono desaparezca durante el breve recorrido del puntero entre la tabla y el propio icono.
2. Permitir seleccionar la tabla completa, una fila o una columna para aplicar los comandos de formato compatibles.

La implementacion no introduce controles ni clases de seleccion persistentes dentro del HTML del documento. La interfaz contextual y los resaltados se montan directamente bajo el root del editor, fuera de `.hwe-page` y `.hwe-page-inner`; por ello no forman parte de las paginas que clona `DocumentSerializer` para guardar HTML o producir el documento PDF.

## Diseno de interaccion

El flujo elegido es deliberadamente progresivo:

1. El usuario mueve el puntero sobre una celda.
2. Aparece el selector de esquina existente.
3. Al pulsarlo se selecciona la tabla completa y aparece la barra contextual `Tabla | Fila | Columna`.
4. `Fila` selecciona la fila de la ultima celda apuntada.
5. `Columna` selecciona la columna visual que ocupaba el puntero dentro de esa celda.
6. Mientras el modo `Fila` o `Columna` permanece activo, el alcance no cambia al mover el puntero hacia las barras de herramientas.
7. Para cambiar el alcance se pulsa directamente cualquier celda de otra fila o columna. No es necesario volver a pulsar el boton contextual.
7. `Escape` o un clic fuera de la tabla cancelan la seleccion.

Esta solucion evita llenar los margenes de la tabla con un boton por cada fila y columna. Tambien reduce la dependencia de controles pequenos y deja abierta una evolucion futura hacia selectores visuales tipo Word.

## Archivos modificados

### `EditorHTML/controllers/TableSelectionController.ts`

Responsabilidad: coordinar el hover, el selector de esquina, la barra contextual, el estado de seleccion, los resaltados y la aplicacion de formato.

Cambios principales:

- Se incorporan tres modos de seleccion: `table`, `row` y `column`.
- Se recuerda la celda activa y la columna visual situada bajo el puntero.
- Un clic sobre una celda bloquea el objetivo explicito. El hover deja de poder sustituir esa celda mientras se desplaza el puntero hacia la barra contextual o la barra principal.
- Una vez seleccionado el modo fila o columna, el alcance permanece fijado hasta que se pulsa otra celda, se selecciona otro modo o se cancela con `Escape`.
- La barra muestra una indicacion contextual: `Pulsa una celda de otra fila` o `Pulsa una celda de otra columna`.
- Se mantiene una tolerancia de 250 ms antes de ocultar el selector de esquina.
- Entrar en el selector o en la barra contextual cancela el ocultado pendiente.
- Se crea la barra contextual una sola vez y se reutiliza durante toda la vida del editor.
- Los botones reflejan el modo activo mediante `aria-pressed` y estado visual.
- `Fila` y `Columna` quedan deshabilitados si todavia no existe una celda objetivo.
- Los resaltados se reconstruyen cuando cambia la seleccion, hay scroll, cambia el tamano de la ventana o se repagina el contenido.
- Los comandos de negrita, subrayado y tamano de fuente pasan a operar sobre el alcance seleccionado, no siempre sobre la tabla completa.
- La seleccion de una fila de cabecera se refleja en las copias de esa cabecera creadas al dividir una tabla entre paginas.
- La seleccion de columna se aplica a todos los fragmentos que comparten `data-hwe-table-flow-id`.
- `destroy()` elimina timers, overlays, selector y barra contextual para no dejar listeners o nodos huerfanos.

Estado interno relevante:

```text
hoveredTable          Tabla bajo el puntero; posiciona el selector de esquina.
activeCell            Ultima celda apuntada; determina la fila objetivo.
activeColumnIndex     Columna visual bajo el puntero; determina la columna objetivo.
isCellTargetLocked    Indica que el objetivo procede de un clic y no debe cambiar por hover.
selectedTable         Tabla o flujo logico actualmente seleccionado.
selectedRow           Fila seleccionada cuando el modo es `row`.
selectedColumnIndex   Indice visual seleccionado cuando el modo es `column`.
selectionMode         Alcance activo: tabla, fila o columna.
```

El controlador conserva referencias DOM solo durante la sesion de edicion. No escribe identificadores de seleccion en el documento.

### `EditorHTML/dom/TableGrid.ts`

Responsabilidad: convertir el DOM de una tabla en una cuadrícula visual utilizable por la seleccion de columnas.

El indice `cellIndex` nativo no es suficiente cuando hay `rowspan` o `colspan`. El nuevo modulo recorre las filas y mantiene un mapa de posiciones ocupadas:

- `getTableGridCells(table)` devuelve, para cada celda, la columna visual inicial y final.
- `getColumnIndexAtClientX(table, cell, clientX)` determina que subcolumna se esta apuntando dentro de una celda con `colspan`.
- `getCellsForColumn(table, columnIndex)` devuelve todas las celdas que intersectan la columna solicitada.

Una celda con `colspan="2"` pertenece a ambas columnas visuales. Por ello se resaltara y formateara completa al seleccionar cualquiera de esas columnas; no se intenta dar formato a una fraccion imposible de la celda.

### `EditorHTML/css/editor.css`

Responsabilidad: presentar la barra contextual y los overlays de seleccion.

Se agregan estilos para:

- `.hwe-table-selection-toolbar`: barra contextual flotante.
- Botones activos, hover, foco visible y estado deshabilitado.
- `.hwe-table-selection-overlay`: rectangulo visual que no recibe eventos del puntero.
- Variantes para tabla, fila y columna con diferentes intensidades de resaltado.

Los overlays usan `pointer-events: none`; no bloquean la edicion de las celdas ni modifican el contenido de la tabla.

### `web/public/fixtures/table-selection-spans.html`

Fixture manual para comprobar la seleccion sobre una tabla con `rowspan` y `colspan`. Se mantiene como caso reproducible dentro del sandbox y puede abrirse con `?fixture=table-selection-spans`.

### `web/src/main.ts`

El harness reconoce `?fixture=table-selection-10-pages` y genera una tabla larga con 135 filas, cabecera repetible y cuatro columnas identificables. La medicion real del paginador local situa unas 15 filas por pagina y produce diez fragmentos. La generacion en JavaScript evita mantener un archivo HTML enorme y permite ajustar de forma controlada el numero de filas si cambia la altura util de pagina.

## Comportamiento con tablas partidas entre paginas

El editor ya relaciona los fragmentos mediante `data-hwe-table-flow-id`. La nueva seleccion reutiliza `getTableFlowFragments`:

- `Tabla`: resalta todos los fragmentos.
- `Columna`: resuelve la misma columna visual en cada fragmento.
- `Fila de cuerpo`: actua sobre la fila real, que el paginador mueve entre fragmentos sin clonar.
- `Fila de cabecera repetida`: localiza la fila equivalente dentro de cada `thead` clonado.

## Formato admitido en esta iteracion

El alcance de tabla, fila o columna se aplica a las operaciones que el controlador ya interceptaba:

- Negrita.
- Subrayado.
- Tamano de fuente.

Los demas comandos siguen usando la seleccion nativa del navegador. No se han agregado operaciones estructurales de columnas ni se ha cambiado el comportamiento de insertar/eliminar filas.

## Accesibilidad y limpieza

- La barra tiene `role="toolbar"` y nombre accesible.
- Cada boton publica su seleccion mediante `aria-pressed`.
- Los overlays tienen `aria-hidden="true"`.
- Los controles usan `contenteditable="false"`.
- `Escape` limpia el estado.
- Al destruir el editor se cancelan timers y se eliminan todos los elementos auxiliares.

## Decisiones descartadas en esta fase

No se han creado selectores individuales en el margen de cada fila y columna porque multiplicarian los objetivos pequenos de hover, que es precisamente el problema que origina esta iteracion. Tampoco se anaden clases a `tr`, `td` o `th`: aunque pudieran eliminarse al serializar, un overlay externo garantiza de forma estructural que el estado de interfaz no llegue al HTML ni al PDF.

## Verificacion ejecutada

1. La compilacion TypeScript aislada de `TableFlow.ts`, `TableGrid.ts` y `TableSelectionController.ts` finaliza sin errores.
2. ESLint finaliza sin errores en los dos archivos TypeScript nuevos o modificados por esta iteracion.
3. `git diff --check` no detecta errores de whitespace en el alcance del cambio.
4. El build del harness Vite finaliza correctamente: 46 modulos transformados.
5. En la tabla sencilla se comprueba que el selector sobrevive al recorrido desde una celda hasta el icono.
6. Al pulsar el icono aparece la barra contextual y un overlay de tabla.
7. El modo fila crea un unico overlay para la fila objetivo.
8. El modo columna sobre la primera columna crea cuatro overlays, uno por cada celda de cabecera/cuerpo que pertenece a ella.
9. Aplicar negrita con esa columna seleccionada modifica exclusivamente la primera columna; las otras dos conservan su peso anterior.
10. `Escape` oculta la barra y elimina todos los overlays.
11. Se verifica que selector, barra y overlays son hijos de `.pcf-html-editor-root` y que ninguno esta dentro de `.hwe-page`.
12. La consola del navegador no registra errores durante estas operaciones.
13. Con `?fixture=table-selection-spans`, apuntar a la mitad derecha de una celda `colspan="2"` selecciona la tercera columna logica. Los cinco objetivos obtenidos son `Detalle`, `Responsable`, `Equipo A`, `Equipo B` y `Salida compartida`.
14. Con `?fixture=table-selection-10-pages`, 135 filas producen exactamente 10 paginas y 10 fragmentos asociados a un unico `data-hwe-table-flow-id`.
15. Seleccionar la primera columna de esa tabla genera 145 objetivos: 135 celdas de cuerpo y 10 cabeceras repetidas.
16. Aplicar negrita modifica los 145 objetivos de la primera columna y ninguna celda de las otras tres columnas; la tabla permanece en 10 paginas despues del rebalanceo.
17. Se selecciona la fila `080` de la pagina 6 y despues la fila `095` de la pagina 7. En ambos casos queda un unico overlay y coincide geometricamente con la fila elegida.
18. Se reproduce el recorrido desde la fila `001` hasta la barra principal: el modo `Fila`, el panel contextual y el overlay permanecen activos durante todo el movimiento.
19. Al pulsar negrita, las cuatro celdas de `001` reciben `font-weight: 700`; las cuatro celdas de `002` conservan su estilo anterior. Esto confirma que el comando es interceptado por la seleccion de fila y no cae en `document.execCommand` sobre la celda nativa.

El type-check global del repositorio sigue informando dos problemas preexistentes ajenos a esta iteracion: `EditorViewController.ts` referencia una variable `context` no definida y las declaraciones de Tesseract requieren el tipo global `Buffer` mientras el `tsconfig` limita `types` a Power Apps. No se modifican esos comportamientos en este trabajo; el build funcional del sandbox si termina correctamente.
