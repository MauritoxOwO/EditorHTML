# Unión de títulos con el párrafo siguiente

Fecha: 30 de octubre de 2026.

Proyecto: `C:\Users\Mauro\Documents\New project 3\EditorHTML-document-dirty-push`.

## Qué se ha implementado

Un bloque cuyo estilo contiene el segmento `tit` se mantiene unido al primer párrafo que lo sigue inmediatamente. Se reconocen, por ejemplo, `art-tit`, `art-Tit`, `art-TIT`, `art_tit` y `art-tit-2`. La comparación ignora las diferencias entre mayúsculas y minúsculas mediante la bandera `i`.

Si varios títulos aparecen consecutivamente, se agrupan con el primer párrafo de contenido. Una línea vacía, un salto manual o un bloque que no sea un párrafo interrumpen esta nueva asociación. Si todo el grupo cabe en una página vacía, se mueve entero cuando no cabe en el espacio restante de la página actual.

Si el grupo supera la altura útil de una página completa, se permite dividir el párrafo conservando su inicio junto a los títulos, siempre que haya espacio físico para ese inicio. No se aumentan la altura de la hoja ni sus márgenes. No se ha implementado en esta entrega una política general de dos líneas para todos los párrafos del documento.

Los dos archivos de implementación modificados son `EditorHTML/pagination/KeepTogetherController.ts` y `EditorHTML/pagination/Paginator.ts`. Las rutas que aparecen a continuación son relativas a la carpeta del proyecto indicada arriba.

## De dónde sale el nombre del estilo

El código existente de `styleApi` convierte el nombre de la fila del catálogo en `className`, quitando un punto inicial si el nombre lo contiene. Cuando se aplica el estilo, `ParagraphStyleManager.applyToBlock` lo añade como clase y lo guarda también en `data-hwe-paragraph-style`.

El siguiente fragmento es contexto existente, no una modificación de esta entrega:

```ts
applyToBlock(block: HTMLElement, className: string | null): void {
    this.classNames.forEach((styleClass) => block.classList.remove(styleClass));
    if (!className) {
      block.removeAttribute("data-hwe-paragraph-style");
      return;
    }

    if (this.fontFamilyClassNames.has(className)) {
      this.clearInlineFontFamily(block);
    }
    block.classList.add(className);
    block.setAttribute("data-hwe-paragraph-style", className);
  }
```

`applyToBlock` recibe `block`, que es el elemento HTML del párrafo seleccionado, y `className`, que es el identificador del estilo elegido. `classList.add` añade la clase al párrafo para aplicar su formato. `setAttribute` guarda el mismo nombre en un atributo que podemos consultar después. Estos datos permiten reconocer un título aunque su etiqueta HTML sea `<p>` y no `<h1>`.

No se añade un campo a Dataverse ni una lista manual con todos los títulos. La nueva regla usa los nombres que ya existen en el documento.

## Funciones existentes que reutilizamos

```ts
import { isEditableBlankBlock, isManualPageBreak, isSplittableTextBlock } from "./PaginatorDom";
```

Usamos esta importación en `KeepTogetherController.ts` para reutilizar tres comprobaciones que ya tenía el paginador. La ruta `./PaginatorDom` apunta al módulo de la misma carpeta. No se copia ni se modifica su implementación.

`isEditableBlankBlock` recibe un elemento HTML y comprueba si representa un bloque de texto vacío. `isManualPageBreak` recibe un nodo y reconoce la marca de salto manual. `isSplittableTextBlock` recibe un elemento y comprueba si es un bloque de texto admitido por el divisor. Las tres devuelven un booleano. Compartirlas evita definir criterios distintos para el detector de títulos y para el resto de la paginación.

## 1. Reconocer el segmento tit

Archivo: `EditorHTML/pagination/KeepTogetherController.ts`.

```ts
export function esEstiloTitulo(nombreEstilo: string): boolean {
  return /(?:^|[-_])tit(?:$|[-_])/i.test(nombreEstilo.trim());
}
```

Usamos `esEstiloTitulo` para determinar si el nombre de un estilo corresponde a un título. La función recibe `nombreEstilo` desde `isTitleBlock`. Su contenido es un `string` con el identificador del estilo aplicado o con el nombre de una clase del bloque.

Usamos `trim()` para eliminar los espacios del principio y del final de ese nombre. No elimina espacios internos ni cambia el nombre guardado en el documento: produce el texto que se utiliza en esta comprobación.

Usamos la expresión `(?:^|[-_])tit(?:$|[-_])` para exigir que `tit` aparezca como un segmento independiente. Antes debe haber el principio del nombre, un guion o un guion bajo. Después debe haber el final del nombre o uno de esos separadores. Por eso `art-tit` se reconoce y `art-titular` no.

Usamos la bandera `i` para aceptar `tit`, `Tit` y `TIT`. No se añade `toLowerCase()` porque la bandera ya resuelve esa necesidad.

Usamos `test()` para obtener un resultado booleano: `true` cuando el nombre cumple el patrón y `false` cuando no lo cumple. Se exporta la función para comprobar esta convención directamente en las pruebas.

## 2. Localizar el párrafo dentro de sus contenedores

Archivo: `EditorHTML/pagination/KeepTogetherController.ts`.

```ts
getFlowTextBlock(node: ChildNode | undefined, includeBlank = false): HTMLElement | null {
    if (!node || node.nodeType !== Node.ELEMENT_NODE) return null;
    const element = node as HTMLElement;
    if (isManualPageBreak(element) || element.hidden ||
        element.getAttribute("contenteditable") === "false" ||
        (!includeBlank && element.hasAttribute("data-hwe-user-blank")) ||
        element.matches("[data-hwe-api-header], [data-hwe-dynamic-header], [data-hwe-runtime-page-header]")) {
      return null;
    }
    if (includeBlank && isEditableBlankBlock(element)) return element;
    if (isSplittableTextBlock(element)) return element;
    if (!element.matches("div, section, article")) return null;
    const children = Array.from(element.childNodes).filter((child) => !this.isFormattingWhitespace(child));
    return children.length === 1 ? this.getFlowTextBlock(children[0], includeBlank) : null;
  }
```

Usamos `getFlowTextBlock` para obtener el bloque de texto real que representa un nodo del flujo. La llaman `isTitleBlock`, `findTitleGroupEnd`, `preserveTitleBlankBarriers` y, desde el paginador, `splitOversizedTitleGroup`.

El parámetro `node` es un nodo del DOM, no una cadena con HTML. Puede ser un párrafo, una columna de texto, otro contenedor o un valor ausente. La primera comprobación descarta los valores ausentes y los nodos que no son elementos HTML. La expresión `as HTMLElement` comunica el tipo a TypeScript; no crea ni transforma un elemento en ejecución.

Usamos `isManualPageBreak`, `hidden`, `getAttribute`, `hasAttribute` y `matches` para descartar barreras que la asociación no debe atravesar: saltos manuales, elementos ocultos, contenido protegido y cabeceras administradas por el editor. En las llamadas normales también descartamos las líneas vacías marcadas. Devolvemos `null` para indicar que ese nodo no es un bloque de texto admisible para esta regla.

Usamos `isSplittableTextBlock`, que ya existía, para reconocer un bloque de texto válido y no vacío. Si lo encontramos, lo devolvemos directamente.

Si recibimos un `div`, `section` o `article`, revisamos sus hijos porque el párrafo puede estar dentro de una columna generada o un contenedor importado. `Array.from` convierte `childNodes` en un array para poder usar `filter`. Quitamos de esa lista solamente los comentarios y espacios de formato identificados por `isFormattingWhitespace`.

Solo descendemos cuando queda exactamente un hijo. La llamada recursiva a `getFlowTextBlock` permite atravesar varios contenedores de un solo hijo, pero impide tratar un contenedor con varios bloques como si fuera un único párrafo.

El segundo parámetro, `includeBlank`, vale `false` por defecto. `preserveTitleBlankBarriers` lo envía como `true` para poder localizar una línea vacía antes de la limpieza del flujo. En ese caso usamos `isEditableBlankBlock` para devolverla. La opción se transmite también a la llamada recursiva, para reconocer una línea vacía aunque esté dentro de una columna. El detector de títulos y el detector del párrafo siguiente conservan el valor por defecto y no consideran esa línea vacía como contenido que deba unirse.

## 3. Diferenciar espacios del HTML y líneas vacías del documento

Archivo: `EditorHTML/pagination/KeepTogetherController.ts`.

```ts
private isFormattingWhitespace(node: ChildNode): boolean {
    return node.nodeType === Node.COMMENT_NODE ||
      (node.nodeType === Node.TEXT_NODE && !(node.textContent ?? "").trim());
  }
```

Usamos `isFormattingWhitespace` desde `getFlowTextBlock` y `findTitleGroupEnd`. Recibe un `ChildNode` para comprobar si es un comentario o un nodo de texto que solo contiene espacios.

La condición `Node.COMMENT_NODE` reconoce comentarios como `<!-- formato -->`. La condición `Node.TEXT_NODE` reconoce el texto que puede quedar entre etiquetas cuando el HTML está indentado. `textContent ?? ""` proporciona una cadena vacía si no hay contenido, y `trim()` permite comprobar si solo había espacios.

No descartamos un elemento `<p><br></p>` mediante esta función. Ese elemento representa una línea vacía del documento, no simplemente la indentación del archivo HTML, y rompe la unión automática entre título y párrafo.

## 3 bis. Preservar una línea vacía que separa un título del texto

Archivos: `EditorHTML/pagination/KeepTogetherController.ts` y `EditorHTML/pagination/Paginator.ts`.

```ts
preserveTitleBlankBarriers(container: HTMLElement): void {
    let afterTitle = false;
    Array.from(container.childNodes).forEach((node) => {
      if (this.isFormattingWhitespace(node)) return;
      if (node.nodeType === Node.ELEMENT_NODE) {
        const element = node as HTMLElement;
        if (element.matches("div, section, article") && !isManualPageBreak(element) &&
            !element.matches("[hidden], [contenteditable='false'], [data-hwe-api-header], [data-hwe-dynamic-header], [data-hwe-runtime-page-header]")) {
          this.preserveTitleBlankBarriers(element);
        }
      }
      const block = this.getFlowTextBlock(node, true);
      if (afterTitle && block && isEditableBlankBlock(block)) {
        block.setAttribute("data-hwe-user-blank", "true");
      }
      afterTitle = this.isTitleBlock(node);
    });
  }
```

Usamos `preserveTitleBlankBarriers` desde `Paginator.collectNodesFromIndex`, antes de que `getMeaningfulChildren` filtre los nodos vacíos. El parámetro `container` es el contenedor de contenido de cada página. También lo recibe recursivamente cuando revisamos columnas, grupos anteriores y contenedores importados.

Usamos `afterTitle` para recordar si el nodo anterior era un título. Recorremos una copia de `container.childNodes` con `Array.from(...).forEach(...)`. Los comentarios y espacios de formato no cambian esa relación.

Si encontramos un contenedor `div`, `section` o `article`, lo revisamos por dentro, excepto cuando es un salto manual, un elemento oculto, un bloque protegido o una cabecera administrada. Esto permite preservar también una línea vacía introducida dentro de un grupo que ya existía.

Usamos `getFlowTextBlock(node, true)` para localizar el bloque real, admitiendo en esta llamada los bloques vacíos. Si sigue a un título y `isEditableBlankBlock` confirma que es una línea vacía, añadimos el atributo existente `data-hwe-user-blank="true"`. La limpieza del paginador ya sabe conservar los elementos con ese atributo, aunque procedan de un documento importado. No estamos afirmando quién escribió la línea: reutilizamos la marca de conservación que ya entiende el editor.

Al final de cada iteración actualizamos `afterTitle` con `isTitleBlock(node)`. Por eso no marcamos indiscriminadamente todos los párrafos vacíos del documento. La primera línea vacía conservada ya establece una barrera entre título y contenido.

La nueva llamada dentro de la recogida del flujo queda antes del filtrado existente:

```ts
this.keepTogetherController.preserveTitleBlankBarriers(inner);
getMeaningfulChildren(inner).forEach((node) => {
  nodes.push(...this.flattenFlowNode(node));
});
```

Usamos `inner` porque contiene el contenido de esa página. El orden de estas llamadas es necesario: si se filtrara primero, una línea vacía importada sin marca podría desaparecer y el título parecería estar pegado al párrafo. Esta marca se elimina en la comparación de cambios por el normalizador existente; no altera por sí sola el estado dirty.

## 4. Consultar el estilo del bloque real

Archivo: `EditorHTML/pagination/KeepTogetherController.ts`.

```ts
private isTitleBlock(node: ChildNode | undefined): boolean {
    const block = this.getFlowTextBlock(node);
    if (!block) return false;
    const appliedStyle = block.getAttribute("data-hwe-paragraph-style");
    return appliedStyle !== null
      ? esEstiloTitulo(appliedStyle)
      : Array.from(block.classList).some(esEstiloTitulo);
  }
```

Usamos `isTitleBlock` desde `findTitleGroupEnd` para comprobar el bloque actual y los posibles títulos consecutivos. Recibe un nodo del flujo y llama a `getFlowTextBlock` para llegar al párrafo real, aunque este se encuentre dentro de una columna.

Usamos `getAttribute("data-hwe-paragraph-style")` para recuperar el nombre del estilo que guardó `ParagraphStyleManager.applyToBlock`. Si el atributo existe, enviamos su valor a `esEstiloTitulo`.

Damos prioridad a ese atributo sobre las clases del elemento. Esto evita que una clase de título residual haga prevalecer una asociación cuando el estilo aplicado actualmente es de cuerpo de texto.

Cuando el atributo no existe, usamos `Array.from(block.classList)` para obtener los nombres de clase como un array. `some(esEstiloTitulo)` llama a la función con cada nombre y devuelve `true` en cuanto encuentra uno que contiene el segmento `tit`. Esta alternativa permite reconocer documentos HTML que tienen la clase aplicada pero no el atributo interno del editor.

## 5. Encontrar el final de una cadena de títulos y su párrafo

Archivo: `EditorHTML/pagination/KeepTogetherController.ts`.

```ts
private findTitleGroupEnd(nodes: ChildNode[], start: number): number {
    if (!this.isTitleBlock(nodes[start])) return start;
    let index = start + 1;
    while (index < nodes.length) {
      if (this.isFormattingWhitespace(nodes[index]) || this.isTitleBlock(nodes[index])) {
        index++;
        continue;
      }
      const paragraph = this.getFlowTextBlock(nodes[index]);
      return paragraph && paragraph.matches("p, div") ? index : start;
    }
    return start;
  }
```

Usamos `findTitleGroupEnd` desde `groupFlowNodes`. Recibe `nodes`, que es el array de nodos que se va a paginar, y `start`, que es la posición del nodo que estamos revisando.

Primero comprobamos si ese nodo es un título. Si no lo es, devolvemos `start`, que significa que no se ha encontrado una asociación nueva a partir de esa posición.

Usamos `while` para recorrer los nodos siguientes. Saltamos los espacios de formato y seguimos avanzando si encontramos otro título. Así podemos mantener juntos un título principal, uno o varios subtítulos y el primer párrafo de contenido.

Cuando encontramos un nodo distinto, lo enviamos a `getFlowTextBlock`. Si el bloque real es un párrafo `p` o un bloque de texto `div`, devolvemos su posición como final del grupo. Los títulos con la convención `tit` ya se han descartado como párrafo final en la comprobación anterior del bucle.

Si el nodo es una línea vacía, una tabla, una imagen, un salto manual u otra barrera, no encontramos un párrafo válido y devolvemos `start`. También devolvemos `start` si se termina el documento sin encontrar contenido posterior. No se crea una asociación incompleta con un párrafo inexistente.

## 6. Crear el grupo que se paginará como una unidad

Archivo: `EditorHTML/pagination/KeepTogetherController.ts`.

```ts
const titleGroupEnd = this.findTitleGroupEnd(nodes, index);
if (titleGroupEnd > index) {
  const group = document.createElement("div");
  group.className = "hwe-keep-together";
  group.setAttribute("data-hwe-keep-together", "true");
  group.setAttribute("data-hwe-generated-wrapper", "true");
  group.setAttribute("data-hwe-title-group", "true");
  nodes.slice(index, titleGroupEnd + 1).forEach((node) => group.appendChild(node));
  grouped.push(group);
  index = titleGroupEnd + 1;
  continue;
}
```

Este fragmento se ejecuta dentro de `groupFlowNodes`. El paginador llama a ese método desde `repaginateFromIndex`, después de recoger los nodos y reunir las continuaciones del documento.

Usamos `findTitleGroupEnd(nodes, index)` para conocer hasta dónde llega la asociación. Solo creamos el grupo cuando el resultado es mayor que `index`, es decir, cuando se ha encontrado contenido posterior al título.

Usamos `document.createElement("div")` para crear un contenedor de disposición. La clase `hwe-keep-together` y el atributo `data-hwe-keep-together="true"` permiten reutilizar el mecanismo del paginador que mueve un bloque completo.

Usamos `data-hwe-generated-wrapper="true"` para identificar el contenedor como un elemento temporal generado por el editor. Los serializadores existentes eliminan estos contenedores al guardar y al preparar el HTML para PDF. Por eso no se ha tenido que modificar el serializador.

Usamos `data-hwe-title-group="true"` para distinguir esta agrupación de las agrupaciones anteriores de imágenes. La distinción es necesaria para reconstruir la asociación después de editar y para tratar los grupos que superan una página.

Usamos `slice(index, titleGroupEnd + 1)` para obtener desde el título inicial hasta el párrafo, ambos incluidos. El `+ 1` es necesario porque el límite final de `slice` no está incluido.

Usamos `forEach` y `appendChild` para introducir esos nodos en el mismo orden dentro del grupo. `appendChild` mueve los nodos existentes; no copia ni duplica su contenido.

Usamos `grouped.push(group)` para añadir la nueva unidad a la lista que se paginará. Después avanzamos `index` hasta el siguiente nodo pendiente y ejecutamos `continue` para no volver a procesar los elementos que ya están dentro del grupo.

Las ramas anteriores que mantienen textos con imágenes o tablas siguen después de este fragmento. No se han sustituido sus criterios por la convención `tit`.

## 7. Reconocer un grupo de títulos

Archivo: `EditorHTML/pagination/KeepTogetherController.ts`.

```ts
isTitleGroup(element: HTMLElement): boolean {
    return this.isKeepTogetherGroup(element) &&
      element.getAttribute("data-hwe-title-group") === "true";
  }
```

Usamos `isTitleGroup` desde `Paginator.flattenFlowNode` y `Paginator.moveOverflowPiece`. Recibe el elemento HTML del posible grupo.

Primero llama a `isKeepTogetherGroup`, que ya existía, para confirmar que es una unidad de disposición. Después exige que `data-hwe-title-group` tenga el valor `"true"`.

Usamos `&&` porque ambas condiciones deben cumplirse. Así no aplicamos la reconstrucción o la división especial de títulos a cualquier agrupación de imágenes que utilice la misma clase CSS general.

## 8. Reconstruir la asociación después de editar o cambiar estilos

Archivo: `EditorHTML/pagination/Paginator.ts`.

```ts
private flattenFlowNode(node: ChildNode): ChildNode[] {
    if (node.nodeType !== Node.ELEMENT_NODE) return [node];

    const element = node as HTMLElement;
    // Rebuild title associations after edits or style changes, and before merging continuations.
    if (this.keepTogetherController.isTitleGroup(element)) {
      return getMeaningfulChildren(element).flatMap((child) => this.flattenFlowNode(child));
    }
    if (!this.isSplittableContainer(element) || shouldPreserveContainerShell(element)) {
      return [node];
    }

    const children = getMeaningfulChildren(element);
    if (children.length === 0) return [];

    return children.flatMap((child) => this.flattenFlowNode(child));
  }
```

Usamos `flattenFlowNode` desde `collectNodesFromIndex`, que recoge el contenido que necesita volver a paginarse. La nueva rama reconoce los grupos de títulos antes de aplicar los criterios generales de conservación de contenedores.

Usamos `getMeaningfulChildren` para recuperar sus hijos de contenido y `flatMap` para reunir el resultado de volver a procesarlos. La llamada recursiva a `flattenFlowNode` permite seguir recogiendo el flujo a través de los contenedores existentes.

Esta extracción se realiza antes de reunir continuaciones y volver a llamar a `groupFlowNodes`. Por eso la unión se reconstruye usando los estilos y la estructura actuales: si el usuario quita el estilo de título, el grupo anterior no queda fijado para siempre; si lo vuelve a aplicar, se puede crear otra vez.

El resto del método es código existente. Mantiene el tratamiento habitual de otros contenedores y no convierte cualquier `div` en un grupo de títulos.

## 9. Tratar un grupo que supera una página completa

Archivo: `EditorHTML/pagination/Paginator.ts`.

```ts
if (this.keepTogetherController.isTitleGroup(lastChild as HTMLElement) &&
    this.splitOversizedTitleGroup(lastChild as HTMLElement, targetInner, page)) {
  return true;
}
this.keepTogetherController.unwrapGeneratedGroup(lastChild as HTMLElement);
return true;
```

Este fragmento está dentro de `moveOverflowPiece`, en la rama que ya trata los grupos que deben mantenerse juntos. Se alcanza después de intentar trasladar el grupo cuando hay otros bloques en la página. En este punto, el grupo es el único contenido significativo de su página y sigue desbordando.

Usamos `isTitleGroup` para limitar el nuevo tratamiento a las asociaciones de títulos. Si lo es, llamamos a `splitOversizedTitleGroup` con el grupo, el contenedor de la página siguiente y la página actual.

Si la función consigue mover la continuación, devolvemos `true` para comunicar al paginador que ha habido progreso. Si no puede hacerlo, se conserva la alternativa existente de deshacer la agrupación. Esto evita intentar trasladar indefinidamente un grupo que nunca podría caber entero en una hoja.

```ts
private splitOversizedTitleGroup(
    group: HTMLElement,
    targetInner: HTMLElement,
    page: HTMLElement
  ): boolean {
    const paragraphContainer = group.lastElementChild as HTMLElement | null;
    const paragraph = paragraphContainer && this.keepTogetherController.getFlowTextBlock(paragraphContainer);
    if (!paragraphContainer || !paragraph) return false;

    // Clone only the paragraph's container path, preserving its text width and flow identities.
    let source = paragraphContainer;
    let continuation: HTMLElement | null = null;
    let target: HTMLElement = targetInner;
    while (source !== paragraph) {
      ensureFlowIdentity(source, CONTAINER_FLOW_ID);
      const clone = source.cloneNode(false) as HTMLElement;
      if (!continuation) {
        continuation = clone;
        targetInner.insertBefore(clone, targetInner.firstChild);
      } else {
        target.appendChild(clone);
      }
      target = clone;
      source = source.firstElementChild as HTMLElement;
    }

    const moved = this.textBlockSplitter.splitTextBlock(paragraph, target, page);
    if (!moved) {
      continuation?.remove();
      return false;
    }
    if (!paragraph.parentNode) {
      // The headings leave no room even for the paragraph's first line.
      if (paragraphContainer !== paragraph) paragraphContainer.remove();
      hweDebugLog("paginator.titleGroup.noRoomForParagraph", {});
    }
    return true;
  }
```

Usamos `splitOversizedTitleGroup` desde `moveOverflowPiece`. Recibe `group`, el contenedor que reúne títulos y párrafo; `targetInner`, el contenedor de contenido de la página siguiente; y `page`, la página cuyo desbordamiento estamos resolviendo.

Usamos `lastElementChild` porque el último bloque del grupo creado es el párrafo de contenido. Después llamamos a `getFlowTextBlock` para obtener su elemento de texto real. Si no encontramos alguno de esos elementos, devolvemos `false` y no intentamos una división sin un párrafo válido.

Usamos `source` para recorrer los contenedores que rodean al párrafo, `target` para recordar dónde colocar la continuación y `continuation` para conservar una referencia al primer contenedor nuevo. La variable `continuation` también permite retirar esa estructura si la división no consigue mover texto.

Usamos `while (source !== paragraph)` para copiar únicamente el camino de contenedores que lleva al párrafo. `ensureFlowIdentity` asigna o conserva su identificador de flujo antes de copiarlo; el paginador podrá reconocer después que ambos fragmentos pertenecen al mismo contenido.

Usamos `cloneNode(false)` para copiar cada contenedor con sus atributos, pero sin copiar sus hijos ni duplicar el texto. Esto conserva, por ejemplo, el contenedor que limita el texto a 140 mm. El primer contenedor se coloca al principio de la página siguiente con `insertBefore`; los interiores se añaden al contenedor anterior con `appendChild`.

Usamos `source.firstElementChild` para descender por el mismo camino que validó `getFlowTextBlock`. Al terminar, `target` es el lugar donde debe introducirse el texto que no cabe.

Usamos el divisor existente `TextBlockSplitter.splitTextBlock(paragraph, target, page)` para trasladar el final del párrafo hasta que la página actual deje de desbordar. Los títulos permanecen en el grupo inicial. No se crean palabras nuevas ni se recalcula el ancho del texto con una regla distinta.

Si el divisor no ha podido mover nada, usamos `continuation?.remove()` para retirar el contenedor vacío que acabamos de crear. El operador `?.` permite hacerlo también cuando no fue necesario crear un contenedor, sin intentar llamar a `remove()` sobre `null`.

Si el párrafo completo ha tenido que salir porque los títulos no dejan sitio ni para su primera línea, retiramos el contenedor de párrafo que haya quedado vacío y registramos `paginator.titleGroup.noRoomForParagraph` en el diagnóstico. Es una excepción física: no se amplía la hoja para fingir que el grupo cabe. Esta entrega no promete una reserva de dos líneas cuando las dimensiones del título y del párrafo hacen imposible cumplirla.

Devolvemos `true` cuando se ha movido contenido para que el paginador pueda continuar con la siguiente comprobación de desbordamiento.

## 10. Pruebas añadidas y qué demuestra cada fragmento

Se añade `tests/title-keep-together.test.cjs` con 28 pruebas unitarias y de integración DOM. Comprueban nombres válidos y falsos positivos, el atributo de estilo aplicado, las columnas generadas, los títulos consecutivos, las barreras, la conservación de una línea vacía importada, los títulos largos, la compatibilidad con imágenes y tablas y la limpieza al guardar o preparar HTML para PDF.

```js
for (const name of ["window", "document", "Node", "HTMLElement", "DOMParser"]) global[name] = dom.window[name];
  global.getComputedStyle = dom.window.getComputedStyle.bind(dom.window);
  const root = document.querySelector("main");
  const controller = new KeepTogetherController();
  const group = () => {
    const nodes = controller.groupFlowNodes([...root.childNodes]);
    root.replaceChildren(...nodes);
    return nodes;
  };
  return { root, controller, group };
}

for (const name of ["tit", "art-tit", "art-Tit", "art-TIT", "art_tit", "tit-general", "art-tit-2", "  art-TIT  "]) {
  test(`reconoce el estilo de título ${JSON.stringify(name)}`, () => assert.equal(esEstiloTitulo(name), true));
}
for (const name of ["", "art-texto", "art-titular", "arttit", "art-tit2"]) {
  test(`no confunde ${JSON.stringify(name)} con un título`, () => assert.equal(esEstiloTitulo(name), false));
}
```

Usamos estos dos recorridos para enviar distintos nombres directamente a `esEstiloTitulo`. En cada iteración, `name` contiene una cadena de la lista y se utiliza como parámetro de la función. `assert.equal` compara el resultado con el booleano esperado. Así queda comprobado que se aceptan las variantes con mayúsculas y que no se confunde `titular` con el segmento `tit`.

```js
test("agrupa el título aplicado y el párrafo inmediato sin incluir los demás", t => {
  const app = setup(t, '<p data-hwe-paragraph-style="art-TIT">Título</p><p>Cuerpo</p><p>Después</p>');
  const nodes = app.group();
  assert.equal(nodes.length, 2);
  assert.equal(app.controller.isTitleGroup(nodes[0]), true);
  assert.equal(nodes[0].children.length, 2);
  assert.equal(nodes[0].textContent, "TítuloCuerpo");
  assert.equal(nodes[1].textContent, "Después");
});
```

Usamos `setup` para crear un DOM de prueba con un título, su párrafo y un tercer bloque. `app.group()` ejecuta el `groupFlowNodes` real. Comprobamos que solo se forma una unidad con los dos primeros elementos y que el tercero permanece fuera. La comparación de `textContent` también verifica que se conserva el orden del contenido.

Se añade `tests/title-keep-together-browser.cjs` para ejecutar cuatro escenarios completos en Chrome, uno por cada zoom de 50 %, 100 %, 118 % y 250 %. Cada escenario comprueba el traslado del grupo, quitar y volver a aplicar el estilo mediante `ParagraphStyleManager`, rellenar de nuevo una página cuando aumenta el espacio disponible, títulos consecutivos, saltos manuales, líneas vacías marcadas e importadas, un párrafo que necesita varias páginas y el HTML real preparado para PDF.

```js
assert.equal(result.together.title, result.together.body);
assert.ok(result.together.title > 0);
assert.ok(result.together.fits && result.together.text);
assert.equal(result.removed.groups, 0);
assert.equal(result.removed.title, 0);
assert.ok(result.removed.text && result.reapplied && result.compacted);
assert.equal(result.chain.count, 3);
assert.equal(result.chain.title, result.chain.body);
assert.ok(result.chain.title > 0 && result.chain.intact && result.chain.fits);
assert.equal(result.manual.groups, 0);
assert.ok(result.manual.separate && result.blank && result.importedBlank);
assert.ok(result.oversized.pages > 2);
assert.equal(result.oversized.title, result.oversized.body);
assert.ok(result.oversized.lines >= 2);
assert.ok(result.oversized.fits && result.oversized.intact && result.oversized.stable);
assert.ok(result.dirtyStable && result.cleanSaved && result.cleanPrint);
```

Usamos `pageIndex` dentro de la prueba para obtener la página en la que está cada bloque. Las primeras comparaciones exigen que título y párrafo terminen en la misma página y que hayan pasado a una página posterior cuando el hueco inicial no permite alojarlos juntos.

Las comprobaciones de `removed`, `reapplied` y `compacted` comprueban lo que ocurre al quitar el estilo, volver a aplicarlo y dejar más espacio disponible. Esto prueba la reconstrucción de los grupos y evita aceptar una asociación que solo funcionase durante la carga inicial.

En `oversized` comprobamos que el documento necesita varias páginas, que el título comparte página con el comienzo del párrafo y que ese comienzo contiene al menos dos líneas en el caso probado. También exigimos conservación del texto, respeto del límite inferior y estabilidad después de volver a paginar.

Las comprobaciones de `dirtyStable`, `cleanSaved` y `cleanPrint` verifican que la estructura temporal no altera el estado de cambios pendientes ni se filtra a los documentos serializados.

El navegador usa los módulos TypeScript reales del proyecto y su CSS. El archivo de prueba los transpila y reúne en memoria para cargarlos en Chrome. Ese empaquetado sirve al entorno de pruebas; no sustituye el código de producción ni cambia la compilación del PCF. No necesita iniciar un servidor web ni conectarse a Dataverse.

## Verificación realizada

- `node --test tests/*.test.cjs`: 197 pruebas correctas, 0 fallidas. Son las 169 anteriores más las 28 nuevas.
- `node --test tests/title-keep-together-browser.cjs`: los cuatro escenarios por zoom son correctos; Node informa cinco pruebas al contar también la agrupadora.
- `node node_modules/typescript/bin/tsc --noEmit --skipLibCheck`: correcto.
- `npm.cmd run lint`: correcto.
- `npm.cmd run build`: correcto; manifiesto validado y control compilado y empaquetado con Webpack.

La compilación muestra el aviso de antigüedad de los datos de Browserslist que ya existía. No ha impedido generar el control ni se han actualizado dependencias para este cambio.

## Alcance de los archivos y del registro

Además de los dos archivos de implementación se añaden las dos pruebas y una copia de este documento en `docs/registro-union-titulos-2026-10-30.md`.

El archivo `union-titulos-parrafo.patch` registra las diferencias de esta entrega respecto al estado de los archivos justo antes de implementar la unión. Ese punto de partida ya incluía la corrección anterior del margen inferior. Por tanto, este nuevo parche conserva esa corrección y no intenta volver a aplicarla. El registro y el parche anteriores del margen se mantienen como entregas separadas.

La compilación regenera sus salidas habituales y los tipos derivados. No se ha publicado el control en Power Apps ni se ha creado un commit. Las modificaciones que ya tenía el proyecto y que no pertenecen a esta entrega se han conservado.

Las explicaciones de este registro acompañan al código añadido: indican el origen de los parámetros, el propósito de las llamadas y las razones de las comprobaciones. Los bloques identificados como contexto existente sirven para explicar el recorrido y no se presentan como cambios nuevos.
