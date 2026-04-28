import { Paginator } from "../Orquestador/Paginator";
import { Toolbar } from "../Resize/Toolbar";
import { fetchHtmlFromFileField, saveHtmlToFileField } from "../execCommand/fileApi";
type PcfContext = ComponentFramework.Context<IInputs>;

export class EditorComponent {
  private container: HTMLElement;

  private root!: HTMLElement;
  private workspace!: HTMLElement;
  private statusMsg!: HTMLElement;
  private pageCountEl!: HTMLElement;

  private pages: HTMLElement[] = [];

  private paginator!: Paginator;
  private toolbar!: Toolbar;

  private baseUrl: string;
  private entityName: string;
  private entityId: string;
  private fieldName: string;

  private isDirty = false;

  constructor(container: HTMLElement, context: PcfContext) {
    this.container = container;

    const page = (context as unknown as {
      page: { getClientUrl: () => string; entityId: string };
    }).page;

    this.baseUrl   = page.getClientUrl();
    this.entityId  = page.entityId ?? "";
    this.entityName = "mcdev_htmldevtests";
    this.fieldName  = "mcdev_htmlarchivooriginal";
  }

  async init(): Promise<void> {
    this.buildShell();
    this.paginator = new Paginator(
      (html?: string) => this.createPageElement(html),
      (pages: HTMLElement[]) => this.onPagesChanged(pages)
    );
    await this.loadContent();
  }

  // ── Shell ─────────────────────────────────────────────────────

  private buildShell(): void {
    this.container.innerHTML = "";
    this.container.style.cssText =
      "width:100%;height:100%;overflow:hidden;display:flex;flex-direction:column;";

    this.root = document.createElement("div");
    this.root.className = "hwe-root";

    this.toolbar = new Toolbar();
    const toolbarEl = this.toolbar.build();
    this.toolbar.getSaveButton().addEventListener("click", () => this.save());
    this.root.appendChild(toolbarEl);

    this.workspace = document.createElement("div");
    this.workspace.className = "hwe-workspace";
    this.root.appendChild(this.workspace);

    const statusBar = document.createElement("div");
    statusBar.className = "hwe-statusbar";

    this.pageCountEl = document.createElement("span");
    this.pageCountEl.textContent = "Páginas: —";
    statusBar.appendChild(this.pageCountEl);

    this.statusMsg = document.createElement("span");
    this.statusMsg.className = "hwe-status-msg";
    statusBar.appendChild(this.statusMsg);

    this.root.appendChild(statusBar);
    this.container.appendChild(this.root);
  }

  // ── Carga ─────────────────────────────────────────────────────

  private async loadContent(): Promise<void> {
    this.setStatus("Cargando contenido...", "saving");
    try {
      const html = await fetchHtmlFromFileField(
        this.baseUrl, this.entityName, this.entityId, this.fieldName
      );

      // 1. Renderizar todo en una sola página provisional
      this.renderInitial(html);

      // 2. Esperar dos frames: primero pinta, segundo calcula layout
      await this.waitFrames(2);

      // 3. Paginar por desbordamiento
      await this.paginateContent();

      this.setStatus("", "");
    } catch (err) {
      this.setStatus(`Error al cargar: ${(err as Error).message}`, "error");
      this.renderInitial("<p><br></p>");
    }
  }

  // ── Renderizado inicial (una sola página) ─────────────────────

  /**
   * Vuelca todo el HTML en una única página provisional.
   * paginateContent() la dividirá después.
   */
  private renderInitial(html: string): void {
    this.workspace.innerHTML = "";
    this.pages = [];

    const page = this.createPageElement(html);
    this.pages.push(page);
    this.workspace.appendChild(page);

    this.paginator.setPages(this.pages);
    this.updatePageCount();
  }

  // ── Paginación por desbordamiento ─────────────────────────────

  /**
   * Recorre las páginas existentes y corta el contenido que desborda
   * en páginas nuevas. Soporta corte fila a fila en tablas.
   */
  private async paginateContent(): Promise<void> {
    let pageIndex = 0;

    // Procesamos página a página; el array puede crecer durante el bucle
    while (pageIndex < this.pages.length) {
      const page  = this.pages[pageIndex];
      const inner = page.querySelector(".hwe-page-inner") as HTMLElement;
      if (!inner) { pageIndex++; continue; }

      const pageLimit = page.offsetTop + page.clientHeight;

      if (inner.scrollHeight <= page.clientHeight + 1) {
        // Esta página no desborda
        pageIndex++;
        continue;
      }

      // Crear la página siguiente donde irá el contenido sobrante
      const nextPage  = this.createPageElement("");
      const nextInner = nextPage.querySelector(".hwe-page-inner") as HTMLElement;
      nextInner.innerHTML = "";

      // Insertar la nueva página en el DOM y en el array
      const divider = this.makePageDivider(pageIndex + 2);
      this.workspace.insertBefore(divider, page.nextSibling);
      this.workspace.insertBefore(nextPage, divider.nextSibling);
      this.pages.splice(pageIndex + 1, 0, nextPage);

      // Mover los nodos que desbordan a la página siguiente
      this.overflowNodeToNext(inner, nextInner, page);

      // Esperar un frame para que el navegador recalcule el layout
      await this.waitFrames(1);

      // Continuar con la misma página por si sigue desbordando
      // (no incrementamos pageIndex)
    }

    // Renumerar divisores y registrar páginas en el Paginator
    this.renumberDividers();
    this.paginator.setPages(this.pages);
    this.updatePageCount();
  }

  /**
   * Mueve desde `sourceInner` a `targetInner` todos los nodos de bloque
   * que superan el límite de altura de `page`.
   * Para tablas, corta fila a fila.
   */
  private overflowNodeToNext(
    sourceInner: HTMLElement,
    targetInner: HTMLElement,
    page: HTMLElement
  ): void {
    const pageBottom = page.offsetTop + page.clientHeight;
    const nodes = Array.from(sourceInner.childNodes);

    // Recorremos de atrás hacia adelante para encontrar el punto de corte
    // sin invalidar índices al mover nodos
    let cutIndex = -1;

    for (let i = 0; i < nodes.length; i++) {
      const node = nodes[i];
      if (node.nodeType !== Node.ELEMENT_NODE) continue;

      const el   = node as HTMLElement;
      const rect = el.getBoundingClientRect();
      const elBottom = rect.bottom + window.scrollY;

      if (elBottom > pageBottom) {
        // Este nodo desborda
        if (el.tagName === "TABLE") {
          // Corte quirúrgico fila a fila
          this.splitTable(el, targetInner, page);
        } else {
          // Nodo de bloque entero → mover a la siguiente página
          // Primero movemos este y todos los siguientes
          cutIndex = i;
          break;
        }
      }
    }

    // Mover nodos desde cutIndex en adelante (si hubo corte de bloque)
    if (cutIndex >= 0) {
      const toMove = Array.from(sourceInner.childNodes).slice(cutIndex);
      toMove.forEach(n => targetInner.appendChild(n));
    }
  }

  /**
   * Divide una tabla en el punto donde sus filas desbordan la página.
   * La parte superior queda en su sitio; la inferior se mueve a targetInner
   * como una nueva tabla con los mismos atributos y colgroup.
   */
  private splitTable(
    table: HTMLElement,
    targetInner: HTMLElement,
    page: HTMLElement
  ): void {
    const pageBottom = page.offsetTop + page.clientHeight;

    // Recopilar todas las filas (thead + tbody + tfoot)
    const allRows = Array.from(table.querySelectorAll("tr")) as HTMLElement[];

    let splitRowIndex = -1;
    for (let i = 0; i < allRows.length; i++) {
      const rect = allRows[i].getBoundingClientRect();
      const rowBottom = rect.bottom + window.scrollY;
      if (rowBottom > pageBottom) {
        splitRowIndex = i;
        break;
      }
    }

    // Si no encontramos punto de corte o la primera fila ya desborda,
    // mover la tabla entera
    if (splitRowIndex <= 0) {
      targetInner.insertBefore(table, targetInner.firstChild);
      return;
    }

    // Filas que van a la página siguiente
    const rowsForNext = allRows.slice(splitRowIndex);

    // Construir tabla nueva con los mismos atributos
    const newTable = document.createElement("table");
    // Copiar atributos de la tabla original
    Array.from(table.attributes).forEach(attr => {
      newTable.setAttribute(attr.name, attr.value);
    });

    // Copiar colgroup si existe (mantiene anchos de columna)
    const colgroup = table.querySelector("colgroup");
    if (colgroup) {
      newTable.appendChild(colgroup.cloneNode(true));
    }

    // Copiar thead como cabecera repetida (comportamiento Word)
    const thead = table.querySelector("thead");
    if (thead) {
      newTable.appendChild(thead.cloneNode(true));
    }

    // Crear tbody para las filas sobrantes
    const newTbody = document.createElement("tbody");
    rowsForNext.forEach(row => {
      // Mover la fila (no clonar) para preservar el contenido editable
      newTbody.appendChild(row);
    });
    newTable.appendChild(newTbody);

    // Insertar la nueva tabla al inicio de la página siguiente
    targetInner.insertBefore(newTable, targetInner.firstChild);

    // Si la tabla original quedó vacía (solo thead), eliminarla
    const remainingRows = table.querySelectorAll("tbody tr, tfoot tr");
    if (remainingRows.length === 0) {
      table.parentElement?.removeChild(table);
    }
  }

  // ── Helpers de layout ─────────────────────────────────────────

  private waitFrames(n: number): Promise<void> {
    return new Promise(resolve => {
      let count = 0;
      const tick = () => {
        if (++count >= n) resolve();
        else requestAnimationFrame(tick);
      };
      requestAnimationFrame(tick);
    });
  }

  private renumberDividers(): void {
    const dividers = this.workspace.querySelectorAll(".hwe-page-divider span");
    dividers.forEach((span, i) => {
      span.textContent = `Página ${i + 2}`;
    });
  }

  // ── Página individual ─────────────────────────────────────────

  private createPageElement(html?: string): HTMLElement {
    const page = document.createElement("div");
    page.className = "hwe-page";

    const inner = document.createElement("div");
    inner.className = "hwe-page-inner";
    inner.setAttribute("contenteditable", "true");
    inner.setAttribute("spellcheck", "false");
    inner.innerHTML = html ?? "<p><br></p>";

    inner.addEventListener("input", () => {
      this.isDirty = true;
      this.toolbar.updateActiveStates();
      this.paginator.rebalanceFromPage(page);
    });
    inner.addEventListener("keydown", (e: KeyboardEvent) => this.onPageKeyDown(e));
    inner.addEventListener("mouseup", () => this.toolbar.updateActiveStates());

    page.appendChild(inner);
    return page;
  }

  private onPagesChanged(pages: HTMLElement[]): void {
    this.pages = pages;
    pages.forEach(page => {
      if (!page.parentElement) this.workspace.appendChild(page);
    });
    this.updatePageCount();
  }

  private makePageDivider(pageNumber: number): HTMLElement {
    const div = document.createElement("div");
    div.className = "hwe-page-divider";
    div.setAttribute("contenteditable", "false");
    const span = document.createElement("span");
    span.textContent = `Página ${pageNumber}`;
    div.appendChild(span);
    return div;
  }

  // ── Teclado ───────────────────────────────────────────────────

  private onPageKeyDown(e: KeyboardEvent): void {
    if (e.ctrlKey && e.key === "s") {
      e.preventDefault();
      this.save();
    }
  }

  // ── Guardar ───────────────────────────────────────────────────

  private async save(): Promise<void> {
    const saveBtn = this.toolbar.getSaveButton();
    saveBtn.disabled = true;
    this.setStatus("Guardando...", "saving");

    try {
      const html = this.collectHtml();
      await saveHtmlToFileField(
        this.baseUrl, this.entityName, this.entityId, this.fieldName, html
      );
      this.isDirty = false;
      this.setStatus("✓ Guardado correctamente", "success");
      setTimeout(() => this.setStatus("", ""), 3000);
    } catch (err) {
      this.setStatus(`✗ Error al guardar: ${(err as Error).message}`, "error");
    } finally {
      saveBtn.disabled = false;
    }
  }

  private collectHtml(): string {
    return this.pages
      .map((page, index) => {
        const inner = page.querySelector(".hwe-page-inner") as HTMLElement;
        const content = inner ? inner.innerHTML : page.innerHTML;
        if (index === 0) return content;
        return `<div style="page-break-before:always">${content}</div>`;
      })
      .join("\n");
  }

  // ── splitHtmlByPageBreaks (intacta, para futuros usos) ────────

  private splitHtmlByPageBreaks(html: string): string[] {
    const temp = document.createElement("div");
    temp.innerHTML = html;
    const segments: string[] = [];
    let current = document.createElement("div");

    const processNode = (node: ChildNode): void => {
      if (node.nodeType !== Node.ELEMENT_NODE) {
        current.appendChild(node.cloneNode(true));
        return;
      }
      const el    = node as HTMLElement;
      const style = el.getAttribute("style") ?? "";
      const isBreakBefore =
        /page-break-before\s*:\s*always/i.test(style) ||
        (el.className && el.className.includes("page-break"));
      const isHrBreak =
        el.tagName === "HR" &&
        (/page-break/i.test(style) || el.className.includes("page-break"));

      if (isBreakBefore || isHrBreak) {
        segments.push(current.innerHTML || "<p><br></p>");
        current = document.createElement("div");
        if (!isHrBreak) {
          const clone = el.cloneNode(true) as HTMLElement;
          clone.style.cssText = clone.style.cssText
            .replace(/page-break-before\s*:\s*always\s*;?/gi, "")
            .replace(/page-break-after\s*:\s*always\s*;?/gi, "")
            .trim();
          current.appendChild(clone);
        }
        return;
      }
      const isBreakAfter = /page-break-after\s*:\s*always/i.test(style);
      if (isBreakAfter) {
        const clone = el.cloneNode(true) as HTMLElement;
        clone.style.cssText = clone.style.cssText
          .replace(/page-break-after\s*:\s*always\s*;?/gi, "")
          .trim();
        current.appendChild(clone);
        segments.push(current.innerHTML);
        current = document.createElement("div");
        return;
      }
      current.appendChild(el.cloneNode(true));
    };

    Array.from(temp.childNodes).forEach(processNode);
    const last = current.innerHTML.trim();
    segments.push(last || "<p><br></p>");
    return segments;
  }

  // ── Auxiliares ────────────────────────────────────────────────

  private updatePageCount(): void {
    this.pageCountEl.textContent = `Páginas: ${this.pages.length}`;
  }

  private setStatus(msg: string, type: "success" | "error" | "saving" | ""): void {
    this.statusMsg.textContent = msg;
    this.statusMsg.className = "hwe-status-msg" + (type ? ` ${type}` : "");
  }

  destroy(): void {
    this.paginator.destroy();
    this.container.innerHTML = "";
  }
}

interface IInputs {
  htmlContent: ComponentFramework.PropertyTypes.StringProperty;
  entityName:  ComponentFramework.PropertyTypes.StringProperty;
  fieldName:   ComponentFramework.PropertyTypes.StringProperty;
}
