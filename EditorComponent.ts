/**
 * EditorComponent.ts
 *
 * Orquestador principal del editor:
 *  1. Construye el DOM (toolbar + workspace + statusbar)
 *  2. Carga el HTML desde Dataverse
 *  3. Divide el HTML en páginas iniciales (por marcadores page-break)
 *  4. Instancia el Paginator para paginación automática por overflow
 *  5. Gestiona el guardado
 */

import { Paginator }                              from "./Paginator";
import { Toolbar }                                from "./Toolbar";
import { fetchHtmlFromFileField, saveHtmlToFileField } from "./fileApi";

export class EditorComponent {
  private container: HTMLElement;

  // DOM principal
  private root!: HTMLElement;
  private workspace!: HTMLElement;
  private statusMsg!: HTMLElement;
  private pageCountEl!: HTMLElement;

  // Páginas actuales en el DOM
  private pages: HTMLElement[] = [];

  // Módulos
  private paginator!: Paginator;
  private toolbar!: Toolbar;

  // Dataverse
  private baseUrl: string;
  private entityName: string;
  private entityId: string;
  private fieldName: string;

  private isDirty = false;

  constructor(
    container: HTMLElement,
    context: ComponentFramework.Context<any>
  ) {
    this.container   = container;
    this.baseUrl     = (window as any).Xrm.Utility.getGlobalContext().getClientUrl();
    this.entityName  = context.parameters.entityName.raw  ?? "";
    this.fieldName   = context.parameters.fieldName.raw   ?? "";
    this.entityId    = (context as any).page.entityId     ?? "";
  }

  async init(): Promise<void> {
    this.buildShell();
    this.paginator = new Paginator(
      (html) => this.createPageElement(html),
      (pages) => this.onPagesChanged(pages)
    );
    await this.loadContent();
  }

  // ── DOM shell ─────────────────────────────────────────────────

  private buildShell(): void {
    this.container.innerHTML = "";
    this.container.style.cssText = "width:100%;height:100%;overflow:hidden;display:flex;flex-direction:column;";

    // Raíz
    this.root = document.createElement("div");
    this.root.className = "hwe-root";

    // Toolbar
    this.toolbar = new Toolbar();
    const toolbarEl = this.toolbar.build();
    this.toolbar.getSaveButton().addEventListener("click", () => this.save());
    this.root.appendChild(toolbarEl);

    // Workspace (zona scrollable con fondo gris)
    this.workspace = document.createElement("div");
    this.workspace.className = "hwe-workspace";
    this.root.appendChild(this.workspace);

    // Status bar inferior
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

  // Carga

  private async loadContent(): Promise<void> {
    this.setStatus("Cargando contenido...", "saving");
    try {
      const html = await fetchHtmlFromFileField(
        this.baseUrl,
        this.entityName,
        this.entityId,
        this.fieldName
      );
      this.renderFromHtml(html);
      this.setStatus("", "");
    } catch (err) {
      this.setStatus(`Error al cargar: ${(err as Error).message}`, "error");
      // Mostrar al menos una página vacía editable
      this.renderFromHtml("<p><br></p>");
    }
  }

  //Renderizado inicial desde HTML 

  /**
   * Divide el HTML en segmentos por marcadores page-break,
   * crea una página por segmento y registra todas en el Paginator.
   */
  private renderFromHtml(html: string): void {
    this.workspace.innerHTML = "";
    this.pages = [];

    const segments = this.splitHtmlByPageBreaks(html);

    segments.forEach((segHtml, index) => {
      if (index > 0) {
        this.workspace.appendChild(this.makePageDivider(index + 1));
      }
      const page = this.createPageElement(segHtml);
      this.pages.push(page);
      this.workspace.appendChild(page);
    });

    if (this.pages.length === 0) {
      const page = this.createPageElement("<p><br></p>");
      this.pages.push(page);
      this.workspace.appendChild(page);
    }

    // Registrar en el Paginator
    this.paginator.setPages(this.pages);
    this.updatePageCount();
  }

  // Página individual 

private createPageElement(html?: string): HTMLElement {
    const page = document.createElement("div");
    page.className = "hwe-page";

    // El inner es el contenedor real del contenido y el que es editable
    const inner = document.createElement("div");
    inner.className = "hwe-page-inner";
    inner.setAttribute("contenteditable", "true");
    inner.setAttribute("spellcheck", "false");
    inner.innerHTML = html ?? "<p><br></p>";

    inner.addEventListener("input", () => {
        this.isDirty = true;
        this.toolbar.updateActiveStates();
    });
    inner.addEventListener("keydown", (e: KeyboardEvent) => this.onPageKeyDown(e));
    inner.addEventListener("mouseup", () => this.toolbar.updateActiveStates());

    page.appendChild(inner);
    return page;
}

  /**
   * Callback del Paginator: se llama cada vez que la lista de páginas
   * cambia (se crea o elimina una página).
   * Reconstruimos los separadores visuales y actualizamos el contador.
   */
  private onPagesChanged(pages: HTMLElement[]): void {
    this.pages = pages;
    this.rebuildWorkspace();
    this.updatePageCount();
  }

  /**
   * Reconstruye el workspace colocando las páginas con sus divisores
   * entre ellas. No recrea las páginas, solo reorganiza el DOM.
   */
  private rebuildWorkspace(): void {
    this.workspace.innerHTML = "";

    this.pages.forEach((page, index) => {
      if (index > 0) {
        this.workspace.appendChild(this.makePageDivider(index + 1));
      }
      this.workspace.appendChild(page);
    });
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

  // Teclado 

  private onPageKeyDown(e: KeyboardEvent): void {
    // Ctrl+S → Guardar
    if (e.ctrlKey && e.key === "s") {
      e.preventDefault();
      this.save();
      return;
    }

    // Enter al final de una página: dejar que el Paginator lo maneje
    // No necesitamos lógica especial; el ResizeObserver detectará el desborde.
  }

  //Guardar 

  private async save(): Promise<void> {
    const saveBtn = this.toolbar.getSaveButton();
    saveBtn.disabled = true;
    this.setStatus("Guardando...", "saving");

    try {
      const html = this.collectHtml();
      await saveHtmlToFileField(
        this.baseUrl,
        this.entityName,
        this.entityId,
        this.fieldName,
        html
      );
      this.isDirty = false;
      this.setStatus("✓ Guardado correctamente", "success");
      setTimeout(() => this.setStatus("", ""), 3000);
    } catch (err) {
      this.setStatus(
        `✗ Error al guardar: ${(err as Error).message}`,
        "error"
      );
    } finally {
      saveBtn.disabled = false;
    }
  }

  /**
   * Reconstruye el HTML completo desde todas las páginas,
   * reinsertando marcadores page-break-before entre ellas
   * para que al recargar se vuelvan a separar correctamente.
   */
  private collectHtml(): string {
    return this.pages
      .map((page, index) => {
        const inner = page.innerHTML;
        if (index === 0) {
          return inner;
        }
        // Envolver en div con page-break para persistir la separación
        return `<div style="page-break-before:always">${inner}</div>`;
      })
      .join("\n");
  }

  // Split por page-break 

  /**
   * Divide el HTML en segmentos según atributos page-break-before/after.
   * Devuelve al menos un segmento.
   */
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
        // Cerrar segmento actual
        segments.push(current.innerHTML || "<p><br></p>");
        current = document.createElement("div");

        if (!isHrBreak) {
          // Clonar el elemento sin el atributo page-break
          const clone = el.cloneNode(true) as HTMLElement;
          clone.style.cssText = clone.style.cssText
            .replace(/page-break-before\s*:\s*always\s*;?/gi, "")
            .replace(/page-break-after\s*:\s*always\s*;?/gi, "")
            .trim();
          current.appendChild(clone);
        }
        return;
      }

      // page-break-after
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

  // Helpers 

  private updatePageCount(): void {
    this.pageCountEl.textContent = `Páginas: ${this.pages.length}`;
  }

  private setStatus(
    msg: string,
    type: "success" | "error" | "saving" | ""
  ): void {
    this.statusMsg.textContent = msg;
    this.statusMsg.className =
      "hwe-status-msg" + (type ? ` ${type}` : "");
  }

  // Cleanup

  destroy(): void {
    this.paginator.destroy();
    this.container.innerHTML = "";
  }
}
