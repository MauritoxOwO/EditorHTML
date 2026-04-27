import { Paginator } from "../Orquestador/Paginator";
import { Toolbar } from "../Resize/Toolbar";
import { fetchHtmlFromFileField, saveHtmlToFileField } from "../execCommand/fileApi";
type PcfContext = ComponentFramework.Context<IInputs>;

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
    context: PcfContext
  ) {
    this.container = container;
    this.entityName = "mcdev_htmldevtests";
    this.fieldName = "mcdev_htmlarchivooriginal";
    this.entityId = (context as unknown as { page: { entityId: string } }).page.entityId ?? "";
  }

  async init(): Promise<void> {
    this.buildShell();
    this.paginator = new Paginator(
      (html?: string) => this.createPageElement(html),
      (pages: HTMLElement[]) => this.onPagesChanged(pages)
    );
    await this.loadContent();
  }

  // DOM shell

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
      this.renderFromHtml("<p><br></p>");
    }
  }

  // Renderizado inicial desde HTML 

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

  this.paginator.setPages(this.pages);

  // ← Esperar dos frames: el primero pinta, el segundo calcula layout
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      this.pages.forEach((page) => this.paginator.rebalanceFromPage(page));
      this.updatePageCount();
    });
  });
}

  // ── Página individual ─────────────────────────────────────────


  private createPageElement(html?: string): HTMLElement {
    const page = document.createElement("div");
    page.className = "hwe-page";

    page.style.height = "297mm";
    page.style.overflow = "hidden";

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

    pages.forEach((page) => {
      if (!page.parentElement) {
        this.workspace.appendChild(page);
      }
    })
    this.updatePageCount();
  }

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

  // ── Teclado ───────────────────────────────────────────────────

  private onPageKeyDown(e: KeyboardEvent): void {
    // Ctrl+S → Guardar
    if (e.ctrlKey && e.key === "s") {
      e.preventDefault();
      this.save();
      return;
    }
  }

  // Guardar

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

  // Split por page-break

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

  // Funciones auxiliares 
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
