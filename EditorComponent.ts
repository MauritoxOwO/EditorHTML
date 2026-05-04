import { Paginator } from "../Orquestador/Paginator";
import { Toolbar } from "../Resize/Toolbar";
import { fetchHtmlFromFileField, saveHtmlToFileField } from "../execCommand/fileApi";

type PcfContext = ComponentFramework.Context<IInputs>;
type StatusType = "success" | "error" | "saving" | "";

export class EditorComponent {
  private readonly container: HTMLElement;

  private root!: HTMLElement;
  private workspace!: HTMLElement;
  private statusMsg!: HTMLElement;
  private pageCountEl!: HTMLElement;

  private pages: HTMLElement[] = [];
  private paginator!: Paginator;
  private toolbar!: Toolbar;

  private readonly baseUrl: string;
  private readonly entityName: string;
  private readonly entityId: string;
  private readonly fieldName: string;

  private isDirty = false;

  constructor(container: HTMLElement, context: PcfContext) {
    this.container = container;

    const runtime = context as unknown as {
      page?: { getClientUrl?: () => string; entityId?: string };
      mode?: { contextInfo?: { entityId?: string; entityTypeName?: string } };
    };

    this.baseUrl = this.getClientUrl(runtime);
    this.entityId = this.cleanGuid(
      runtime.page?.entityId ?? runtime.mode?.contextInfo?.entityId ?? ""
    );
    this.entityName = runtime.mode?.contextInfo?.entityTypeName ?? "mcdev_htmldevtests";
    this.fieldName = "mcdev_htmlarchivooriginal";
  }

  async init(): Promise<void> {
    this.buildShell();
    this.paginator = new Paginator(
      (html?: string) => this.createPageElement(html),
      (pages: HTMLElement[]) => this.onPagesChanged(pages)
    );
    await this.loadContent();
  }

  private buildShell(): void {
    this.container.innerHTML = "";
    this.container.style.cssText =
      "width:100%;height:100%;overflow:hidden;display:flex;flex-direction:column;";

    this.root = document.createElement("div");
    this.root.className = "hwe-root";

    this.toolbar = new Toolbar();
    const toolbarEl = this.toolbar.build();
    this.toolbar.getSaveButton().addEventListener("click", () => void this.save());
    this.root.appendChild(toolbarEl);

    this.workspace = document.createElement("div");
    this.workspace.className = "hwe-workspace";
    this.root.appendChild(this.workspace);

    const statusBar = document.createElement("div");
    statusBar.className = "hwe-statusbar";

    this.pageCountEl = document.createElement("span");
    this.pageCountEl.textContent = "Paginas: 0";
    statusBar.appendChild(this.pageCountEl);

    this.statusMsg = document.createElement("span");
    this.statusMsg.className = "hwe-status-msg";
    statusBar.appendChild(this.statusMsg);

    this.root.appendChild(statusBar);
    this.container.appendChild(this.root);
  }

  private getClientUrl(runtime: {
    page?: { getClientUrl?: () => string };
  }): string {
    const pageClientUrl = runtime.page?.getClientUrl?.();
    if (pageClientUrl) return pageClientUrl;

    const globalContext = (window as unknown as {
      Xrm?: { Utility?: { getGlobalContext?: () => { getClientUrl?: () => string } } };
    }).Xrm?.Utility?.getGlobalContext?.();

    return globalContext?.getClientUrl?.() ?? "";
  }

  private cleanGuid(value: string): string {
    return value.replace(/[{}]/g, "");
  }

  private async loadContent(): Promise<void> {
    this.setStatus("Cargando contenido...", "saving");

    try {
      if (!this.baseUrl) {
        throw new Error("No se pudo obtener la URL de Dataverse.");
      }

      if (!this.entityId) {
        throw new Error("No se pudo obtener el Id del registro actual.");
      }

      const html = await fetchHtmlFromFileField(
        this.baseUrl,
        this.entityName,
        this.entityId,
        this.fieldName
      );

      this.renderInitial(html || "<p><br></p>");
      await this.waitFrames(2);
      await this.paginateContent();

      this.setStatus("", "");
    } catch (err) {
      this.setStatus(`Error al cargar: ${(err as Error).message}`, "error");
      this.renderInitial("<p><br></p>");
      this.paginator.setPages(this.pages);
    }
  }

  private renderInitial(html: string): void {
    this.workspace.innerHTML = "";
    this.pages = [];

    const page = this.createPageElement(this.normalizeHtmlForPagination(html));
    this.pages.push(page);
    this.workspace.appendChild(page);
    this.updatePageCount();
  }

  private async paginateContent(): Promise<void> {
    let pageIndex = 0;
    let safety = 0;

    while (pageIndex < this.pages.length && safety++ < 100) {
      const page = this.pages[pageIndex];
      const inner = this.getInner(page);

      if (!inner || !this.pageOverflows(page)) {
        pageIndex++;
        continue;
      }

      const nextPage = this.createPageElement("");
      const nextInner = this.getInner(nextPage);
      if (!nextInner) {
        pageIndex++;
        continue;
      }

      nextInner.innerHTML = "";
      const moved = this.extractOverflowToNextPage(inner, nextInner, page);

      if (!moved) {
        pageIndex++;
        continue;
      }

      const divider = this.makePageDivider(pageIndex + 2);
      this.workspace.insertBefore(divider, page.nextSibling);
      this.workspace.insertBefore(nextPage, divider.nextSibling);
      this.pages.splice(pageIndex + 1, 0, nextPage);

      await this.waitFrames(1);
    }

    this.renumberDividers();
    this.paginator.setPages(this.pages);
    this.updatePageCount();
  }

  private extractOverflowToNextPage(
    sourceInner: HTMLElement,
    targetInner: HTMLElement,
    page: HTMLElement
  ): boolean {
    let movedAny = false;
    let safety = 0;

    while (this.pageOverflows(page) && safety++ < 200) {
      const lastChild = this.getLastMeaningfulChild(sourceInner);

      if (!lastChild) {
        break;
      }

      if (
        lastChild.nodeType === Node.ELEMENT_NODE &&
        this.unwrapIfSplittableContainer(lastChild as HTMLElement)
      ) {
        continue;
      }

      if (
        lastChild.nodeType === Node.ELEMENT_NODE &&
        (lastChild as HTMLElement).tagName === "TABLE"
      ) {
        const split = this.splitTable(lastChild as HTMLElement, targetInner, page);
        if (!split) break;
        movedAny = true;
        continue;
      }

      targetInner.insertBefore(lastChild, targetInner.firstChild);
      movedAny = true;
    }

    return movedAny;
  }

  private normalizeHtmlForPagination(html: string): string {
    const temp = document.createElement("div");
    temp.innerHTML = html || "<p><br></p>";

    const body = temp.querySelector("body");
    if (body) temp.innerHTML = body.innerHTML;

    this.stripPageBreakStyles(temp);

    let safety = 0;
    while (safety++ < 20) {
      const children = this.getMeaningfulChildren(temp);
      if (children.length !== 1 || children[0].nodeType !== Node.ELEMENT_NODE) break;

      const onlyChild = children[0] as HTMLElement;
      if (!this.isSplittableContainer(onlyChild)) break;

      temp.innerHTML = onlyChild.innerHTML;
    }

    return temp.innerHTML.trim() || "<p><br></p>";
  }

  private stripPageBreakStyles(root: HTMLElement): void {
    root.querySelectorAll<HTMLElement>("[style]").forEach((element) => {
      element.style.cssText = element.style.cssText
        .replace(/page-break-before\s*:\s*always\s*;?/gi, "")
        .replace(/page-break-after\s*:\s*always\s*;?/gi, "")
        .replace(/break-before\s*:\s*page\s*;?/gi, "")
        .replace(/break-after\s*:\s*page\s*;?/gi, "")
        .trim();
    });
  }

  private splitTable(
    table: HTMLElement,
    targetInner: HTMLElement,
    page: HTMLElement
  ): boolean {
    const pageBottom = page.getBoundingClientRect().bottom;
    const rows = Array.from(table.querySelectorAll("tr")) as HTMLElement[];

    let splitRowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].getBoundingClientRect().bottom > pageBottom) {
        splitRowIndex = i;
        break;
      }
    }

    if (splitRowIndex <= 0) {
      targetInner.insertBefore(table, targetInner.firstChild);
      return true;
    }

    const rowsForNext = rows.slice(splitRowIndex);
    if (rowsForNext.length === 0) return false;

    const newTable = document.createElement("table");
    Array.from(table.attributes).forEach((attr) => {
      newTable.setAttribute(attr.name, attr.value);
    });

    const colgroup = table.querySelector("colgroup");
    if (colgroup) newTable.appendChild(colgroup.cloneNode(true));

    const thead = table.querySelector("thead");
    if (thead) newTable.appendChild(thead.cloneNode(true));

    const tbody = document.createElement("tbody");
    rowsForNext.forEach((row) => tbody.appendChild(row));
    newTable.appendChild(tbody);

    targetInner.insertBefore(newTable, targetInner.firstChild);

    const remainingRows = table.querySelectorAll("tbody tr, tfoot tr");
    if (remainingRows.length === 0) table.remove();

    return true;
  }

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

    pages.forEach((page, index) => {
      if (!page.parentElement) {
        const previousPage = pages[index - 1];
        if (previousPage?.parentElement) {
          const divider = this.makePageDivider(index + 1);
          this.workspace.insertBefore(divider, previousPage.nextSibling);
          this.workspace.insertBefore(page, divider.nextSibling);
        } else {
          this.workspace.appendChild(page);
        }
      }
    });

    this.removeEmptyDividers();
    this.renumberDividers();
    this.updatePageCount();
  }

  private makePageDivider(pageNumber: number): HTMLElement {
    const divider = document.createElement("div");
    divider.className = "hwe-page-divider";
    divider.setAttribute("contenteditable", "false");

    const label = document.createElement("span");
    label.textContent = `Pagina ${pageNumber}`;
    divider.appendChild(label);

    return divider;
  }

  private removeEmptyDividers(): void {
    const dividers = Array.from(this.workspace.querySelectorAll(".hwe-page-divider"));
    dividers.forEach((divider) => {
      const next = divider.nextElementSibling;
      if (!next || !next.classList.contains("hwe-page")) divider.remove();
    });
  }

  private renumberDividers(): void {
    const dividers = this.workspace.querySelectorAll(".hwe-page-divider span");
    dividers.forEach((span, index) => {
      span.textContent = `Pagina ${index + 2}`;
    });
  }

  private pageOverflows(page: HTMLElement): boolean {
    const inner = this.getInner(page);
    if (!inner) return false;
    return inner.scrollHeight > page.clientHeight + 1;
  }

  private getLastMeaningfulChild(container: HTMLElement): ChildNode | null {
    const children = this.getMeaningfulChildren(container);
    if (children.length > 1) return children[children.length - 1];

    const onlyChild = children[0];
    if (
      onlyChild?.nodeType === Node.ELEMENT_NODE &&
      this.isSplittableContainer(onlyChild as HTMLElement)
    ) {
      return onlyChild;
    }

    return null;
  }

  private getMeaningfulChildren(container: HTMLElement): ChildNode[] {
    return Array.from(container.childNodes).filter((node) => !this.isEmptyNode(node));
  }

  private unwrapIfSplittableContainer(element: HTMLElement): boolean {
    if (!this.isSplittableContainer(element) || !element.parentNode) return false;

    const parent = element.parentNode;
    while (element.firstChild) {
      parent.insertBefore(element.firstChild, element);
    }
    parent.removeChild(element);
    return true;
  }

  private isSplittableContainer(element: HTMLElement): boolean {
    const splittableTags = new Set(["DIV", "SECTION", "ARTICLE", "MAIN", "BODY"]);
    if (!splittableTags.has(element.tagName)) return false;
    if (element.classList.contains("hwe-page") || element.classList.contains("hwe-page-inner")) {
      return false;
    }

    return this.getMeaningfulChildren(element).length > 0;
  }

  private isEmptyNode(node: ChildNode): boolean {
    if (node.nodeType === Node.TEXT_NODE) {
      return (node.textContent ?? "").trim() === "";
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return false;

    const el = node as HTMLElement;
    if (el.tagName === "BR") return true;

    return (
      el.tagName === "P" &&
      el.childNodes.length <= 1 &&
      (el.textContent ?? "").trim() === ""
    );
  }

  private getInner(page: HTMLElement): HTMLElement | null {
    return page.querySelector(".hwe-page-inner");
  }

  private waitFrames(n: number): Promise<void> {
    return new Promise((resolve) => {
      let count = 0;
      const tick = () => {
        count++;
        if (count >= n) {
          resolve();
          return;
        }
        requestAnimationFrame(tick);
      };
      requestAnimationFrame(tick);
    });
  }

  private onPageKeyDown(e: KeyboardEvent): void {
    if (e.ctrlKey && e.key.toLowerCase() === "s") {
      e.preventDefault();
      void this.save();
    }
  }

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
      this.setStatus("Guardado correctamente", "success");
      window.setTimeout(() => this.setStatus("", ""), 3000);
    } catch (err) {
      this.setStatus(`Error al guardar: ${(err as Error).message}`, "error");
    } finally {
      saveBtn.disabled = false;
    }
  }

  private collectHtml(): string {
    return this.pages
      .map((page, index) => {
        const inner = this.getInner(page);
        const content = inner ? inner.innerHTML : page.innerHTML;
        if (index === 0) return content;
        return `<div style="page-break-before:always">${content}</div>`;
      })
      .join("\n");
  }

  private updatePageCount(): void {
    this.pageCountEl.textContent = `Paginas: ${this.pages.length}`;
  }

  private setStatus(message: string, type: StatusType): void {
    this.statusMsg.textContent = message;
    this.statusMsg.className = "hwe-status-msg" + (type ? ` ${type}` : "");
  }

  destroy(): void {
    this.paginator?.destroy();
    this.container.innerHTML = "";
  }
}

interface IInputs {
  htmlContent: ComponentFramework.PropertyTypes.StringProperty;
  entityName: ComponentFramework.PropertyTypes.StringProperty;
  fieldName: ComponentFramework.PropertyTypes.StringProperty;
}
