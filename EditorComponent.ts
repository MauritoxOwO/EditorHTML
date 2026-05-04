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

  private rebalanceTimer: number | undefined;
  private isComposing = false;
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
      if (!this.baseUrl) throw new Error("No se pudo obtener la URL de Dataverse.");
      if (!this.entityId) throw new Error("No se pudo obtener el Id del registro actual.");

      const html = await fetchHtmlFromFileField(
        this.baseUrl,
        this.entityName,
        this.entityId,
        this.fieldName
      );

      await this.renderAndPaginate(html || "<p><br></p>");
      this.setStatus("", "");
    } catch (err) {
      this.setStatus(`Error al cargar: ${(err as Error).message}`, "error");
      await this.renderAndPaginate("<p><br></p>");
    }
  }

  private async renderAndPaginate(html: string): Promise<void> {
    this.workspace.innerHTML = "";
    this.pages = [this.createPageElement(this.normalizeHtmlForPagination(html))];
    this.workspace.appendChild(this.pages[0]);

    this.paginator.setPages(this.pages);
    await this.waitForImages(this.workspace);
    await this.waitFrames(2);

    this.paginator.repaginateAll();
    this.pages = this.paginator.getPages();
    this.syncWorkspace();
    this.updatePageCount();
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
      if (!this.isComposing) this.scheduleRebalance(page);
    });
    inner.addEventListener("compositionstart", () => {
      this.isComposing = true;
    });
    inner.addEventListener("compositionend", () => {
      this.isComposing = false;
      this.scheduleRebalance(page);
    });
    inner.addEventListener("keydown", (event: KeyboardEvent) => this.onPageKeyDown(event));
    inner.addEventListener("mouseup", () => this.toolbar.updateActiveStates());

    page.appendChild(inner);
    return page;
  }

  private scheduleRebalance(page: HTMLElement): void {
    if (this.rebalanceTimer !== undefined) {
      window.clearTimeout(this.rebalanceTimer);
    }

    this.rebalanceTimer = window.setTimeout(() => {
      this.rebalanceTimer = undefined;
      if (!this.pages.includes(page)) return;

      this.paginator.rebalanceFromPage(page);
      this.pages = this.paginator.getPages();
      this.syncWorkspace();
      this.updatePageCount();
    }, 180);
  }

  private onPagesChanged(pages: HTMLElement[]): void {
    this.pages = pages;
    this.syncWorkspace();
    this.updatePageCount();
  }

  private syncWorkspace(): void {
    const pageSet = new Set(this.pages);

    Array.from(this.workspace.querySelectorAll(".hwe-page")).forEach((page) => {
      if (!pageSet.has(page as HTMLElement)) page.remove();
    });

    this.workspace.querySelectorAll(".hwe-page-divider").forEach((divider) => {
      divider.remove();
    });

    let previousPage: HTMLElement | null = null;

    this.pages.forEach((page, index) => {
      if (index === 0) {
        if (this.workspace.firstElementChild !== page) {
          this.workspace.insertBefore(page, this.workspace.firstChild);
        }
      } else {
        if (previousPage && page.previousElementSibling !== previousPage) {
          this.workspace.insertBefore(page, previousPage.nextSibling);
        } else if (!page.parentElement) {
          this.workspace.appendChild(page);
        }

        this.workspace.insertBefore(this.makePageDivider(index + 1), page);
      }

      previousPage = page;
    });
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

  private normalizeHtmlForPagination(html: string): string {
    const temp = document.createElement("div");
    temp.innerHTML = html || "<p><br></p>";

    const body = temp.querySelector("body");
    if (body) temp.innerHTML = body.innerHTML;

    this.cleanImportedHtml(temp);

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

  private cleanImportedHtml(root: HTMLElement): void {
    this.removeComments(root);
    root.querySelectorAll("style, meta, link, xml, script, object, hr").forEach((node) => {
      node.remove();
    });
    this.removeOfficeNamespacedNodes(root);

    root.querySelectorAll<HTMLElement>("[style]").forEach((element) => {
      element.style.cssText = this.sanitizeInlineStyle(element.style.cssText);
      if (!element.getAttribute("style")) element.removeAttribute("style");
    });

    root.querySelectorAll<HTMLElement>("[width]").forEach((element) => {
      if (element.tagName !== "IMG") element.removeAttribute("width");
    });
    root.querySelectorAll<HTMLElement>("[height]").forEach((element) => {
      if (element.tagName !== "IMG") element.removeAttribute("height");
    });

    this.unwrapKnownWordContainers(root);
    this.removeVisuallyEmptyNodes(root);
    this.removeBorderOnlyBlocks(root);
  }

  private sanitizeInlineStyle(style: string): string {
    return style
      .replace(/mso-[^:;]+:[^;]+;?/gi, "")
      .replace(/page-break-before\s*:\s*always\s*;?/gi, "")
      .replace(/page-break-after\s*:\s*always\s*;?/gi, "")
      .replace(/break-before\s*:\s*page\s*;?/gi, "")
      .replace(/break-after\s*:\s*page\s*;?/gi, "")
      .replace(/margin-left\s*:[^;]+;?/gi, "")
      .replace(/margin-right\s*:[^;]+;?/gi, "")
      .replace(/text-indent\s*:[^;]+;?/gi, "")
      .replace(/position\s*:[^;]+;?/gi, "")
      .replace(/left\s*:[^;]+;?/gi, "")
      .replace(/right\s*:[^;]+;?/gi, "")
      .replace(/transform\s*:[^;]+;?/gi, "")
      .replace(/min-width\s*:[^;]+;?/gi, "")
      .replace(/max-width\s*:[^;]+;?/gi, "")
      .replace(/width\s*:[^;]+;?/gi, "")
      .replace(/white-space\s*:[^;]+;?/gi, "")
      .replace(/tab-stops\s*:[^;]+;?/gi, "")
      .trim();
  }

  private removeComments(root: HTMLElement): void {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_COMMENT);
    const comments: Node[] = [];

    while (walker.nextNode()) {
      comments.push(walker.currentNode);
    }

    comments.forEach((comment) => comment.parentNode?.removeChild(comment));
  }

  private removeOfficeNamespacedNodes(root: HTMLElement): void {
    Array.from(root.querySelectorAll("*")).forEach((node) => {
      const tagName = node.tagName.toLowerCase();
      if (tagName.includes(":") && /^(o|v|w|m):/i.test(tagName)) node.remove();
    });
  }

  private unwrapKnownWordContainers(root: HTMLElement): void {
    root.querySelectorAll<HTMLElement>("div.WordSection1, div[class*='WordSection']").forEach(
      (element) => this.unwrapElement(element)
    );
  }

  private removeVisuallyEmptyNodes(root: HTMLElement): void {
    Array.from(root.childNodes).forEach((node) => {
      if (node.nodeType === Node.ELEMENT_NODE) {
        this.removeVisuallyEmptyNodes(node as HTMLElement);
      }

      if (this.isEmptyNode(node)) {
        node.parentNode?.removeChild(node);
      }
    });
  }

  private removeBorderOnlyBlocks(root: HTMLElement): void {
    Array.from(root.querySelectorAll<HTMLElement>("p, div, section, article, span")).forEach(
      (element) => {
        const text = (element.textContent ?? "").replace(/\u00a0/g, " ").trim();
        const hasMedia = !!element.querySelector("img, table, tr, td, th, video, canvas, svg");
        const hasBorder = /border/i.test(element.getAttribute("style") ?? "");

        if (!text && !hasMedia && hasBorder) element.remove();
      }
    );
  }

  private unwrapElement(element: HTMLElement): void {
    if (!element.parentNode) return;

    const parent = element.parentNode;
    while (element.firstChild) {
      parent.insertBefore(element.firstChild, element);
    }
    parent.removeChild(element);
  }

  private isSplittableContainer(element: HTMLElement): boolean {
    const splittableTags = new Set(["DIV", "SECTION", "ARTICLE", "MAIN", "BODY"]);
    if (!splittableTags.has(element.tagName)) return false;
    if (element.classList.contains("hwe-page") || element.classList.contains("hwe-page-inner")) {
      return false;
    }

    return this.getMeaningfulChildren(element).length > 0;
  }

  private getMeaningfulChildren(container: HTMLElement): ChildNode[] {
    return Array.from(container.childNodes).filter((node) => !this.isEmptyNode(node));
  }

  private isEmptyNode(node: ChildNode): boolean {
    if (node.nodeType === Node.COMMENT_NODE) return true;

    if (node.nodeType === Node.TEXT_NODE) {
      return (node.textContent ?? "").replace(/\u00a0/g, " ").trim() === "";
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return false;

    const element = node as HTMLElement;
    if (element.tagName === "BR") return true;
    if (["META", "LINK", "STYLE", "SCRIPT", "XML"].includes(element.tagName)) return true;
    if (element.querySelector("img, table, tr, td, th, video, canvas, svg")) return false;

    return (
      ["P", "DIV", "SECTION", "ARTICLE", "SPAN"].includes(element.tagName) &&
      (element.textContent ?? "").replace(/\u00a0/g, " ").trim() === "" &&
      Array.from(element.childNodes).every((child) => this.isEmptyNode(child))
    );
  }

  private waitForImages(root: HTMLElement): Promise<void> {
    const pending = Array.from(root.querySelectorAll("img")).filter((img) => !img.complete);
    if (pending.length === 0) return Promise.resolve();

    return Promise.all(
      pending.map(
        (img) =>
          new Promise<void>((resolve) => {
            img.addEventListener("load", () => resolve(), { once: true });
            img.addEventListener("error", () => resolve(), { once: true });
          })
      )
    ).then(() => undefined);
  }

  private waitFrames(count: number): Promise<void> {
    return new Promise((resolve) => {
      let frame = 0;
      const tick = () => {
        frame++;
        if (frame >= count) {
          resolve();
          return;
        }
        requestAnimationFrame(tick);
      };
      requestAnimationFrame(tick);
    });
  }

  private onPageKeyDown(event: KeyboardEvent): void {
    if (event.ctrlKey && event.key.toLowerCase() === "s") {
      event.preventDefault();
      void this.save();
    }
  }

  private async save(): Promise<void> {
    const saveButton = this.toolbar.getSaveButton();
    saveButton.disabled = true;
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
      saveButton.disabled = false;
    }
  }

  private collectHtml(): string {
    return this.pages
      .map((page, index) => {
        const inner = page.querySelector(".hwe-page-inner") as HTMLElement | null;
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
    if (this.rebalanceTimer !== undefined) {
      window.clearTimeout(this.rebalanceTimer);
    }
    this.paginator?.destroy();
    this.container.innerHTML = "";
  }
}

interface IInputs {
  htmlContent: ComponentFramework.PropertyTypes.StringProperty;
  entityName: ComponentFramework.PropertyTypes.StringProperty;
  fieldName: ComponentFramework.PropertyTypes.StringProperty;
}
