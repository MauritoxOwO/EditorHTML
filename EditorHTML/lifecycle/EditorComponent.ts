import { Paginator } from "../Orquestador/Paginator";
import { CaretManager } from "../Orquestador/CaretManager";
import { Toolbar } from "../Resize/Toolbar";
import { fetchHtmlFromFileField, saveHtmlToFileField } from "../execCommand/fileApi";
import {
  applyPageSetup,
  DEFAULT_PAGE_SETUP,
  makePageSetupWrapper,
  normalizePageSetup,
  PageSetup,
  readPageSetupFromElement,
} from "../Orquestador/PageGeometry";
import { WordPasteImporter } from "../import/WordPasteImporter";

type PcfContext = ComponentFramework.Context<IInputs>;
type StatusType = "success" | "error" | "saving" | "";

export interface EditorComponentOptions {
  initialHtml?: string;
  loadHtml?: () => Promise<string> | string;
  saveHtml?: (html: string) => Promise<void> | void;
}

export class EditorComponent {
  private readonly container: HTMLElement;
  private readonly options: EditorComponentOptions;

  private root!: HTMLElement;
  private workspace!: HTMLElement;
  private statusMsg!: HTMLElement;
  private pageCountEl!: HTMLElement;

  private pages: HTMLElement[] = [];
  private paginator!: Paginator;
  private toolbar!: Toolbar;
  private readonly wordPasteImporter = new WordPasteImporter();
  private pageSetup: PageSetup = DEFAULT_PAGE_SETUP;

  private readonly baseUrl: string;
  private readonly entityName: string;
  private readonly entityId: string;
  private readonly fieldName: string;

  private rebalanceFrame: number | undefined;
  private pendingRebalance: { page: HTMLElement; pullFromNextPages: boolean } | null = null;
  private readonly pagesNeedingPull = new WeakSet<HTMLElement>();
  private readonly pendingInputTypes = new WeakMap<HTMLElement, string>();
  private isComposing = false;
  private isDirty = false;

  constructor(container: HTMLElement, context?: PcfContext, options: EditorComponentOptions = {}) {
    this.container = container;
    this.options = options;

    const runtime = (context ?? {}) as unknown as {
      page?: { getClientUrl?: () => string; entityId?: string };
      mode?: { contextInfo?: { entityId?: string; entityTypeName?: string } };
    };

    this.baseUrl = this.getClientUrl(runtime);
    this.entityId = this.cleanGuid(
      runtime.page?.entityId ?? runtime.mode?.contextInfo?.entityId ?? ""
    );
    this.entityName = "mcdev_htmldevtests";
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
    this.applyCurrentPageSetup();

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

  private setPageSetup(setup: Partial<PageSetup> | PageSetup): void {
    this.pageSetup = normalizePageSetup(setup);
    this.applyCurrentPageSetup();
  }

  private applyCurrentPageSetup(): void {
    if (this.root) applyPageSetup(this.root, this.pageSetup);
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
      if (this.options.loadHtml || this.options.initialHtml !== undefined) {
        const html = this.options.loadHtml
          ? await this.options.loadHtml()
          : this.options.initialHtml ?? "<p><br></p>";

        await this.renderAndPaginate(html || "<p><br></p>");
        this.setStatus("", "");
        return;
      }

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
    this.setPageSetup(DEFAULT_PAGE_SETUP);

    const normalizedHtml = this.normalizeHtmlForPagination(html);
    this.pages = [this.createPageElement(normalizedHtml)];
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

    inner.addEventListener("beforeinput", (event: InputEvent) => {
      this.pendingInputTypes.set(page, event.inputType);
      if (this.isDeleteInput(event.inputType)) this.pagesNeedingPull.add(page);
    });
    inner.addEventListener("input", () => {
      this.isDirty = true;
      this.toolbar.updateActiveStates();
      if (!this.isComposing) {
        const inputType = this.pendingInputTypes.get(page) ?? "";
        this.pendingInputTypes.delete(page);
        this.syncEditableBlankBlocks(inner, this.isEnterInput(inputType));
        const shouldPullFromNextPages =
          this.isDeleteInput(inputType) || this.pagesNeedingPull.has(page);
        this.pagesNeedingPull.delete(page);
        this.scheduleRebalance(page, shouldPullFromNextPages);
      }
    });
    inner.addEventListener("compositionstart", () => {
      this.isComposing = true;
    });
    inner.addEventListener("compositionend", () => {
      this.isComposing = false;
      this.syncEditableBlankBlocks(inner, false);
      this.scheduleRebalance(page);
    });
    inner.addEventListener("paste", (event: ClipboardEvent) => this.onPaste(event, page));
    inner.addEventListener("keydown", (event: KeyboardEvent) => this.onPageKeyDown(event));
    inner.addEventListener("mouseup", () => this.toolbar.updateActiveStates());

    page.appendChild(inner);
    return page;
  }

  private scheduleRebalance(page: HTMLElement, pullFromNextPages = false): void {
    this.queueRebalance(page, pullFromNextPages);
    if (this.rebalanceFrame !== undefined) return;

    this.rebalanceFrame = window.requestAnimationFrame(() => {
      this.rebalanceFrame = undefined;

      const pending = this.pendingRebalance;
      this.pendingRebalance = null;
      if (!pending || !this.pages.includes(pending.page)) return;

      const pageIndex = this.pages.indexOf(pending.page);
      if (!this.shouldRebalancePage(pending.page, pageIndex, pending.pullFromNextPages)) return;

      const startPage = this.pages[Math.max(0, pageIndex - 1)] ?? pending.page;
      const activeEditable = this.getActiveEditable();
      const marker = CaretManager.createMarker(this.root);
      const caretViewportTop = marker?.getBoundingClientRect().top ?? null;
      this.paginator.rebalanceFromPage(startPage);
      this.pages = this.paginator.getPages();
      this.pages.forEach((currentPage) => {
        const inner = currentPage.querySelector<HTMLElement>(".hwe-page-inner");
        if (inner) this.syncEditableBlankBlocks(inner, false);
      });
      this.syncWorkspace();
      this.updatePageCount();
      const fallbackEditable = this.getEditableForPageIndex(pageIndex) ?? activeEditable;
      this.restoreCaretViewport(marker, caretViewportTop);
      CaretManager.restoreMarker(marker, fallbackEditable);
      CaretManager.removeMarkers(this.root);
    });
  }

  private queueRebalance(page: HTMLElement, pullFromNextPages: boolean): void {
    if (!this.pendingRebalance) {
      this.pendingRebalance = { page, pullFromNextPages };
      return;
    }

    const currentIndex = this.pages.indexOf(this.pendingRebalance.page);
    const nextIndex = this.pages.indexOf(page);
    const shouldUseNextPage =
      currentIndex === -1 || (nextIndex !== -1 && nextIndex < currentIndex);

    this.pendingRebalance = {
      page: shouldUseNextPage ? page : this.pendingRebalance.page,
      pullFromNextPages: this.pendingRebalance.pullFromNextPages || pullFromNextPages,
    };
  }

  private onPagesChanged(pages: HTMLElement[]): void {
    this.pages = pages;
    this.syncWorkspace();
    this.updatePageCount();
  }

  private syncWorkspace(): void {
    const shouldRestoreScroll = this.shouldPreserveWorkspaceScroll();
    const scrollTop = this.workspace.scrollTop;
    const scrollLeft = this.workspace.scrollLeft;
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

    if (shouldRestoreScroll) {
      this.restoreWorkspaceScroll(scrollTop, scrollLeft);
    }
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
    if (this.wordPasteImporter.isWordHtml(html)) {
      const imported = this.wordPasteImporter.importFromHtml(html);
      if (imported.pageSetup) this.setPageSetup(imported.pageSetup);
      return imported.html;
    }

    const temp = document.createElement("div");
    temp.innerHTML = html || "<p><br></p>";

    const body = temp.querySelector("body");
    if (body) temp.innerHTML = body.innerHTML;

    const savedDocument = this.getSavedDocumentWrapper(temp);
    if (savedDocument) {
      const savedPageSetup = readPageSetupFromElement(savedDocument);
      if (savedPageSetup) this.setPageSetup(savedPageSetup);
      temp.innerHTML = savedDocument.innerHTML;
    }

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

  private getSavedDocumentWrapper(root: HTMLElement): HTMLElement | null {
    const directChildren = this.getMeaningfulChildren(root, true);
    if (directChildren.length === 1 && directChildren[0].nodeType === Node.ELEMENT_NODE) {
      const onlyChild = directChildren[0] as HTMLElement;
      if (onlyChild.matches("[data-hwe-document='true']")) return onlyChild;
    }

    return root.querySelector<HTMLElement>("[data-hwe-document='true']");
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
      .replace(/position\s*:[^;]+;?/gi, "")
      .replace(/left\s*:[^;]+;?/gi, "")
      .replace(/right\s*:[^;]+;?/gi, "")
      .replace(/transform\s*:[^;]+;?/gi, "")
      .replace(/overflow\s*:[^;]+;?/gi, "")
      .replace(/overflow-x\s*:[^;]+;?/gi, "")
      .replace(/overflow-y\s*:[^;]+;?/gi, "")
      .replace(/tab-stops\s*:[^;]+;?/gi, "")
      .replace(/behavior\s*:[^;]+;?/gi, "")
      .replace(/-moz-binding\s*:[^;]+;?/gi, "")
      .replace(/url\s*\([^)]*\)\s*;?/gi, "")
      .replace(/expression\s*\([^)]*\)\s*;?/gi, "")
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
    const splittableTags = new Set([
      "DIV",
      "SECTION",
      "ARTICLE",
      "MAIN",
      "BODY",
      "CENTER",
      "HEADER",
      "FOOTER",
      "ASIDE",
      "NAV",
    ]);
    if (!splittableTags.has(element.tagName)) return false;
    if (element.classList.contains("hwe-page") || element.classList.contains("hwe-page-inner")) {
      return false;
    }

    return this.getMeaningfulChildren(element).length > 0;
  }

  private getMeaningfulChildren(
    container: HTMLElement,
    preserveEditableBlankBlocks = false
  ): ChildNode[] {
    return Array.from(container.childNodes).filter(
      (node) => !this.isEmptyNode(node, preserveEditableBlankBlocks)
    );
  }

  private isEmptyNode(node: ChildNode, preserveEditableBlankBlocks = false): boolean {
    if (node.nodeType === Node.COMMENT_NODE) return true;

    if (node.nodeType === Node.TEXT_NODE) {
      return (node.textContent ?? "").replace(/\u00a0/g, " ").trim() === "";
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return false;

    const element = node as HTMLElement;
    if (element.hasAttribute("data-hwe-user-blank")) return false;
    if (element.tagName === "BR") return true;
    if (["META", "LINK", "STYLE", "SCRIPT", "XML"].includes(element.tagName)) return true;
    if (element.querySelector("img, table, tr, td, th, video, canvas, svg")) return false;
    if (preserveEditableBlankBlocks && this.isEditableBlankBlock(element)) return false;

    return (
      ["P", "DIV", "SECTION", "ARTICLE", "SPAN"].includes(element.tagName) &&
      (element.textContent ?? "").replace(/\u00a0/g, " ").trim() === "" &&
      Array.from(element.childNodes).every((child) => this.isEmptyNode(child, false))
    );
  }

  private isEditableBlankBlock(element: HTMLElement): boolean {
    const blankBlockTags = new Set([
      "P",
      "DIV",
      "LI",
      "H1",
      "H2",
      "H3",
      "H4",
      "H5",
      "H6",
      "BLOCKQUOTE",
      "PRE",
    ]);

    if (!blankBlockTags.has(element.tagName)) return false;
    if ((element.textContent ?? "").replace(/\u00a0/g, " ").trim() !== "") return false;

    return Array.from(element.childNodes).every((child) => {
      if (child.nodeType === Node.TEXT_NODE) {
        return (child.textContent ?? "").replace(/\u00a0/g, " ").trim() === "";
      }

      return child.nodeType === Node.ELEMENT_NODE && (child as HTMLElement).tagName === "BR";
    });
  }

  private onPaste(event: ClipboardEvent, page: HTMLElement): void {
    const html = event.clipboardData?.getData("text/html") ?? "";
    if (!html || !this.wordPasteImporter.isWordHtml(html)) return;

    event.preventDefault();

    const targetEditable = event.currentTarget as HTMLElement;
    const insertionMarker = this.createPasteInsertionMarker(targetEditable);
    const imported = this.wordPasteImporter.importFromHtml(html);
    if (imported.pageSetup) this.setPageSetup(imported.pageSetup);

    const affectedPage =
      this.insertHtmlAtPasteMarker(imported.html, insertionMarker, targetEditable) ?? page;
    const inner = affectedPage.querySelector<HTMLElement>(".hwe-page-inner");
    if (inner) this.syncEditableBlankBlocks(inner, false);

    this.isDirty = true;
    this.toolbar.updateActiveStates();
    this.scheduleRebalance(affectedPage, true);
    void this.waitForImages(affectedPage).then(() => this.scheduleRebalance(affectedPage, true));
  }

  private createPasteInsertionMarker(targetEditable: HTMLElement): HTMLElement {
    const marker = document.createElement("span");
    marker.setAttribute("data-hwe-paste-marker", "true");
    marker.style.cssText = "display:inline-block;width:0;height:0;overflow:hidden;line-height:0;";

    const selection = window.getSelection();
    targetEditable.focus({ preventScroll: true });

    if (!selection || selection.rangeCount === 0) {
      targetEditable.appendChild(marker);
      return marker;
    }

    const range = selection.getRangeAt(0);
    if (!targetEditable.contains(range.commonAncestorContainer)) {
      targetEditable.appendChild(marker);
      return marker;
    }

    range.deleteContents();
    range.insertNode(marker);
    return marker;
  }

  private insertHtmlAtPasteMarker(
    html: string,
    marker: HTMLElement,
    targetEditable: HTMLElement
  ): HTMLElement | null {
    const template = document.createElement("template");
    template.innerHTML = html;
    const fragment = template.content;
    const insertedNodes = Array.from(fragment.childNodes);

    if (!marker.parentNode) {
      targetEditable.appendChild(fragment);
    } else {
      marker.replaceWith(fragment);
    }

    const selection = window.getSelection();
    const lastInserted = insertedNodes[insertedNodes.length - 1];
    if (selection && lastInserted?.parentNode) {
      const nextRange = document.createRange();
      nextRange.setStartAfter(lastInserted);
      nextRange.collapse(true);
      targetEditable.focus({ preventScroll: true });
      selection.removeAllRanges();
      selection.addRange(nextRange);
    }

    return targetEditable.closest<HTMLElement>(".hwe-page");
  }

  private syncEditableBlankBlocks(root: HTMLElement, markNewBlanks: boolean): void {
    root
      .querySelectorAll<HTMLElement>("p, div, li, h1, h2, h3, h4, h5, h6, blockquote, pre")
      .forEach((element) => {
        if (!this.isEditableBlankBlock(element)) {
          element.removeAttribute("data-hwe-user-blank");
          return;
        }

        if (markNewBlanks || element.hasAttribute("data-hwe-user-blank")) {
          element.setAttribute("data-hwe-user-blank", "true");
          this.ensureBlankBlockHasCaretStop(element);
        }
      });
  }

  private ensureBlankBlockHasCaretStop(element: HTMLElement): void {
    if (element.childNodes.length === 0) {
      element.appendChild(document.createElement("br"));
      return;
    }

    const hasBreak = Array.from(element.childNodes).some(
      (child) => child.nodeType === Node.ELEMENT_NODE && (child as HTMLElement).tagName === "BR"
    );
    if (!hasBreak && (element.textContent ?? "").length === 0) {
      element.appendChild(document.createElement("br"));
    }
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

    if (event.key === "Backspace" && this.handleBackspaceAtPageStart(event)) {
      return;
    }

    if (event.key === "Backspace" || event.key === "Delete") {
      const page = (event.currentTarget as HTMLElement).closest<HTMLElement>(".hwe-page");
      if (page) this.pagesNeedingPull.add(page);
    }
  }

  private handleBackspaceAtPageStart(event: KeyboardEvent): boolean {
    const inner = event.currentTarget as HTMLElement;
    const page = inner.closest<HTMLElement>(".hwe-page");
    if (!page || !this.isCaretAtStartOfEditable(inner)) return false;

    const pageIndex = this.pages.indexOf(page);
    if (pageIndex <= 0) return false;

    const previousPage = this.pages[pageIndex - 1];
    const previousInner = previousPage.querySelector<HTMLElement>(".hwe-page-inner");
    if (!previousInner) return false;

    event.preventDefault();
    this.deleteLastContent(previousInner);
    this.placeCaretAtEnd(previousInner);
    this.isDirty = true;
    this.toolbar.updateActiveStates();
    this.scheduleRebalance(previousPage, true);
    return true;
  }

  private isCaretAtStartOfEditable(inner: HTMLElement): boolean {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0 || !selection.isCollapsed) return false;

    const range = selection.getRangeAt(0);
    if (!inner.contains(range.startContainer)) return false;

    const beforeRange = document.createRange();
    beforeRange.selectNodeContents(inner);
    beforeRange.setEnd(range.startContainer, range.startOffset);

    const fragment = beforeRange.cloneContents();
    return !this.fragmentHasVisibleContent(fragment);
  }

  private fragmentHasVisibleContent(fragment: DocumentFragment): boolean {
    return Array.from(fragment.childNodes).some((node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        return (node.textContent ?? "").replace(/\u00a0/g, " ").length > 0;
      }

      if (node.nodeType !== Node.ELEMENT_NODE) return false;

      const element = node as HTMLElement;
      if (element.tagName === "BR") return false;
      if (element.querySelector("img, table, tr, td, th, video, canvas, svg")) return true;
      return (element.textContent ?? "").replace(/\u00a0/g, " ").length > 0;
    });
  }

  private deleteLastContent(container: HTMLElement): void {
    const removed = this.deleteLastContentFromNode(container);
    if (!removed && this.getMeaningfulChildren(container, true).length === 0) {
      container.innerHTML = "<p><br></p>";
    }
  }

  private deleteLastContentFromNode(node: Node): boolean {
    for (let index = node.childNodes.length - 1; index >= 0; index--) {
      const child = node.childNodes[index];

      if (child.nodeType === Node.TEXT_NODE) {
        const text = child.textContent ?? "";
        if (text.length === 0) {
          child.remove();
          continue;
        }

        child.textContent = text.slice(0, -1);
        if (child.textContent.length === 0) child.remove();
        return true;
      }

      if (child.nodeType !== Node.ELEMENT_NODE) {
        child.remove();
        return true;
      }

      const element = child as HTMLElement;
      if (element.hasAttribute("data-hwe-caret")) {
        element.remove();
        continue;
      }

      if (element.tagName === "BR" || this.isAtomicEditableElement(element)) {
        element.remove();
        return true;
      }

      if (this.deleteLastContentFromNode(element)) {
        if (this.isEmptyNode(element, false)) element.remove();
        return true;
      }

      if (this.isEditableBlankBlock(element)) {
        element.remove();
        return true;
      }
    }

    return false;
  }

  private isAtomicEditableElement(element: HTMLElement): boolean {
    return ["IMG", "TABLE", "VIDEO", "CANVAS", "SVG"].includes(element.tagName);
  }

  private placeCaretAtEnd(inner: HTMLElement): void {
    inner.focus({ preventScroll: true });

    const range = document.createRange();
    const endPosition = this.getLastCaretPosition(inner);
    if (endPosition) {
      range.setStart(endPosition.node, endPosition.offset);
    } else {
      range.selectNodeContents(inner);
      range.collapse(false);
    }

    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
  }

  private getLastCaretPosition(root: Node): { node: Node; offset: number } | null {
    for (let index = root.childNodes.length - 1; index >= 0; index--) {
      const child = root.childNodes[index];

      if (child.nodeType === Node.TEXT_NODE) {
        return { node: child, offset: child.textContent?.length ?? 0 };
      }

      if (child.nodeType !== Node.ELEMENT_NODE) continue;

      const element = child as HTMLElement;
      if (element.hasAttribute("data-hwe-caret")) continue;
      if (element.tagName === "BR") {
        return { node: root, offset: index };
      }

      const nested = this.getLastCaretPosition(element);
      if (nested) return nested;

      if (!this.isEmptyNode(element, false)) {
        return { node: root, offset: index + 1 };
      }
    }

    return null;
  }

  private isDeleteInput(inputType: string): boolean {
    return inputType.startsWith("delete") || inputType === "historyUndo";
  }

  private isEnterInput(inputType: string): boolean {
    return inputType === "insertParagraph" || inputType === "insertLineBreak";
  }

  private async save(): Promise<void> {
    const saveButton = this.toolbar.getSaveButton();
    saveButton.disabled = true;
    this.setStatus("Guardando...", "saving");

    try {
      const html = this.collectHtml();
      if (this.options.saveHtml) {
        await this.options.saveHtml(html);
      } else {
        await saveHtmlToFileField(
          this.baseUrl,
          this.entityName,
          this.entityId,
          this.fieldName,
          html
        );
      }

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
    CaretManager.removeMarkers(this.root);

    const html = this.pages
      .map((page, index) => {
        const inner = page.querySelector(".hwe-page-inner") as HTMLElement | null;
        const content = inner ? inner.innerHTML : page.innerHTML;
        if (index === 0) return content;
        return `<div data-hwe-page-break="before" style="page-break-before:always">${content}</div>`;
      })
      .join("\n");

    return makePageSetupWrapper(html, this.pageSetup);
  }

  private updatePageCount(): void {
    this.pageCountEl.textContent = `Paginas: ${this.pages.length}`;
  }

  private setStatus(message: string, type: StatusType): void {
    this.statusMsg.textContent = message;
    this.statusMsg.className = "hwe-status-msg" + (type ? ` ${type}` : "");
  }

  private getActiveEditable(): HTMLElement | null {
    const active = document.activeElement as HTMLElement | null;
    if (active?.matches("[contenteditable='true']") && this.root.contains(active)) {
      return active;
    }

    const selection = window.getSelection();
    const anchorNode = selection?.anchorNode;
    if (!anchorNode || !this.root.contains(anchorNode)) return null;

    const element =
      anchorNode.nodeType === Node.ELEMENT_NODE
        ? (anchorNode as HTMLElement)
        : anchorNode.parentElement;

    return element?.closest<HTMLElement>("[contenteditable='true']") ?? null;
  }

  private getEditableForPageIndex(pageIndex: number): HTMLElement | null {
    if (this.pages.length === 0) return null;

    const safeIndex = Math.max(0, Math.min(pageIndex, this.pages.length - 1));
    return this.pages[safeIndex].querySelector<HTMLElement>(".hwe-page-inner");
  }

  private shouldRebalancePage(
    page: HTMLElement,
    pageIndex: number,
    pullFromNextPages: boolean
  ): boolean {
    if (this.pageOverflows(page)) return true;
    return pullFromNextPages && pageIndex >= 0 && pageIndex < this.pages.length - 1;
  }

  private pageOverflows(page: HTMLElement): boolean {
    const inner = page.querySelector<HTMLElement>(".hwe-page-inner");
    if (!inner) return false;

    const contentBottom = this.getContentBottom(inner);
    if (contentBottom === null) return false;

    return contentBottom > this.getContentLimitBottom(inner) + 1;
  }

  private getContentLimitBottom(inner: HTMLElement): number {
    const styles = getComputedStyle(inner);
    const paddingBottom = parseFloat(styles.paddingBottom) || 0;
    return inner.getBoundingClientRect().bottom - paddingBottom;
  }

  private getContentBottom(inner: HTMLElement): number | null {
    const children = this.getMeaningfulChildren(inner, true);
    if (children.length === 0) return null;

    return children.reduce<number | null>((bottom, child) => {
      const childBottom = this.getNodeBottom(child);
      if (childBottom === null) return bottom;
      return bottom === null ? childBottom : Math.max(bottom, childBottom);
    }, null);
  }

  private getNodeBottom(node: ChildNode): number | null {
    if (node.nodeType === Node.ELEMENT_NODE) {
      return (node as HTMLElement).getBoundingClientRect().bottom;
    }

    const range = document.createRange();
    range.selectNodeContents(node);
    const rect = range.getBoundingClientRect();
    return rect.width > 0 || rect.height > 0 ? rect.bottom : null;
  }

  private shouldPreserveWorkspaceScroll(): boolean {
    const active = document.activeElement;
    return !!active && this.root.contains(active);
  }

  private restoreWorkspaceScroll(scrollTop: number, scrollLeft: number): void {
    this.workspace.scrollTop = scrollTop;
    this.workspace.scrollLeft = scrollLeft;
  }

  private restoreCaretViewport(marker: HTMLElement | null, previousTop: number | null): void {
    if (!marker || !marker.parentNode || previousTop === null) return;

    const currentTop = marker.getBoundingClientRect().top;
    this.workspace.scrollTop += currentTop - previousTop;
  }

  async loadHtml(html: string): Promise<void> {
    await this.renderAndPaginate(html || "<p><br></p>");
    this.isDirty = false;
    this.setStatus("", "");
  }

  getHtml(): string {
    return this.collectHtml();
  }

  destroy(): void {
    if (this.rebalanceFrame !== undefined) {
      window.cancelAnimationFrame(this.rebalanceFrame);
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
