import { Paginator, RebalanceOptions } from "../Orquestador/Paginator";
import { CaretManager } from "../Orquestador/CaretManager";
import {
  CLEAR_PARAGRAPH_STYLE_VALUE,
  ParagraphStyleOption,
  Toolbar,
} from "../Resize/Toolbar";
import { fetchHtmlFromFileField, saveHtmlToFileField } from "../execCommand/fileApi";
import {
  fetchParagraphStyles,
  ParagraphStyleDefinition,
  ParagraphStyleTableConfig,
} from "../execCommand/styleApi";
import {
  applyPageSetup,
  DEFAULT_PAGE_SETUP,
  normalizePageSetup,
  PageSetup,
} from "../Orquestador/PageGeometry";
import { BlankLineController } from "./BlankLineController";
import { DocumentSerializer } from "./DocumentSerializer";
import { PasteController } from "./PasteController";
import { AssetLayoutManager } from "./AssetLayoutManager";
import {
  getMeaningfulChildren,
  isEditableBlankBlock,
  isEmptyNode,
} from "../dom/EditableDom";
import { ImageOcrService } from "../ocr/ImageOcrService";

type PcfContext = ComponentFramework.Context<IInputs>;
type StatusType = "success" | "error" | "saving" | "";
type EditorView = "visual" | "source";
type QueuedRebalance = Required<RebalanceOptions> & {
  page: HTMLElement;
  pullFromNextPages: boolean;
};

const PARAGRAPH_STYLE_BLOCK_SELECTOR = "p, li, h1, h2, h3, h4, h5, h6, blockquote, pre";
const DEFAULT_STYLE_TABLE_CONFIG: ParagraphStyleTableConfig = {
  entitySetName: "mcdev_htmlstyles",
  classField: "mcdev_cssclass",
  cssField: "mcdev_css",
};
const LOCAL_PARAGRAPH_STYLES: ParagraphStyleDefinition[] = [
  {
    label: "Texto general",
    className: "texto-general",
    cssText: `.texto-general {
  display: inline-block;
  text-indent: 20pt;
  margin: 0;
  text-align: justify;
  hyphens: auto;
  -webkit-hyphens: auto;
  -ms-hyphens: auto;
  orphans: 2;
  widows: 2;
  font-family: 'Swis721 BT','SwissRoman', Helvetica, Arial, sans-serif;
  font-weight: normal;
  font-style: normal;
  font-size: 11pt;
}`,
  },
];

export interface EditorComponentOptions {
  initialHtml?: string;
  loadHtml?: () => Promise<string> | string;
  saveHtml?: (html: string) => Promise<void> | void;
  paragraphStyles?: ParagraphStyleDefinition[];
}

export class EditorComponent {
  private readonly container: HTMLElement;
  private readonly options: EditorComponentOptions;

  private root!: HTMLElement;
  private editorHeader!: HTMLElement;
  private workspace!: HTMLElement;
  private viewTabs!: HTMLElement;
  private visualTabBtn!: HTMLButtonElement;
  private sourceTabBtn!: HTMLButtonElement;
  private sourceEditor!: HTMLTextAreaElement;
  private statusMsg!: HTMLElement;
  private pageCountEl!: HTMLElement;

  private pages: HTMLElement[] = [];
  private paginator!: Paginator;
  private toolbar!: Toolbar;
  private readonly blankLineController = new BlankLineController();
  private readonly documentSerializer = new DocumentSerializer();
  private readonly assetLayoutManager = new AssetLayoutManager();
  private readonly pasteController = new PasteController();
  private readonly imageOcrService = new ImageOcrService({
    setStatus: (message, type) => this.setStatus(message, type),
  });
  private pageSetup: PageSetup = DEFAULT_PAGE_SETUP;

  private readonly baseUrl: string;
  private readonly entityName: string;
  private readonly entityId: string;
  private readonly fieldName: string;
  private readonly styleTableConfig: ParagraphStyleTableConfig;
  private currentFileName = "content.html";

  private rebalanceFrame: number | undefined;
  private pendingRebalance: QueuedRebalance | null = null;
  private readonly pagesNeedingPull = new WeakSet<HTMLElement>();
  private readonly pendingInputTypes = new WeakMap<HTMLElement, string>();
  private readonly paragraphStyleClasses = new Set<string>();
  private lastSelectedTableRow: HTMLTableRowElement | null = null;
  private lastTextSelection: Range | null = null;
  private lastStyleBlock: HTMLElement | null = null;
  private paragraphStyleElement: HTMLStyleElement | null = null;
  private readonly handleSelectionChange = (): void => this.rememberTextSelection();
  private isComposing = false;
  private isDirty = false;
  private activeView: EditorView = "visual";
  private sourceDirty = false;

  constructor(container: HTMLElement, context?: PcfContext, options: EditorComponentOptions = {}) {
    this.container = container;
    this.options = options;

    const runtime = (context ?? {}) as unknown as {
      page?: { getClientUrl?: () => string; entityId?: string };
      mode?: { contextInfo?: { entityId?: string; entityTypeName?: string } };
      parameters?: Record<string, { raw?: string | null }>;
    };

    this.baseUrl = this.getClientUrl(runtime);
    this.entityId = this.cleanGuid(
      runtime.page?.entityId ?? runtime.mode?.contextInfo?.entityId ?? ""
    );
    this.entityName = "mcdev_htmldevtests";
    this.fieldName = "mcdev_htmlarchivooriginal";
    this.styleTableConfig = {
      entitySetName:
        this.getParameterValue(runtime.parameters, "styleEntitySetName") ??
        DEFAULT_STYLE_TABLE_CONFIG.entitySetName,
      classField:
        this.getParameterValue(runtime.parameters, "styleClassField") ??
        DEFAULT_STYLE_TABLE_CONFIG.classField,
      cssField:
        this.getParameterValue(runtime.parameters, "styleCssField") ??
        DEFAULT_STYLE_TABLE_CONFIG.cssField,
    };
  }

  async init(): Promise<void> {
    this.buildShell();
    void this.loadParagraphStyles();
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

    this.editorHeader = document.createElement("div");
    this.editorHeader.className = "hwe-editor-header";
    this.root.appendChild(this.editorHeader);

    this.toolbar = new Toolbar({
      onInsertTable: () => this.insertTable(),
      onInsertRowAfter: () => this.insertTableRowAfter(),
      onApplyParagraphStyle: (className) => this.applyParagraphStyle(className),
      onExportPdf: () => {
        void this.exportPdf();
      },
    });
    const toolbarEl = this.toolbar.build();
    this.toolbar.getSaveButton().addEventListener("click", () => void this.save());
    this.editorHeader.appendChild(toolbarEl);
    document.addEventListener("selectionchange", this.handleSelectionChange);

    this.buildViewTabs();

    this.workspace = document.createElement("div");
    this.workspace.className = "hwe-workspace";
    this.root.appendChild(this.workspace);

    this.sourceEditor = document.createElement("textarea");
    this.sourceEditor.className = "hwe-source-editor";
    this.sourceEditor.setAttribute("spellcheck", "false");
    this.sourceEditor.addEventListener("input", () => {
      this.sourceDirty = true;
      this.isDirty = true;
    });
    this.root.appendChild(this.sourceEditor);
    this.updateViewTabs();

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

  private buildViewTabs(): void {
    this.viewTabs = document.createElement("div");
    this.viewTabs.className = "hwe-view-tabs";

    this.visualTabBtn = this.makeViewTabButton("Editor", "visual");
    this.sourceTabBtn = this.makeViewTabButton("HTML", "source");

    this.viewTabs.appendChild(this.visualTabBtn);
    this.viewTabs.appendChild(this.sourceTabBtn);
    this.editorHeader.appendChild(this.viewTabs);
    this.updateViewTabs();
  }

  private makeViewTabButton(label: string, view: EditorView): HTMLButtonElement {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = label;
    button.addEventListener("click", () => {
      void this.switchView(view);
    });
    return button;
  }

  private async switchView(view: EditorView): Promise<void> {
    if (view === this.activeView) return;

    if (view === "source") {
      this.sourceEditor.value = this.collectHtml();
      this.sourceDirty = false;
      this.activeView = "source";
      this.updateViewTabs();
      this.sourceEditor.focus();
      return;
    }

    if (this.sourceDirty) {
      await this.renderAndPaginate(this.sourceEditor.value || "<p><br></p>");
      this.sourceDirty = false;
    }

    this.activeView = "visual";
    this.updateViewTabs();
  }

  private updateViewTabs(): void {
    this.visualTabBtn.classList.toggle("hwe-active", this.activeView === "visual");
    this.sourceTabBtn.classList.toggle("hwe-active", this.activeView === "source");

    if (this.workspace) {
      this.workspace.hidden = this.activeView !== "visual";
    }
    if (this.sourceEditor) {
      this.sourceEditor.hidden = this.activeView !== "source";
    }
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

  private getParameterValue(
    parameters: Record<string, { raw?: string | null }> | undefined,
    name: string
  ): string | undefined {
    const value = parameters?.[name]?.raw?.trim();
    return value || undefined;
  }

  private async loadParagraphStyles(): Promise<void> {
    try {
      const styles = this.options.paragraphStyles
        ? this.options.paragraphStyles
        : this.baseUrl
          ? await fetchParagraphStyles(this.baseUrl, this.styleTableConfig)
          : LOCAL_PARAGRAPH_STYLES;

      this.setParagraphStyles(styles.length > 0 ? styles : LOCAL_PARAGRAPH_STYLES);
    } catch (error) {
      console.warn("[HtmlWordEditor] paragraph styles fallback:", error);
      this.setParagraphStyles(LOCAL_PARAGRAPH_STYLES);
    }
  }

  private setParagraphStyles(styles: ParagraphStyleDefinition[]): void {
    const validStyles = styles.filter((style) => this.isValidCssClassName(style.className));

    this.paragraphStyleClasses.clear();
    validStyles.forEach((style) => this.paragraphStyleClasses.add(style.className));
    this.toolbar.setParagraphStyles(
      validStyles.map<ParagraphStyleOption>((style) => ({
        label: style.label,
        className: style.className,
      }))
    );
    this.injectParagraphStyleCss(validStyles);
  }

  private injectParagraphStyleCss(styles: ParagraphStyleDefinition[]): void {
    if (!this.paragraphStyleElement) {
      this.paragraphStyleElement = document.createElement("style");
      this.paragraphStyleElement.setAttribute("data-hwe-style-catalog", "true");
      this.root.insertBefore(this.paragraphStyleElement, this.root.firstChild);
    }

    this.paragraphStyleElement.textContent = styles
      .map((style) => this.formatParagraphStyleCss(style))
      .join("\n\n");
  }

  private formatParagraphStyleCss(style: ParagraphStyleDefinition): string {
    const declarations = this.extractCssDeclarations(style.cssText);
    return declarations ? `.hwe-page-inner .${style.className} { ${declarations} }` : "";
  }

  private extractCssDeclarations(cssText: string): string {
    const css = cssText.trim();
    if (!css) return "";

    if (css.startsWith("{") && css.endsWith("}")) {
      return css.slice(1, -1).trim();
    }

    if (css.includes("{")) {
      const ruleMatch = /[^{]+\{([\s\S]*?)\}/.exec(css);
      return ruleMatch?.[1]?.trim() ?? "";
    }

    return css;
  }

  private isValidCssClassName(className: string): boolean {
    return /^-?[_a-zA-Z]+[_a-zA-Z0-9-]*$/.test(className);
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

      const fileContent = await fetchHtmlFromFileField(
        this.baseUrl,
        this.entityName,
        this.entityId,
        this.fieldName
      );
      this.currentFileName = fileContent.fileName || this.currentFileName;

      await this.renderAndPaginate(fileContent.html || "<p><br></p>");
      this.setStatus("", "");
    } catch (err) {
      this.setStatus(`Error al cargar: ${(err as Error).message}`, "error");
      await this.renderAndPaginate("<p><br></p>");
    }
  }

  private async renderAndPaginate(html: string): Promise<void> {
    this.workspace.innerHTML = "";
    this.setPageSetup(DEFAULT_PAGE_SETUP);

    const normalizedDocument = this.documentSerializer.normalizeHtmlForPagination(html);

    this.pages = [this.createPageElement(normalizedDocument.html)];
    this.workspace.appendChild(this.pages[0]);
    this.applyOfficialTableWidths(this.workspace);

    this.paginator.setPages(this.pages);
    await this.assetLayoutManager.waitForStableLayout(this.workspace);

    this.paginator.repaginateAll();
    this.pages = this.paginator.getPages();
    this.pages.forEach((page) => this.applyOfficialTableWidths(page));
    this.syncWorkspace();
    this.updatePageCount();
    this.imageOcrService.queue(this.workspace);
  }

  private createPageElement(html?: string): HTMLElement {
    const page = document.createElement("div");
    page.className = "hwe-page";

    const inner = document.createElement("div");
    inner.className = "hwe-page-inner";
    inner.setAttribute("contenteditable", "true");
    inner.setAttribute("spellcheck", "false");
    inner.innerHTML = html ?? "<p><br></p>";
    this.applyOfficialTableWidths(inner);

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
        this.blankLineController.syncEditableBlankBlocks(inner, this.isEnterInput(inputType));
        const shouldPullFromNextPages =
          this.isDeleteInput(inputType) || this.pagesNeedingPull.has(page);
        this.pagesNeedingPull.delete(page);
        this.scheduleRebalance(page, shouldPullFromNextPages, {
          includePreviousPage: shouldPullFromNextPages,
        });
      }
    });
    inner.addEventListener("compositionstart", () => {
      this.isComposing = true;
    });
    inner.addEventListener("compositionend", () => {
      this.isComposing = false;
      this.blankLineController.syncEditableBlankBlocks(inner, false);
      this.scheduleRebalance(page, false, { includePreviousPage: false });
    });
    inner.addEventListener("paste", (event: ClipboardEvent) => this.onPaste(event, page));
    inner.addEventListener("keydown", (event: KeyboardEvent) => this.onPageKeyDown(event));
    inner.addEventListener("keyup", () => {
      this.rememberTextSelection();
      this.rememberSelectedTableRow();
    });
    inner.addEventListener("click", (event) => {
      this.rememberTableRowFromEvent(event);
      this.rememberStyleBlockFromEvent(event);
    });
    inner.addEventListener("mouseup", (event) => {
      this.rememberStyleBlockFromEvent(event);
      this.rememberTextSelection();
      this.rememberSelectedTableRow();
      this.toolbar.updateActiveStates();
    });

    page.appendChild(inner);
    return page;
  }

  private scheduleRebalance(
    page: HTMLElement,
    pullFromNextPages = false,
    options: RebalanceOptions = {}
  ): void {
    this.queueRebalance(page, pullFromNextPages, options);
    if (this.rebalanceFrame !== undefined) return;

    this.rebalanceFrame = window.requestAnimationFrame(() => {
      this.rebalanceFrame = undefined;

      const pending = this.pendingRebalance;
      this.pendingRebalance = null;
      if (!pending || !this.pages.includes(pending.page)) return;

      const pageIndex = this.pages.indexOf(pending.page);
      if (!this.shouldRebalancePage(pending.page, pageIndex, pending.pullFromNextPages)) return;

      const startPageIndex = pending.includePreviousPage ? Math.max(0, pageIndex - 1) : pageIndex;
      const startPage = this.pages[startPageIndex] ?? pending.page;
      const activeEditable = this.getActiveEditable();
      const marker = CaretManager.createMarker(this.root);
      const caretViewportTop = marker?.getBoundingClientRect().top ?? null;
      this.paginator.rebalanceFromPage(startPage, {
        includePreviousPage: false,
        compactPages: pending.compactPages,
      });
      this.pages = this.paginator.getPages();
      this.pages.forEach((currentPage) => {
        const inner = currentPage.querySelector<HTMLElement>(".hwe-page-inner");
        if (inner) this.blankLineController.syncEditableBlankBlocks(inner, false);
      });
      this.syncWorkspace();
      this.updatePageCount();
      const fallbackEditable = this.getEditableForPageIndex(pageIndex) ?? activeEditable;
      this.restoreCaretViewport(marker, caretViewportTop);
      CaretManager.restoreMarker(marker, fallbackEditable);
      CaretManager.removeMarkers(this.root);
    });
  }

  private queueRebalance(
    page: HTMLElement,
    pullFromNextPages: boolean,
    options: RebalanceOptions
  ): void {
    const nextRebalance: QueuedRebalance = {
      page,
      pullFromNextPages,
      includePreviousPage: options.includePreviousPage ?? pullFromNextPages,
      compactPages: options.compactPages ?? true,
    };

    if (!this.pendingRebalance) {
      this.pendingRebalance = nextRebalance;
      return;
    }

    const currentIndex = this.pages.indexOf(this.pendingRebalance.page);
    const nextIndex = this.pages.indexOf(page);
    const shouldUseNextPage =
      currentIndex === -1 || (nextIndex !== -1 && nextIndex < currentIndex);

    this.pendingRebalance = {
      page: shouldUseNextPage ? page : this.pendingRebalance.page,
      pullFromNextPages: this.pendingRebalance.pullFromNextPages || pullFromNextPages,
      includePreviousPage:
        this.pendingRebalance.includePreviousPage || nextRebalance.includePreviousPage,
      compactPages: this.pendingRebalance.compactPages || nextRebalance.compactPages,
    };
  }

  private onPagesChanged(pages: HTMLElement[]): void {
    this.pages = pages;
    this.pages.forEach((page) => this.applyOfficialTableWidths(page));
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

  private applyOfficialTableWidths(root: HTMLElement): void {
    const inners = root.classList.contains("hwe-page-inner")
      ? [root]
      : Array.from(root.querySelectorAll<HTMLElement>(".hwe-page-inner"));

    inners.forEach((inner) => {
      Array.from(inner.children).forEach((child) => {
        const table = this.getDirectFlowTable(child as HTMLElement);
        if (!table) return;

        table.style.setProperty("width", "100%", "important");
        table.style.setProperty("max-width", "100%", "important");
        table.style.setProperty("margin-left", "0", "important");
        table.style.setProperty("margin-right", "0", "important");
      });
    });
  }

  private getDirectFlowTable(element: HTMLElement): HTMLTableElement | null {
    if (element.tagName === "TABLE") return element as HTMLTableElement;

    const children = getMeaningfulChildren(element, true);
    if (children.length !== 1 || children[0].nodeType !== Node.ELEMENT_NODE) return null;

    const onlyChild = children[0] as HTMLElement;
    return onlyChild.tagName === "TABLE" ? (onlyChild as HTMLTableElement) : null;
  }

  private onPaste(event: ClipboardEvent, page: HTMLElement): void {
    const result = this.pasteController.handlePaste(event, page);
    if (!result.handled) return;

    const affectedPage = result.affectedPage ?? page;
    const inner = affectedPage.querySelector<HTMLElement>(".hwe-page-inner");
    if (inner) this.blankLineController.syncEditableBlankBlocks(inner, false);
    this.applyOfficialTableWidths(affectedPage);

    this.isDirty = true;
    this.toolbar.updateActiveStates();
    this.scheduleRebalance(affectedPage, true, { includePreviousPage: false });
    void this.assetLayoutManager
      .waitForStableLayout(affectedPage)
      .then(() => {
        this.scheduleRebalance(affectedPage, true, { includePreviousPage: false });
        this.imageOcrService.queue(affectedPage);
      });
  }

  private applyParagraphStyle(className: string): void {
    const shouldClearStyle = className === CLEAR_PARAGRAPH_STYLE_VALUE;
    if (!shouldClearStyle && !this.paragraphStyleClasses.has(className)) return;
    if (this.activeView !== "visual") {
      this.setStatus("Vuelve al editor visual para aplicar estilos.", "error");
      return;
    }

    this.restoreTextSelection();
    const blocks = this.getSelectedStyleBlocks();
    if (blocks.length === 0) {
      this.setStatus("Selecciona un parrafo para aplicar el estilo.", "error");
      return;
    }

    blocks.forEach((block) => {
      this.paragraphStyleClasses.forEach((styleClass) => block.classList.remove(styleClass));
      if (!shouldClearStyle) block.classList.add(className);
    });

    this.rememberTextSelection();
    this.markEditedAfterStyleChange(blocks[0]);
  }

  private markEditedAfterStyleChange(element: HTMLElement): void {
    const page = element.closest<HTMLElement>(".hwe-page");
    if (!page) return;

    this.isDirty = true;
    this.toolbar.updateActiveStates();

    if (this.pageOverflows(page)) {
      this.scheduleRebalance(page, false, {
        includePreviousPage: false,
        compactPages: false,
      });
    }
  }

  private getSelectedStyleBlocks(): HTMLElement[] {
    const selection = window.getSelection();
    const range = selection && selection.rangeCount > 0 ? selection.getRangeAt(0) : null;
    if (!range || !this.root.contains(range.commonAncestorContainer)) {
      return this.lastStyleBlock && this.root.contains(this.lastStyleBlock)
        ? [this.lastStyleBlock]
        : [];
    }

    const editable = this.getActiveEditable();
    const scope = editable ?? this.root;
    const blocks = Array.from(
      scope.querySelectorAll<HTMLElement>(PARAGRAPH_STYLE_BLOCK_SELECTOR)
    ).filter((block) => this.rangeOverlapsBlock(range, block));

    if (blocks.length > 0) {
      this.lastStyleBlock = blocks[0];
      return blocks;
    }

    const element =
      range.startContainer.nodeType === Node.ELEMENT_NODE
        ? (range.startContainer as HTMLElement)
        : range.startContainer.parentElement;
    const currentBlock = element?.closest<HTMLElement>(PARAGRAPH_STYLE_BLOCK_SELECTOR);
    if (currentBlock && this.root.contains(currentBlock)) {
      this.lastStyleBlock = currentBlock;
      return [currentBlock];
    }

    return this.lastStyleBlock && this.root.contains(this.lastStyleBlock)
      ? [this.lastStyleBlock]
      : [];
  }

  private rangeOverlapsBlock(range: Range, block: HTMLElement): boolean {
    if (range.collapsed) return block.contains(range.startContainer);

    const blockRange = document.createRange();
    blockRange.selectNodeContents(block);

    const startsBeforeBlockEnds =
      range.compareBoundaryPoints(Range.START_TO_END, blockRange) < 0;
    const endsAfterBlockStarts =
      range.compareBoundaryPoints(Range.END_TO_START, blockRange) > 0;
    return startsBeforeBlockEnds && endsAfterBlockStarts;
  }

  private rememberTextSelection(): void {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;

    const range = selection.getRangeAt(0);
    if (!this.root?.contains(range.commonAncestorContainer)) return;

    const element =
      range.commonAncestorContainer.nodeType === Node.ELEMENT_NODE
        ? (range.commonAncestorContainer as HTMLElement)
        : range.commonAncestorContainer.parentElement;
    if (!element?.closest("[contenteditable='true']")) return;

    this.lastTextSelection = range.cloneRange();
    this.lastStyleBlock = element.closest<HTMLElement>(PARAGRAPH_STYLE_BLOCK_SELECTOR);
  }

  private rememberStyleBlockFromEvent(event: Event): void {
    const target = event.target as HTMLElement | null;
    const block = target?.closest<HTMLElement>(PARAGRAPH_STYLE_BLOCK_SELECTOR);
    if (block && this.root.contains(block)) this.lastStyleBlock = block;
  }

  private restoreTextSelection(): void {
    if (!this.lastTextSelection) return;

    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(this.lastTextSelection);
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
    if (!page) return false;

    const caretAtPageStart = this.isCaretAtStartOfEditable(inner);
    const caretAtTableContinuationStart =
      !caretAtPageStart && this.isCaretAtStartOfFirstTableFragment(inner);
    if (!caretAtPageStart && !caretAtTableContinuationStart) return false;

    const pageIndex = this.pages.indexOf(page);
    if (pageIndex <= 0) return false;

    const previousPage = this.pages[pageIndex - 1];
    const previousInner = previousPage.querySelector<HTMLElement>(".hwe-page-inner");
    if (!previousInner) return false;

    event.preventDefault();
    if (caretAtTableContinuationStart) {
      this.removeTrailingPaginationBlanks(previousInner);
    } else {
      this.deleteLastContent(previousInner);
    }
    this.placeCaretAtEnd(previousInner);
    this.isDirty = true;
    this.toolbar.updateActiveStates();
    this.scheduleRebalance(previousPage, true, { includePreviousPage: true });
    return true;
  }

  private isCaretAtStartOfFirstTableFragment(inner: HTMLElement): boolean {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0 || !selection.isCollapsed) return false;

    const range = selection.getRangeAt(0);
    if (!inner.contains(range.startContainer)) return false;

    const first = getMeaningfulChildren(inner)[0];
    if (!first || first.nodeType !== Node.ELEMENT_NODE) return false;

    const firstElement = first as HTMLElement;
    if (firstElement.tagName !== "TABLE" || !firstElement.contains(range.startContainer)) {
      return false;
    }

    const beforeRange = document.createRange();
    beforeRange.selectNodeContents(firstElement);
    beforeRange.setEnd(range.startContainer, range.startOffset);

    return !this.fragmentHasTextOrMediaContent(beforeRange.cloneContents());
  }

  private fragmentHasTextOrMediaContent(fragment: DocumentFragment): boolean {
    return Array.from(fragment.childNodes).some((node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        return (node.textContent ?? "").replace(/\u00a0/g, " ").trim().length > 0;
      }

      if (node.nodeType !== Node.ELEMENT_NODE) return false;

      const element = node as HTMLElement;
      if (element.tagName === "BR") return false;
      if (element.querySelector("img, video, canvas, svg")) return true;
      return (element.textContent ?? "").replace(/\u00a0/g, " ").trim().length > 0;
    });
  }

  private removeTrailingPaginationBlanks(container: HTMLElement): boolean {
    let removedAny = false;
    let child = container.lastChild;

    while (child) {
      const previous = child.previousSibling;
      if (!this.isRemovablePaginationBlank(child)) break;

      child.remove();
      removedAny = true;
      child = previous;
    }

    return removedAny;
  }

  private isRemovablePaginationBlank(node: ChildNode): boolean {
    if (node.nodeType === Node.TEXT_NODE) {
      return (node.textContent ?? "").replace(/\u00a0/g, " ").trim() === "";
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return false;

    const element = node as HTMLElement;
    if (element.hasAttribute("data-hwe-user-blank")) return false;
    return isEmptyNode(element, false) || isEditableBlankBlock(element);
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
    if (!removed && getMeaningfulChildren(container, true).length === 0) {
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
        if (isEmptyNode(element, false)) element.remove();
        return true;
      }

      if (isEditableBlankBlock(element)) {
        element.remove();
        return true;
      }
    }

    return false;
  }

  private isAtomicEditableElement(element: HTMLElement): boolean {
    return ["IMG", "TABLE", "VIDEO", "CANVAS", "SVG"].includes(element.tagName);
  }

  private insertTable(): void {
    const editable = this.getActiveEditable() ?? this.getEditableForPageIndex(0);
    if (!editable) return;

    editable.focus({ preventScroll: true });
    document.execCommand(
      "insertHTML",
      false,
      '<table class="hwe-word-table" data-hwe-source="manual" data-hwe-table="word" style="border-collapse:collapse;width:100%"><tbody><tr><td style="border:solid windowtext 1.0pt;padding:2.85pt 4.25pt"><p><br></p></td><td style="border:solid windowtext 1.0pt;padding:2.85pt 4.25pt"><p><br></p></td></tr><tr><td style="border:solid windowtext 1.0pt;padding:2.85pt 4.25pt"><p><br></p></td><td style="border:solid windowtext 1.0pt;padding:2.85pt 4.25pt"><p><br></p></td></tr></tbody></table><p><br></p>'
    );
    this.markEditedAndRebalance(editable);
  }

  private insertTableRowAfter(): void {
    const selectionElement = this.getSelectionElement();
    const row = selectionElement?.closest("tr") ?? this.lastSelectedTableRow;
    if (!row || !this.root.contains(row)) {
      this.insertTable();
      return;
    }

    const newRow = row.cloneNode(true) as HTMLTableRowElement;
    Array.from(newRow.cells).forEach((cell) => {
      cell.innerHTML = "<p><br></p>";
    });
    row.parentNode?.insertBefore(newRow, row.nextSibling);
    this.lastSelectedTableRow = newRow;
    this.placeCaretInElement(newRow.cells[0] as HTMLElement);
    this.markEditedAndRebalance(row);
  }

  private rememberSelectedTableRow(): void {
    const row = this.getSelectionElement()?.closest("tr");
    if (row && this.root.contains(row)) this.lastSelectedTableRow = row;
  }

  private rememberTableRowFromEvent(event: Event): void {
    const target = event.target;
    if (!(target instanceof Element)) return;

    const row = target.closest("tr");
    if (row && this.root.contains(row)) this.lastSelectedTableRow = row as HTMLTableRowElement;
  }

  private getSelectionElement(): HTMLElement | null {
    const selection = window.getSelection();
    const node = selection?.anchorNode;
    if (!node) return null;
    return node.nodeType === Node.ELEMENT_NODE
      ? (node as HTMLElement)
      : node.parentElement;
  }

  private markEditedAndRebalance(element: HTMLElement): void {
    const page = element.closest<HTMLElement>(".hwe-page");
    if (!page) return;
    this.isDirty = true;
    this.toolbar.updateActiveStates();
    this.scheduleRebalance(page, true, { includePreviousPage: false });
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

  private placeCaretInElement(element: HTMLElement): void {
    const editable = element.closest<HTMLElement>("[contenteditable='true']");
    editable?.focus({ preventScroll: true });

    const range = document.createRange();
    range.selectNodeContents(element);
    range.collapse(false);

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

      if (!isEmptyNode(element, false)) {
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
      if (this.activeView === "source" && this.sourceDirty) {
        await this.renderAndPaginate(this.sourceEditor.value || "<p><br></p>");
        this.sourceEditor.value = this.collectHtml();
        this.sourceDirty = false;
        this.activeView = "source";
        this.updateViewTabs();
      }

      await this.imageOcrService.run(this.workspace, true);

      const html = this.collectHtml();
      if (this.options.saveHtml) {
        await this.options.saveHtml(html);
      } else {
        await saveHtmlToFileField(
          this.baseUrl,
          this.entityName,
          this.entityId,
          this.fieldName,
          html,
          this.currentFileName
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

  private async exportPdf(): Promise<void> {
    if (this.activeView === "source" && this.sourceDirty) {
      this.setStatus("Vuelve al editor visual o guarda los cambios del HTML antes de exportar.", "error");
      return;
    }

    await this.imageOcrService.run(this.workspace, true);

    const html = this.documentSerializer.collectPdfHtml(
      this.root,
      this.pages,
      this.pageSetup,
      this.paragraphStyleElement?.textContent ?? ""
    );
    const frame = document.createElement("iframe");
    frame.title = "Exportar PDF";
    frame.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0;";
    frame.setAttribute("aria-hidden", "true");
    this.root.appendChild(frame);

    const doc = frame.contentDocument;
    if (!doc) {
      frame.remove();
      this.setStatus("No se pudo preparar la exportacion a PDF.", "error");
      return;
    }

    doc.open();
    doc.write(html);
    doc.close();

    window.setTimeout(() => {
      frame.contentWindow?.focus();
      frame.contentWindow?.print();
      window.setTimeout(() => frame.remove(), 1000);
    }, 250);
    this.setStatus("Selecciona Guardar como PDF en el dialogo de impresion.", "success");
  }

  private collectHtml(): string {
    return this.documentSerializer.collectHtml(
      this.root,
      this.pages,
      this.pageSetup
    );
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
    return inner.getBoundingClientRect().bottom - Math.min(paddingBottom, 12);
  }

  private getContentBottom(inner: HTMLElement): number | null {
    const children = getMeaningfulChildren(inner);
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
    if (this.activeView === "source") {
      this.sourceEditor.value = this.collectHtml();
      this.sourceDirty = false;
    }
    this.isDirty = false;
    this.setStatus("", "");
  }

  getHtml(): string {
    if (this.activeView === "source" && this.sourceDirty) {
      return this.sourceEditor.value;
    }

    return this.collectHtml();
  }

  destroy(): void {
    if (this.rebalanceFrame !== undefined) {
      window.cancelAnimationFrame(this.rebalanceFrame);
    }
    document.removeEventListener("selectionchange", this.handleSelectionChange);
    this.imageOcrService.destroy();
    this.toolbar?.destroy();
    this.paginator?.destroy();
    this.container.innerHTML = "";
  }
}

interface IInputs {
  htmlContent: ComponentFramework.PropertyTypes.StringProperty;
  entityName: ComponentFramework.PropertyTypes.StringProperty;
  fieldName: ComponentFramework.PropertyTypes.StringProperty;
  styleEntitySetName: ComponentFramework.PropertyTypes.StringProperty;
  styleClassField: ComponentFramework.PropertyTypes.StringProperty;
  styleCssField: ComponentFramework.PropertyTypes.StringProperty;
}
