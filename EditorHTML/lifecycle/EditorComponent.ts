import { Paginator, RebalanceOptions } from "../Orquestador/Paginator";
import { CaretManager } from "../Orquestador/CaretManager";
import {
  CLEAR_PARAGRAPH_STYLE_VALUE,
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
import { EditorDiagnosticsController } from "./EditorDiagnosticsController";
import { EditorLayoutService } from "./EditorLayoutService";
import { ImageResizeController } from "./ImageResizeController";
import { PageBackspaceController } from "./PageBackspaceController";
import { ParagraphStyleManager } from "./ParagraphStyleManager";
import { StyleSelectionTracker } from "./StyleSelectionTracker";
import { TableCommandController } from "./TableCommandController";
import { EditorView, EditorViewController } from "./EditorViewController";
import {
  hweDebugLog,
  hweDebugStart,
  installHweDebugGlobals,
} from "../debug/DebugLogger";

type PcfContext = ComponentFramework.Context<IInputs>;
type StatusType = "success" | "error" | "saving" | "";
type QueuedRebalance = Required<RebalanceOptions> & {
  page: HTMLElement;
  pullFromNextPages: boolean;
};

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
  private readonly layoutService = new EditorLayoutService();
  private imageResizeController!: ImageResizeController;
  private readonly pageBackspaceController = new PageBackspaceController();
  private diagnosticsController!: EditorDiagnosticsController;
  private paragraphStyleManager!: ParagraphStyleManager;
  private styleSelectionTracker!: StyleSelectionTracker;
  private tableCommandController!: TableCommandController;
  private viewController!: EditorViewController;
  private pageSetup: PageSetup = DEFAULT_PAGE_SETUP;
  private allocatedWidth?: number;
  private allocatedHeight?: number;
  private deferredRenderHtml: string | null = null;
  private deferredRenderFrame: number | undefined;
  private resizeObserver: ResizeObserver | null = null;
  private imageHydrationRun = 0;

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
  private readonly handleSelectionChange = (): void => this.styleSelectionTracker.rememberTextSelection();
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
    this.paginator = new Paginator(
      (html?: string) => this.createPageElement(html),
      (pages: HTMLElement[]) => this.onPagesChanged(pages),
      (page: HTMLElement, afterPage: HTMLElement | null) =>
        this.attachPageForMeasurement(page, afterPage)
    );
    await this.loadParagraphStyles();
    await this.loadContent();
  }

  private buildShell(): void {
    this.container.innerHTML = "";
    this.container.style.cssText =
      "width:100%;height:100%;min-height:0;overflow:hidden;display:flex;flex-direction:column;";
    this.applyAllocatedSize();

    this.root = document.createElement("div");
    this.root.className = "hwe-root";
    this.applyCurrentPageSetup();
    installHweDebugGlobals(() => this.root ?? null);
    hweDebugLog("editor.buildShell", {
      allocatedHeight: this.allocatedHeight,
      allocatedWidth: this.allocatedWidth,
    });

    this.editorHeader = document.createElement("div");
    this.editorHeader.className = "hwe-editor-header";
    this.root.appendChild(this.editorHeader);
    this.tableCommandController = new TableCommandController({
      rootProvider: () => this.root,
      getActiveEditable: () => this.getActiveEditable(),
      getEditableForPageIndex: (pageIndex) => this.getEditableForPageIndex(pageIndex),
      markEdited: (element) => this.markEditedAndRebalance(element),
    });
    this.viewController = new EditorViewController(this.editorHeader, (view) => {
      void this.switchView(view);
    });

    this.toolbar = new Toolbar({
      onInsertTable: () => this.tableCommandController.insertTable(),
      onInsertRowAfter: () => this.tableCommandController.insertTableRowAfter(),
      onApplyParagraphStyle: (className) => this.applyParagraphStyle(className),
      onExportPdf: () => {
        void this.exportPdf();
      },
    });
    const toolbarEl = this.toolbar.build();
    this.toolbar.getSaveButton().addEventListener("click", () => void this.save());
    this.editorHeader.appendChild(toolbarEl);
    this.paragraphStyleManager = new ParagraphStyleManager(() => this.root, this.toolbar);
    this.styleSelectionTracker = new StyleSelectionTracker(
      () => this.root,
      () => this.getActiveEditable()
    );
    this.diagnosticsController = new EditorDiagnosticsController(
      () => this.root ?? null,
      (message, type) => this.setStatus(message, type)
    );
    document.addEventListener("selectionchange", this.handleSelectionChange);

    this.viewController.build();

    this.workspace = document.createElement("div");
    this.workspace.className = "hwe-workspace";
    this.root.appendChild(this.workspace);
    this.imageResizeController = new ImageResizeController({
      rootProvider: () => this.root ?? null,
      onImageChanged: (image) => this.markImageEdited(image),
    });
    this.imageResizeController.start();

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

    statusBar.appendChild(this.diagnosticsController.makeButton());

    this.root.appendChild(statusBar);
    this.container.appendChild(this.root);
    this.diagnosticsController.start();

    if (typeof ResizeObserver !== "undefined") {
      this.resizeObserver = new ResizeObserver(() => this.renderDeferredWhenVisible());
      this.resizeObserver.observe(this.container);
    }
  }

  private async switchView(view: EditorView): Promise<void> {
    if (view === this.activeView) return;

    if (view === "source") {
      this.imageResizeController.clearSelection();
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
    this.viewController.update(this.activeView, this.workspace, this.sourceEditor);
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

      this.paragraphStyleManager.setStyles(styles.length > 0 ? styles : LOCAL_PARAGRAPH_STYLES);
    } catch (error) {
      console.warn("[HtmlWordEditor] paragraph styles fallback:", error);
      this.paragraphStyleManager.setStyles(LOCAL_PARAGRAPH_STYLES);
    }
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
    const done = hweDebugStart("editor.renderAndPaginate", {
      htmlLength: html.length,
    });
    this.imageHydrationRun++;
    this.workspace.innerHTML = "";

    const normalizedDocument = this.documentSerializer.normalizeHtmlForPagination(html);
    this.setPageSetup(normalizedDocument.pageSetup ?? DEFAULT_PAGE_SETUP);
    hweDebugLog("editor.renderAndPaginate.normalized", {
      htmlLength: normalizedDocument.html.length,
      pageSetup: normalizedDocument.pageSetup ?? null,
    });

    this.pages = [this.createPageElement(normalizedDocument.html)];
    this.workspace.appendChild(this.pages[0]);
    this.layoutService.applyOfficialTableWidths(this.workspace);

    this.paginator.setPages(this.pages);
    if (!this.canMeasureLayout()) {
      this.deferredRenderHtml = html;
      this.syncWorkspace();
      this.updatePageCount();
      hweDebugLog("editor.renderAndPaginate.deferred", {
        root: this.root.getBoundingClientRect(),
        workspace: this.workspace.getBoundingClientRect(),
      });
      done({
        deferred: true,
        pages: this.pages.length,
      });
      return;
    }

    this.deferredRenderHtml = null;
    await this.assetLayoutManager.waitForStableLayout(this.workspace);

    this.paginator.repaginateAll();
    this.pages = this.paginator.getPages();
    this.pages.forEach((page) => this.layoutService.applyOfficialTableWidths(page));
    this.syncWorkspace();
    this.updatePageCount();
    this.imageResizeController.refresh();
    this.startDetachedImageHydration();
    done({
      pages: this.pages.length,
    });
  }

  private createPageElement(html?: string): HTMLElement {
    const page = document.createElement("div");
    page.className = "hwe-page";

    const inner = document.createElement("div");
    inner.className = "hwe-page-inner";
    inner.setAttribute("contenteditable", "true");
    inner.setAttribute("spellcheck", "false");
    inner.innerHTML = html ?? "<p><br></p>";
    this.layoutService.applyOfficialTableWidths(inner);

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
        const isEnterInput = this.isEnterInput(inputType);
        const shouldPullFromNextPages =
          this.isDeleteInput(inputType) || this.pagesNeedingPull.has(page);
        this.blankLineController.syncEditableBlankBlocks(inner, isEnterInput);
        this.pagesNeedingPull.delete(page);
        this.scheduleRebalance(page, shouldPullFromNextPages, {
          includePreviousPage: shouldPullFromNextPages,
          compactPages: shouldPullFromNextPages || !isEnterInput,
          overflowOnly: isEnterInput && !shouldPullFromNextPages,
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
      this.styleSelectionTracker.rememberTextSelection();
      this.tableCommandController.rememberSelectedTableRow();
    });
    inner.addEventListener("click", (event) => {
      this.tableCommandController.rememberTableRowFromEvent(event);
      this.styleSelectionTracker.rememberStyleBlockFromEvent(event);
    });
    inner.addEventListener("mouseup", (event) => {
      this.styleSelectionTracker.rememberStyleBlockFromEvent(event);
      this.styleSelectionTracker.rememberTextSelection();
      this.tableCommandController.rememberSelectedTableRow();
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
      const done = hweDebugStart("editor.scheduleRebalance.flush", {
        compactPages: pending.compactPages,
        includePreviousPage: pending.includePreviousPage,
        overflowOnly: pending.overflowOnly,
        pageIndex,
        pages: this.pages.length,
        pullFromNextPages: pending.pullFromNextPages,
      });

      const startPageIndex = pending.includePreviousPage ? Math.max(0, pageIndex - 1) : pageIndex;
      const startPage = this.pages[startPageIndex] ?? pending.page;
      const activeEditable = this.getActiveEditable();
      const marker = CaretManager.createMarker(this.root);
      const caretViewportTop = marker?.getBoundingClientRect().top ?? null;
      if (pending.overflowOnly) {
        const resolved = this.paginator.pushOverflowForwardFromPage(pending.page);
        this.pages = this.paginator.getPages();
        const fallbackPageIndex = Math.max(0, Math.min(pageIndex, this.pages.length - 1));
        const fallbackPage = this.pages[fallbackPageIndex];
        if (!resolved && fallbackPage && this.layoutService.pageOverflows(fallbackPage)) {
          hweDebugLog("editor.scheduleRebalance.overflowFallback", {
            fallbackPageIndex,
            pages: this.pages.length,
          });
          window.requestAnimationFrame(() => {
            if (!this.pages.includes(fallbackPage) || !this.layoutService.pageOverflows(fallbackPage)) return;
            this.scheduleRebalance(fallbackPage, false, {
              compactPages: false,
              includePreviousPage: false,
            });
          });
        }
      } else {
        this.paginator.rebalanceFromPage(startPage, {
          includePreviousPage: false,
          compactPages: pending.compactPages,
        });
      }
      this.pages = this.paginator.getPages();
      this.pages.forEach((currentPage) => {
        const inner = currentPage.querySelector<HTMLElement>(".hwe-page-inner");
        if (inner) this.blankLineController.syncEditableBlankBlocks(inner, false);
      });
      this.syncWorkspace();
      this.updatePageCount();
      this.imageResizeController.refresh();
      const fallbackEditable = this.getEditableForPageIndex(pageIndex) ?? activeEditable;
      this.restoreCaretViewport(marker, caretViewportTop);
      CaretManager.restoreMarker(marker, fallbackEditable);
      CaretManager.removeMarkers(this.root);
      done({
        pages: this.pages.length,
      });
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
      overflowOnly: options.overflowOnly ?? false,
    };

    if (!this.pendingRebalance) {
      this.pendingRebalance = nextRebalance;
      return;
    }

    const currentIndex = this.pages.indexOf(this.pendingRebalance.page);
    const nextIndex = this.pages.indexOf(page);
    const shouldUseNextPage =
      currentIndex === -1 || (nextIndex !== -1 && nextIndex < currentIndex);

    const mergedPullFromNextPages =
      this.pendingRebalance.pullFromNextPages || nextRebalance.pullFromNextPages;

    this.pendingRebalance = {
      page: shouldUseNextPage ? page : this.pendingRebalance.page,
      pullFromNextPages: mergedPullFromNextPages,
      includePreviousPage:
        this.pendingRebalance.includePreviousPage || nextRebalance.includePreviousPage,
      compactPages: mergedPullFromNextPages
        ? true
        : this.pendingRebalance.compactPages && nextRebalance.compactPages,
      overflowOnly:
        !mergedPullFromNextPages &&
        this.pendingRebalance.overflowOnly &&
        nextRebalance.overflowOnly,
    };
  }

  private onPagesChanged(pages: HTMLElement[]): void {
    this.pages = pages;
    this.pages.forEach((page) => this.layoutService.applyOfficialTableWidths(page));
    this.syncWorkspace();
    this.updatePageCount();
    this.imageResizeController?.refresh();
  }

  private attachPageForMeasurement(page: HTMLElement, afterPage: HTMLElement | null): void {
    this.layoutService.applyOfficialTableWidths(page);
    if (page.parentElement === this.workspace) return;

    const reference =
      afterPage?.parentElement === this.workspace ? afterPage.nextSibling : null;
    this.workspace.insertBefore(page, reference);
  }

  private startDetachedImageHydration(): void {
    const runId = ++this.imageHydrationRun;
    void this.assetLayoutManager.hydrateDetachedImages(
      this.workspace,
      (id) => this.documentSerializer.getDetachedLargeImageSrc(id),
      async (image) => {
        if (runId !== this.imageHydrationRun) return;

        const page = image.closest<HTMLElement>(".hwe-page");
        if (!page || !this.pages.includes(page)) return;

        this.layoutService.applyOfficialTableWidths(page);
        await this.assetLayoutManager.waitForStableLayout(page);
        if (runId !== this.imageHydrationRun || !this.pages.includes(page)) return;

        this.scheduleRebalance(page, false, {
          compactPages: false,
          includePreviousPage: false,
        });
      }
    );
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
    this.imageResizeController?.refresh();
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

  private onPaste(event: ClipboardEvent, page: HTMLElement): void {
    const result = this.pasteController.handlePaste(event, page);
    if (!result.handled) return;

    const affectedPage = result.affectedPage ?? page;
    const inner = affectedPage.querySelector<HTMLElement>(".hwe-page-inner");
    if (inner) this.blankLineController.syncEditableBlankBlocks(inner, false);
    this.layoutService.applyOfficialTableWidths(affectedPage);

    this.isDirty = true;
    this.toolbar.updateActiveStates();
    this.scheduleRebalance(affectedPage, false, {
      compactPages: false,
      includePreviousPage: false,
    });
    void this.assetLayoutManager
      .waitForStableLayout(affectedPage)
      .then(() => {
        this.scheduleRebalance(affectedPage, false, {
          compactPages: false,
          includePreviousPage: false,
        });
      });
  }

  private applyParagraphStyle(className: string): void {
    const shouldClearStyle = className === CLEAR_PARAGRAPH_STYLE_VALUE;
    if (!shouldClearStyle && !this.paragraphStyleManager.hasClass(className)) return;
    if (this.activeView !== "visual") {
      this.setStatus("Vuelve al editor visual para aplicar estilos.", "error");
      return;
    }

    this.styleSelectionTracker.restoreTextSelection();
    const blocks = this.styleSelectionTracker.getSelectedStyleBlocks();
    if (blocks.length === 0) {
      this.setStatus("Selecciona un parrafo para aplicar el estilo.", "error");
      return;
    }

    blocks.forEach((block) => {
      this.paragraphStyleManager.applyToBlock(block, shouldClearStyle ? null : className);
    });

    this.styleSelectionTracker.rememberTextSelection();
    this.markEditedAfterStyleChange(blocks[0]);
  }

  private markEditedAfterStyleChange(element: HTMLElement): void {
    const page = element.closest<HTMLElement>(".hwe-page");
    if (!page) return;

    this.isDirty = true;
    this.toolbar.updateActiveStates();

    if (this.layoutService.pageOverflows(page)) {
      this.scheduleRebalance(page, false, {
        includePreviousPage: false,
        compactPages: false,
      });
    }
  }

  private onPageKeyDown(event: KeyboardEvent): void {
    if (event.ctrlKey && event.key.toLowerCase() === "s") {
      event.preventDefault();
      void this.save();
    }

    if (
      event.key === "Backspace" &&
      this.pageBackspaceController.handleBackspaceAtPageStart(event, {
        pages: this.pages,
        onContentChanged: (previousPage) => {
          this.isDirty = true;
          this.toolbar.updateActiveStates();
          this.scheduleRebalance(previousPage, true, { includePreviousPage: true });
        },
      })
    ) {
      return;
    }

    if (event.key === "Backspace" || event.key === "Delete") {
      const page = (event.currentTarget as HTMLElement).closest<HTMLElement>(".hwe-page");
      if (page) this.pagesNeedingPull.add(page);
    }
  }

  private markEditedAndRebalance(element: HTMLElement): void {
    const page = element.closest<HTMLElement>(".hwe-page");
    if (!page) return;
    this.isDirty = true;
    this.toolbar.updateActiveStates();
    this.scheduleRebalance(page, true, { includePreviousPage: false });
  }

  private markImageEdited(image: HTMLImageElement): void {
    const page = image.closest<HTMLElement>(".hwe-page");
    if (!page) return;

    this.isDirty = true;
    this.toolbar.updateActiveStates();
    this.layoutService.applyOfficialTableWidths(page);
    this.scheduleRebalance(page, false, {
      compactPages: false,
      includePreviousPage: false,
    });
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

    const html = this.documentSerializer.collectPdfHtml(
      this.root,
      this.pages,
      this.pageSetup,
      this.paragraphStyleManager.cssText
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
    if (this.layoutService.pageOverflows(page)) return true;
    return pullFromNextPages && pageIndex >= 0 && pageIndex < this.pages.length - 1;
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

  resize(width?: number, height?: number): void {
    this.allocatedWidth = width;
    this.allocatedHeight = height;
    this.applyAllocatedSize();
    this.renderDeferredWhenVisible();
  }

  getHtml(): string {
    if (this.activeView === "source" && this.sourceDirty) {
      return this.sourceEditor.value;
    }

    return this.collectHtml();
  }

  destroy(): void {
    this.imageHydrationRun++;
    this.diagnosticsController?.destroy();
    if (this.rebalanceFrame !== undefined) {
      window.cancelAnimationFrame(this.rebalanceFrame);
    }
    if (this.deferredRenderFrame !== undefined) {
      window.cancelAnimationFrame(this.deferredRenderFrame);
    }
    this.resizeObserver?.disconnect();
    document.removeEventListener("selectionchange", this.handleSelectionChange);
    this.imageResizeController?.destroy();
    this.toolbar?.destroy();
    this.paginator?.destroy();
    this.container.innerHTML = "";
  }

  private applyAllocatedSize(): void {
    const width = this.formatAllocatedSize(this.allocatedWidth);
    const height = this.formatAllocatedSize(this.allocatedHeight);

    this.container.style.width = width;
    this.container.style.height = height;
    this.container.style.minHeight = "0";
    this.container.style.overflow = "hidden";
    this.container.style.display = "flex";
    this.container.style.flexDirection = "column";
  }

  private renderDeferredWhenVisible(): void {
    if (!this.deferredRenderHtml || this.deferredRenderFrame !== undefined) return;

    this.deferredRenderFrame = window.requestAnimationFrame(() => {
      this.deferredRenderFrame = undefined;
      if (!this.deferredRenderHtml || !this.canMeasureLayout()) return;

      const html = this.deferredRenderHtml;
      this.deferredRenderHtml = null;
      hweDebugLog("editor.renderDeferredWhenVisible", {
        htmlLength: html.length,
      });
      void this.renderAndPaginate(html);
    });
  }

  private canMeasureLayout(): boolean {
    if (!this.root?.isConnected || !this.workspace?.isConnected) return false;

    const rootRect = this.root.getBoundingClientRect();
    const workspaceRect = this.workspace.getBoundingClientRect();
    return rootRect.width > 20 && rootRect.height > 20 && workspaceRect.height > 20;
  }

  private formatAllocatedSize(value: number | undefined): string {
    return typeof value === "number" && Number.isFinite(value) && value > 0
      ? `${value}px`
      : "100%";
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
