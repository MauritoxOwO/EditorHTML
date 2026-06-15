import { Paginator, RebalanceOptions } from "../pagination/Paginator";
import { CaretManager } from "../pagination/CaretManager";
import { MANUAL_PAGE_BREAK_ATTR } from "../pagination/PaginatorDom";
import {
  CLEAR_PARAGRAPH_STYLE_VALUE,
  Toolbar,
} from "../ui/Toolbar";
import { fetchHtmlFromFileField, saveHtmlToFileField } from "../services/dataverse/fileApi";
import {
  fetchParagraphStyleCatalog,
  ParagraphFontFaceDefinition,
  ParagraphStyleCatalog,
  ParagraphStyleDefinition,
  ParagraphStyleTableConfig,
} from "../services/dataverse/styleApi";
import { fetchDynamicDocumentHeaderHtml } from "../services/dataverse/dynamicHeaderApi";
import {
  fetchRuntimeHeaderLogoSrc,
  RuntimeHeaderLogoTableConfig,
} from "../services/dataverse/runtimeHeaderLogoApi";
import {
  applyPageSetup,
  DEFAULT_PAGE_SETUP,
  normalizePageSetup,
  PageSetup,
} from "../pagination/PageGeometry";
import { BlankLineController } from "../controllers/BlankLineController";
import { DocumentSerializer } from "../services/DocumentSerializer";
import { PasteController } from "../controllers/PasteController";
import { AssetLayoutManager } from "../services/AssetLayoutManager";
import { EditorDiagnosticsController } from "../controllers/EditorDiagnosticsController";
import { EditorLayoutService } from "../services/EditorLayoutService";
import { PrintHtmlExportService } from "../services/PrintHtmlExportService";
import { EditorHistoryController } from "../controllers/EditorHistoryController";
import { ImageResizeController } from "../controllers/ImageResizeController";
import { PageBackspaceController } from "../controllers/PageBackspaceController";
import { ParagraphStyleManager } from "../services/ParagraphStyleManager";
import { StyleSelectionTracker } from "../controllers/StyleSelectionTracker";
import { TableDomIntegrityController } from "../controllers/TableDomIntegrityController";
import { TableColumnResizeController } from "../controllers/TableColumnResizeController";
import { TableCommandController } from "../controllers/TableCommandController";
import { EditorView, EditorViewController } from "../ui/EditorViewController";
import { RuntimePageHeaderRenderer } from "../ui/RuntimePageHeaderRenderer";
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
  stateField: "statecode",
  dropdownField: "mostrarendesplegable",
  documentTypeField: "tipodoc",
  documentTypeDropdownValue: "Anuncio",
};
const DEFAULT_MODEL_DRIVEN_EDITOR_HEIGHT_PX = 900;
const API_HEADER_ATTR = "data-hwe-api-header";
const API_HEADER_SOURCE_ATTR = "data-hwe-api-header-source";
const API_HEADER_SOURCE = "ays_GenerarCabeceraAnuncio";
const API_HEADER_SELECTOR = `[${API_HEADER_ATTR}='true']`;
const LEGACY_DYNAMIC_HEADER_SELECTOR = "[data-hwe-dynamic-header='true']";
const MANAGED_HEADER_SELECTOR = `${API_HEADER_SELECTOR}, ${LEGACY_DYNAMIC_HEADER_SELECTOR}`;
const DEFAULT_RUNTIME_HEADER_LOGO_NAME_VALUE = "logo-bocm.jpg";
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
  font-family: "Swiss721 BT", "Swis721 BT", "SwissRoman", Helvetica, Arial, sans-serif;
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
  savePrintHtml?: (html: string) => Promise<void> | void;
  paragraphStyles?: ParagraphStyleDefinition[];
  paragraphFonts?: ParagraphFontFaceDefinition[];
  paragraphStyleCatalog?: ParagraphStyleCatalog;
  runtimeHeaderLogoConfig?: RuntimeHeaderLogoTableConfig;
  runtimeHeaderLogoSrc?: string;
}

export class EditorComponent {
  private readonly container: HTMLElement;
  private readonly options: EditorComponentOptions;
  private readonly isPcfHost: boolean;

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
  private readonly printHtmlExportService = new PrintHtmlExportService(this.documentSerializer);
  private readonly assetLayoutManager = new AssetLayoutManager();
  private readonly pasteController = new PasteController();
  private readonly layoutService = new EditorLayoutService();
  private readonly historyController = new EditorHistoryController(10);
  private readonly tableDomIntegrityController = new TableDomIntegrityController();
  private readonly runtimePageHeaderRenderer = new RuntimePageHeaderRenderer();
  private imageResizeController!: ImageResizeController;
  private readonly pageBackspaceController = new PageBackspaceController();
  private diagnosticsController!: EditorDiagnosticsController;
  private paragraphStyleManager!: ParagraphStyleManager;
  private styleSelectionTracker!: StyleSelectionTracker;
  private tableColumnResizeController!: TableColumnResizeController;
  private tableCommandController!: TableCommandController;
  private viewController!: EditorViewController;
  private pageSetup: PageSetup = DEFAULT_PAGE_SETUP;
  private allocatedWidth?: number;
  private allocatedHeight?: number;
  private deferredRenderHtml: string | null = null;
  private deferredRenderFrame: number | undefined;
  private historySnapshotTimer: number | undefined;
  private resizeObserver: ResizeObserver | null = null;
  private imageHydrationRun = 0;

  private readonly baseUrl: string;
  private readonly entityName: string;
  private readonly entityId: string;
  private readonly fieldName: string;
  private readonly printHtmlFieldName: string | undefined;
  private readonly styleTableConfig: ParagraphStyleTableConfig;
  private readonly runtimeHeaderLogoConfig: RuntimeHeaderLogoTableConfig;
  private runtimeHeaderLogoSrc = "";
  private dynamicHeaderHtml = "";
  private dynamicHeaderLoaded = false;
  private currentFileName = "content.html";

  private rebalanceFrame: number | undefined;
  private pendingRebalance: QueuedRebalance | null = null;
  private readonly pagesNeedingPull = new WeakSet<HTMLElement>();
  private readonly pendingInputTypes = new WeakMap<HTMLElement, string>();
  private readonly handleSelectionChange = (): void => this.styleSelectionTracker.rememberTextSelection();
  private isComposing = false;
  private isDirty = false;
  private isRestoringHistory = false;
  private activeView: EditorView = "visual";
  private sourceDirty = false;

  constructor(container: HTMLElement, context?: PcfContext, options: EditorComponentOptions = {}) {
    this.container = container;
    this.options = options;
    this.isPcfHost = context !== undefined;

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
    this.printHtmlFieldName = this.getParameterValue(runtime.parameters, "printHtmlFieldName");
    this.runtimeHeaderLogoConfig = {
      entitySetName:
        options.runtimeHeaderLogoConfig?.entitySetName ??
        this.getParameterValue(runtime.parameters, "runtimeHeaderLogoEntitySetName"),
      imageField:
        options.runtimeHeaderLogoConfig?.imageField ??
        this.getParameterValue(runtime.parameters, "runtimeHeaderLogoImageField"),
      recordId:
        options.runtimeHeaderLogoConfig?.recordId ??
        this.getParameterValue(runtime.parameters, "runtimeHeaderLogoRecordId"),
      idField:
        options.runtimeHeaderLogoConfig?.idField ??
        this.getParameterValue(runtime.parameters, "runtimeHeaderLogoIdField"),
      nameField:
        options.runtimeHeaderLogoConfig?.nameField ??
        this.getParameterValue(runtime.parameters, "runtimeHeaderLogoNameField"),
      nameValue:
        options.runtimeHeaderLogoConfig?.nameValue ??
        this.getParameterValue(runtime.parameters, "runtimeHeaderLogoNameValue") ??
        DEFAULT_RUNTIME_HEADER_LOGO_NAME_VALUE,
    };
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
      stateField:
        this.getParameterValue(runtime.parameters, "styleStateField") ??
        DEFAULT_STYLE_TABLE_CONFIG.stateField,
      dropdownField:
        this.getParameterValue(runtime.parameters, "styleDropdownField") ??
        DEFAULT_STYLE_TABLE_CONFIG.dropdownField,
      documentTypeField:
        this.getParameterValue(runtime.parameters, "styleDocumentTypeField") ??
        DEFAULT_STYLE_TABLE_CONFIG.documentTypeField,
      documentTypeDropdownValue:
        this.getParameterValue(runtime.parameters, "styleDocumentTypeDropdownValue") ??
        DEFAULT_STYLE_TABLE_CONFIG.documentTypeDropdownValue,
      typeField:
        this.getParameterValue(runtime.parameters, "styleTypeField") ??
        DEFAULT_STYLE_TABLE_CONFIG.typeField,
      styleTypeValue:
        this.getParameterValue(runtime.parameters, "styleTypeStyleValue") ??
        DEFAULT_STYLE_TABLE_CONFIG.styleTypeValue,
      fontTypeValue:
        this.getParameterValue(runtime.parameters, "styleTypeFontValue") ??
        DEFAULT_STYLE_TABLE_CONFIG.fontTypeValue,
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
    await Promise.all([
      this.loadParagraphStyles(),
      this.loadDynamicHeader(),
      this.loadRuntimeHeaderLogo(),
    ]);
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
    this.tableCommandController.start();
    this.tableColumnResizeController = new TableColumnResizeController({
      onColumnsChanged: (table) => this.markTableColumnsChanged(table),
      rootProvider: () => this.root ?? null,
    });
    this.viewController = new EditorViewController(this.editorHeader, (view) => {
      void this.switchView(view);
    });

    this.toolbar = new Toolbar({
      onInsertTable: () => this.tableCommandController.insertTable(),
      onInsertRowAfter: () => this.tableCommandController.insertTableRowAfter(),
      onDeleteRow: () => this.tableCommandController.deleteTableRow(),
      onInsertPageBreak: () => this.insertManualPageBreak(),
      onApplyParagraphStyle: (className) => this.applyParagraphStyle(className),
      onCommand: (command) => this.imageResizeController?.handleToolbarCommand(command) ?? false,
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
    this.tableColumnResizeController.start();

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
      this.tableColumnResizeController.clear();
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
    this.recordHistorySnapshotNow();
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
      const catalog = this.options.paragraphStyleCatalog
        ? this.options.paragraphStyleCatalog
        : this.options.paragraphStyles
          ? {
              styles: this.options.paragraphStyles,
              fonts: this.options.paragraphFonts ?? [],
            }
        : this.baseUrl
          ? await fetchParagraphStyleCatalog(this.baseUrl, this.styleTableConfig)
          : {
              styles: LOCAL_PARAGRAPH_STYLES,
              fonts: [],
            };

      this.paragraphStyleManager.setCatalog(
        catalog.styles.length > 0
          ? catalog
          : {
              styles: LOCAL_PARAGRAPH_STYLES,
              fonts: catalog.fonts,
            }
      );
    } catch (error) {
      console.warn("[HtmlWordEditor] paragraph styles fallback:", error);
      this.paragraphStyleManager.setStyles(LOCAL_PARAGRAPH_STYLES);
    }
  }

  private async loadDynamicHeader(): Promise<void> {
    if (!this.shouldLoadDynamicHeader()) {
      this.dynamicHeaderHtml = "";
      this.dynamicHeaderLoaded = false;
      return;
    }

    try {
      this.dynamicHeaderHtml = await this.fetchWrappedDynamicHeaderHtml();
      this.dynamicHeaderLoaded = true;
    } catch (error) {
      console.warn("[HtmlWordEditor] dynamic header fallback:", error);
      this.dynamicHeaderHtml = "";
      this.dynamicHeaderLoaded = false;
    }
  }

  private shouldLoadDynamicHeader(): boolean {
    return Boolean(this.baseUrl && this.entityId);
  }

  private wrapDynamicHeaderHtml(html: string): string {
    if (!html.trim()) return "";
    return `<div ${API_HEADER_ATTR}="true" ${API_HEADER_SOURCE_ATTR}="${API_HEADER_SOURCE}" contenteditable="false">${html}</div>`;
  }

  private async fetchWrappedDynamicHeaderHtml(): Promise<string> {
    const html = await fetchDynamicDocumentHeaderHtml(
      this.baseUrl,
      this.entityId
    );
    return this.wrapDynamicHeaderHtml(html);
  }

  private async loadRuntimeHeaderLogo(): Promise<void> {
    const staticLogoSrc = this.options.runtimeHeaderLogoSrc?.trim();
    if (staticLogoSrc) {
      this.setRuntimeHeaderLogoSrc(staticLogoSrc);
      return;
    }

    if (!this.baseUrl || !this.runtimeHeaderLogoConfig.entitySetName) {
      this.setRuntimeHeaderLogoSrc("");
      return;
    }

    try {
      const logoSrc = await fetchRuntimeHeaderLogoSrc(this.baseUrl, this.runtimeHeaderLogoConfig);
      this.setRuntimeHeaderLogoSrc(logoSrc);
    } catch (error) {
      console.warn("[HtmlWordEditor] runtime BOCM logo fallback:", error);
      this.setRuntimeHeaderLogoSrc("");
    }
  }

  private setRuntimeHeaderLogoSrc(src: string): void {
    if (this.runtimeHeaderLogoSrc === src) return;

    this.revokeRuntimeHeaderLogoSrc();
    this.runtimeHeaderLogoSrc = src;
    this.runtimePageHeaderRenderer.setLogoSrc(src);
    this.pages.forEach((page) => this.runtimePageHeaderRenderer.ensureHeader(page));
  }

  private revokeRuntimeHeaderLogoSrc(): void {
    if (this.runtimeHeaderLogoSrc.startsWith("blob:")) {
      URL.revokeObjectURL(this.runtimeHeaderLogoSrc);
    }
    this.runtimeHeaderLogoSrc = "";
  }

  private async loadContent(): Promise<void> {
    this.setStatus("Cargando contenido...", "saving");

    try {
      if (this.options.loadHtml || this.options.initialHtml !== undefined) {
        const html = this.options.loadHtml
          ? await this.options.loadHtml()
          : this.options.initialHtml ?? "<p><br></p>";

        await this.renderAndPaginate(html || "<p><br></p>");
        this.resetHistorySnapshot();
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
      this.resetHistorySnapshot();
      this.setStatus("", "");
    } catch (err) {
      this.setStatus(`Error al cargar: ${(err as Error).message}`, "error");
      await this.renderAndPaginate("<p><br></p>");
      this.resetHistorySnapshot();
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

    this.pages = [this.createPageElement(this.withDynamicHeaderHtml(normalizedDocument.html))];
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

  private withDynamicHeaderHtml(html: string): string {
    const container = document.createElement("div");
    container.innerHTML = html || "<p><br></p>";

    if (this.dynamicHeaderLoaded) {
      this.removeExistingDynamicHeaders(container);
    } else {
      this.promoteLegacyDynamicHeaders(container);
    }

    if (this.dynamicHeaderLoaded && this.dynamicHeaderHtml) {
      const header = document.createElement("div");
      header.innerHTML = this.dynamicHeaderHtml;
      Array.from(header.childNodes).forEach((node) => {
        container.insertBefore(node, container.firstChild);
      });
    }

    return container.innerHTML || "<p><br></p>";
  }

  private removeExistingDynamicHeaders(container: HTMLElement): void {
    container.querySelectorAll<HTMLElement>(MANAGED_HEADER_SELECTOR).forEach(
      (element) => element.remove()
    );

    const headerElements = this.getDynamicHeaderElements();
    if (headerElements.length === 0) return;

    let safety = 0;
    while (safety++ < 5 && this.removeLeadingHeaderCopy(container, headerElements)) {
      // Keep removing stale generated headers left by older saves.
    }
  }

  private getDynamicHeaderElements(): HTMLElement[] {
    if (!this.dynamicHeaderHtml) return [];

    const template = document.createElement("div");
    template.innerHTML = this.dynamicHeaderHtml;
    const wrapper = template.querySelector<HTMLElement>(MANAGED_HEADER_SELECTOR);
    const source = wrapper ?? template;

    return Array.from(source.children).filter(
      (child): child is HTMLElement => child instanceof HTMLElement
    );
  }

  private promoteLegacyDynamicHeaders(container: HTMLElement): void {
    container.querySelectorAll<HTMLElement>(LEGACY_DYNAMIC_HEADER_SELECTOR).forEach((element) => {
      element.removeAttribute("data-hwe-dynamic-header");
      element.setAttribute(API_HEADER_ATTR, "true");
      element.setAttribute(API_HEADER_SOURCE_ATTR, API_HEADER_SOURCE);
      element.setAttribute("contenteditable", "false");
    });
  }

  private removeLeadingHeaderCopy(
    container: HTMLElement,
    headerElements: HTMLElement[]
  ): boolean {
    const candidates = Array.from(container.children).filter(
      (child): child is HTMLElement =>
        child instanceof HTMLElement && child.tagName !== "STYLE"
    );
    if (candidates.length < headerElements.length) return false;

    const leadingCandidates = candidates.slice(0, headerElements.length);
    const matches = leadingCandidates.every((candidate, index) =>
      this.looksLikeSameHeaderElement(candidate, headerElements[index])
    );
    if (!matches) return false;

    leadingCandidates.forEach((candidate) => candidate.remove());
    return true;
  }

  private looksLikeSameHeaderElement(candidate: HTMLElement, expected: HTMLElement): boolean {
    if (candidate.tagName !== expected.tagName) return false;

    const candidateClass = this.normalizeClassName(candidate.className);
    const expectedClass = this.normalizeClassName(expected.className);
    if (candidateClass && candidateClass === expectedClass) return true;

    const candidateStyle = this.normalizeHeaderStyle(candidate.getAttribute("style"));
    const expectedStyle = this.normalizeHeaderStyle(expected.getAttribute("style"));
    if (candidateStyle && candidateStyle === expectedStyle) return true;

    const candidateText = (candidate.textContent ?? "").replace(/\s+/g, " ").trim();
    const expectedText = (expected.textContent ?? "").replace(/\s+/g, " ").trim();
    if (candidateText && candidateText === expectedText) return true;

    return (
      this.hasHeaderStructuralMarker(expected) &&
      this.getElementShape(candidate) === this.getElementShape(expected)
    );
  }

  private normalizeClassName(value: string): string {
    return value.split(/\s+/).filter(Boolean).sort().join(" ");
  }

  private normalizeHeaderStyle(value: string | null): string {
    return (value ?? "")
      .replace(/\s+/g, " ")
      .replace(/\s*([:;])\s*/g, "$1")
      .trim()
      .toLowerCase();
  }

  private getElementShape(element: HTMLElement): string {
    const children = Array.from(element.children).filter(
      (child): child is HTMLElement => child instanceof HTMLElement
    );
    const self = [
      element.tagName.toLowerCase(),
      this.normalizeClassName(element.className),
      this.normalizeHeaderStyle(element.getAttribute("style")),
    ].join("|");

    return children.length > 0
      ? `${self}>${children.map((child) => this.getElementShape(child)).join(",")}`
      : self;
  }

  private hasHeaderStructuralMarker(element: HTMLElement): boolean {
    if (this.normalizeClassName(element.className)) return true;
    if (this.normalizeHeaderStyle(element.getAttribute("style"))) return true;

    return Array.from(element.children).some(
      (child) => child instanceof HTMLElement && this.hasHeaderStructuralMarker(child)
    );
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
      if (this.tableDomIntegrityController.handleBeforeInput(event, inner)) {
        this.markTableDomIntegrityChanged(page, inner);
        return;
      }

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
        if (this.tableDomIntegrityController.normalize(inner)) {
          this.layoutService.applyOfficialTableWidths(page);
        }
        this.pagesNeedingPull.delete(page);
        this.scheduleRebalance(page, shouldPullFromNextPages, {
          includePreviousPage: shouldPullFromNextPages,
          compactPages: shouldPullFromNextPages || !isEnterInput,
          overflowOnly: isEnterInput && !shouldPullFromNextPages,
        });
        this.scheduleHistorySnapshot();
      }
    });
    inner.addEventListener("compositionstart", () => {
      this.isComposing = true;
    });
    inner.addEventListener("compositionend", () => {
      this.isComposing = false;
      this.blankLineController.syncEditableBlankBlocks(inner, false);
      this.scheduleRebalance(page, false, { includePreviousPage: false });
      this.scheduleHistorySnapshot();
    });
    inner.addEventListener("paste", (event: ClipboardEvent) => {
      if (this.tableDomIntegrityController.handlePaste(event, inner)) {
        this.markTableDomIntegrityChanged(page, inner);
        return;
      }

      this.onPaste(event, page);
    });
    inner.addEventListener("drop", (event: DragEvent) => {
      if (this.tableDomIntegrityController.handleDrop(event, inner)) {
        this.markTableDomIntegrityChanged(page, inner);
      }
    });
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

    this.runtimePageHeaderRenderer.ensureHeader(page);
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
      if (
        !pending.force &&
        !this.shouldRebalancePage(pending.page, pageIndex, pending.pullFromNextPages)
      ) {
        return;
      }
      const done = hweDebugStart("editor.scheduleRebalance.flush", {
        compactPages: pending.compactPages,
        force: pending.force,
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
      force: options.force ?? false,
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
      force: this.pendingRebalance.force || nextRebalance.force,
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
      this.runtimePageHeaderRenderer.ensureHeader(page);
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
    if (inner) {
      this.tableDomIntegrityController.normalize(inner);
      this.blankLineController.syncEditableBlankBlocks(inner, false);
    }
    this.layoutService.applyOfficialTableWidths(affectedPage);

    this.isDirty = true;
    this.toolbar.updateActiveStates();
    this.scheduleRebalance(affectedPage, false, {
      compactPages: false,
      includePreviousPage: false,
    });
    this.recordHistorySnapshotNow();
    void this.assetLayoutManager
      .waitForStableLayout(affectedPage)
      .then(() => {
        this.scheduleRebalance(affectedPage, false, {
          compactPages: false,
          includePreviousPage: false,
        });
      });
  }

  private markTableDomIntegrityChanged(page: HTMLElement, inner: HTMLElement): void {
    this.tableDomIntegrityController.normalize(inner);
    this.blankLineController.syncEditableBlankBlocks(inner, false);
    this.layoutService.applyOfficialTableWidths(page);

    this.isDirty = true;
    this.toolbar.updateActiveStates();
    this.scheduleRebalance(page, false, {
      compactPages: false,
      includePreviousPage: false,
    });
    this.recordHistorySnapshotNow();
  }

  private insertManualPageBreak(): void {
    if (this.activeView !== "visual") {
      this.setStatus("Vuelve al editor visual para insertar un salto.", "error");
      return;
    }

    const editable = this.getActiveEditable();
    const page = editable?.closest<HTMLElement>(".hwe-page") ?? null;
    if (!editable || !page) {
      this.setStatus("Coloca el cursor donde quieres insertar el salto.", "error");
      return;
    }

    const marker = this.createManualPageBreakMarker();
    const caretTarget = this.insertManualPageBreakMarker(marker, editable);
    this.tableDomIntegrityController.normalize(editable);
    this.blankLineController.syncEditableBlankBlocks(editable, false);
    this.layoutService.applyOfficialTableWidths(page);
    this.placeCaretAtStart(caretTarget ?? marker.nextSibling, editable);

    this.isDirty = true;
    this.toolbar.updateActiveStates();
    this.scheduleRebalance(page, false, {
      compactPages: true,
      force: true,
      includePreviousPage: false,
    });
    this.recordHistorySnapshotNow();
  }

  private createManualPageBreakMarker(): HTMLElement {
    const marker = document.createElement("div");
    marker.className = "hwe-manual-page-break";
    marker.setAttribute(MANUAL_PAGE_BREAK_ATTR, "true");
    marker.setAttribute("aria-hidden", "true");
    return marker;
  }

  private insertManualPageBreakMarker(
    marker: HTMLElement,
    editable: HTMLElement
  ): ChildNode | null {
    const selection = window.getSelection();
    const range =
      selection && selection.rangeCount > 0 ? selection.getRangeAt(0) : null;

    if (!range || !editable.contains(range.commonAncestorContainer)) {
      editable.appendChild(marker);
      return this.ensureEditableBlockAfter(marker);
    }

    const startElement = this.nodeToElement(range.startContainer);
    const table = startElement?.closest<HTMLTableElement>("table");
    if (table && editable.contains(table)) {
      const flowRoot = table.closest<HTMLElement>(".hwe-table-flow-wrapper") ?? table;
      flowRoot.parentNode?.insertBefore(marker, flowRoot.nextSibling);
      return this.ensureEditableBlockAfter(marker);
    }

    const textBlock = startElement?.closest<HTMLElement>(
      "p, h1, h2, h3, h4, h5, h6, blockquote, pre"
    );
    if (textBlock && editable.contains(textBlock) && !textBlock.closest("td, th")) {
      return this.splitTextBlockWithManualBreak(textBlock, range, marker);
    }

    if (range.startContainer === editable) {
      const reference = editable.childNodes[range.startOffset] ?? null;
      editable.insertBefore(marker, reference);
      return this.ensureEditableBlockAfter(marker);
    }

    const flowChild = this.getTopLevelFlowChild(range.startContainer, editable);
    if (flowChild?.parentNode) {
      flowChild.parentNode.insertBefore(marker, flowChild.nextSibling);
      return this.ensureEditableBlockAfter(marker);
    }

    editable.appendChild(marker);
    return this.ensureEditableBlockAfter(marker);
  }

  private splitTextBlockWithManualBreak(
    block: HTMLElement,
    sourceRange: Range,
    marker: HTMLElement
  ): ChildNode | null {
    const parent = block.parentNode;
    if (!parent) return null;

    const range = sourceRange.cloneRange();
    if (!range.collapsed) range.deleteContents();

    const afterRange = range.cloneRange();
    afterRange.setEnd(block, block.childNodes.length);
    const afterContent = afterRange.extractContents();

    parent.insertBefore(marker, block.nextSibling);
    if (this.isVisuallyEmptyBlock(block)) block.innerHTML = "<br>";

    if (this.fragmentHasContent(afterContent)) {
      const afterBlock = block.cloneNode(false) as HTMLElement;
      afterBlock.appendChild(afterContent);
      parent.insertBefore(afterBlock, marker.nextSibling);
      return afterBlock;
    }

    return this.ensureEditableBlockAfter(marker);
  }

  private ensureEditableBlockAfter(marker: HTMLElement): ChildNode | null {
    let next = marker.nextSibling;
    while (next && this.isWhitespaceTextNode(next)) next = next.nextSibling;
    if (next) return next;

    const blank = document.createElement("p");
    blank.setAttribute("data-hwe-user-blank", "true");
    blank.appendChild(document.createElement("br"));
    marker.parentNode?.insertBefore(blank, marker.nextSibling);
    return blank;
  }

  private placeCaretAtStart(node: ChildNode | null, fallbackEditable: HTMLElement): void {
    const selection = window.getSelection();
    if (!selection) return;

    fallbackEditable.focus({ preventScroll: true });
    const range = document.createRange();
    if (!node) {
      range.selectNodeContents(fallbackEditable);
      range.collapse(false);
    } else if (node.nodeType === Node.TEXT_NODE) {
      range.setStart(node, 0);
      range.collapse(true);
    } else {
      range.selectNodeContents(node);
      range.collapse(true);
    }

    selection.removeAllRanges();
    selection.addRange(range);
  }

  private getTopLevelFlowChild(node: Node, editable: HTMLElement): ChildNode | null {
    let current: Node | null =
      node.nodeType === Node.ELEMENT_NODE ? node : node.parentNode;

    while (current?.parentNode && current.parentNode !== editable) {
      current = current.parentNode;
    }

    return current?.parentNode === editable ? (current as ChildNode) : null;
  }

  private nodeToElement(node: Node): HTMLElement | null {
    return node.nodeType === Node.ELEMENT_NODE
      ? (node as HTMLElement)
      : node.parentElement;
  }

  private fragmentHasContent(fragment: DocumentFragment): boolean {
    return Array.from(fragment.childNodes).some((node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        return (node.textContent ?? "").replace(/\u00a0/g, " ").trim() !== "";
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return false;
      const element = node as HTMLElement;
      if (element.tagName === "BR") return false;
      return true;
    });
  }

  private isVisuallyEmptyBlock(block: HTMLElement): boolean {
    const text = (block.textContent ?? "").replace(/\u00a0/g, " ").trim();
    if (text) return false;
    return !block.querySelector("img, table, tr, td, th, video, canvas, svg");
  }

  private isWhitespaceTextNode(node: ChildNode): boolean {
    return (
      node.nodeType === Node.TEXT_NODE &&
      (node.textContent ?? "").replace(/\u00a0/g, " ").trim() === ""
    );
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
    this.markEditedAfterStyleChange(blocks);
    this.recordHistorySnapshotNow();
  }

  private markEditedAfterStyleChange(blocks: HTMLElement[]): void {
    const affectedPageSet = new Set(
      blocks
        .map((block) => block.closest<HTMLElement>(".hwe-page"))
        .filter((page): page is HTMLElement => Boolean(page))
    );
    const affectedPages = this.pages.filter((page) => affectedPageSet.has(page));
    if (affectedPages.length === 0) return;

    this.isDirty = true;
    this.toolbar.updateActiveStates();

    const firstOverflowPage = affectedPages.find((page) => this.layoutService.pageOverflows(page));
    if (firstOverflowPage) {
      this.scheduleRebalance(firstOverflowPage, false, {
        includePreviousPage: false,
        compactPages: false,
      });
    }
  }

  private onPageKeyDown(event: KeyboardEvent): void {
    if (this.handleHistoryShortcut(event)) return;

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
          this.recordHistorySnapshotNow();
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
    this.recordHistorySnapshotNow();
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
    this.recordHistorySnapshotNow();
  }

  private markTableColumnsChanged(table: HTMLTableElement): void {
    const page = this.getFirstTableFlowPage(table) ?? table.closest<HTMLElement>(".hwe-page");
    if (!page) return;

    this.isDirty = true;
    this.toolbar.updateActiveStates();
    this.layoutService.applyOfficialTableWidths(page);
    this.scheduleRebalance(page, true, {
      compactPages: true,
      includePreviousPage: false,
    });
    this.recordHistorySnapshotNow();
  }

  private getFirstTableFlowPage(table: HTMLTableElement): HTMLElement | null {
    const flowId = table.getAttribute("data-hwe-table-flow-id");
    if (!flowId) return table.closest<HTMLElement>(".hwe-page");

    return (
      this.pages.find((page) =>
        Array.from(page.querySelectorAll<HTMLTableElement>(".hwe-page-inner table")).some(
          (candidate) => candidate.getAttribute("data-hwe-table-flow-id") === flowId
        )
      ) ?? null
    );
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
    const wasSourceView = this.activeView === "source";

    try {
      if (this.activeView === "source" && this.sourceDirty) {
        this.activeView = "visual";
        this.updateViewTabs();
        await this.renderAndPaginate(this.sourceEditor.value || "<p><br></p>");
        this.sourceDirty = false;
      }

      await this.refreshDynamicHeaderBeforeSave();

      const html = this.collectHtml();
      const printHtml = this.collectPrintHtml();
      if (this.options.saveHtml) {
        await this.options.saveHtml(html);
        await this.savePrintHtml(printHtml);
      } else {
        await saveHtmlToFileField(
          this.baseUrl,
          this.entityName,
          this.entityId,
          this.fieldName,
          html,
          this.currentFileName
        );
        await this.savePrintHtml(printHtml);
      }

      if (wasSourceView) {
        this.sourceEditor.value = html;
        this.sourceDirty = false;
        this.activeView = "source";
        this.updateViewTabs();
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

  private async refreshDynamicHeaderBeforeSave(): Promise<void> {
    if (!this.shouldLoadDynamicHeader()) return;

    const nextHeaderHtml = await this.fetchWrappedDynamicHeaderHtml();
    const alreadyCurrent =
      this.dynamicHeaderLoaded &&
      this.dynamicHeaderHtml === nextHeaderHtml &&
      this.pages.some((page) => page.querySelector(API_HEADER_SELECTOR));
    if (alreadyCurrent) return;

    const currentHtml = this.collectHtml();
    this.dynamicHeaderHtml = nextHeaderHtml;
    this.dynamicHeaderLoaded = true;
    const wasSourceView = this.activeView === "source";
    if (wasSourceView) {
      this.activeView = "visual";
      this.updateViewTabs();
    }
    await this.renderAndPaginate(currentHtml || "<p><br></p>");

    if (wasSourceView) {
      this.sourceEditor.value = this.collectHtml();
      this.sourceDirty = false;
      this.activeView = "source";
      this.updateViewTabs();
    }
  }

  private async savePrintHtml(printHtml: string): Promise<void> {
    if (this.options.savePrintHtml) {
      await this.options.savePrintHtml(printHtml);
      return;
    }

    if (!this.printHtmlFieldName) return;

    await saveHtmlToFileField(
      this.baseUrl,
      this.entityName,
      this.entityId,
      this.printHtmlFieldName,
      printHtml,
      this.getPrintHtmlFileName()
    );
  }

  private getPrintHtmlFileName(): string {
    const baseName = this.currentFileName.replace(/\.(?:html?|xhtml)$/i, "") || "content";
    return `${baseName}.print.html`;
  }

  private async exportPdf(): Promise<void> {
    if (this.activeView === "source" && this.sourceDirty) {
      this.setStatus("Vuelve al editor visual o guarda los cambios del HTML antes de exportar.", "error");
      return;
    }

    const html = this.collectPrintHtml();
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

  private handleHistoryShortcut(event: KeyboardEvent): boolean {
    const isModifierPressed = event.ctrlKey || event.metaKey;
    if (!isModifierPressed || event.altKey || this.activeView !== "visual") return false;

    const key = event.key.toLowerCase();
    const isUndo = key === "z" && !event.shiftKey;
    const isRedo = key === "y" || (key === "z" && event.shiftKey);
    if (!isUndo && !isRedo) return false;

    event.preventDefault();
    event.stopPropagation();
    if (this.isRestoringHistory) return true;

    this.flushPendingHistorySnapshot();
    const snapshot = isUndo ? this.historyController.undo() : this.historyController.redo();
    if (snapshot) void this.restoreHistorySnapshot(snapshot);
    return true;
  }

  private resetHistorySnapshot(): void {
    this.clearHistorySnapshotTimer();
    this.historyController.reset(this.collectHistoryHtml());
  }

  private scheduleHistorySnapshot(): void {
    if (this.isRestoringHistory || this.activeView !== "visual") return;
    this.clearHistorySnapshotTimer();
    this.historySnapshotTimer = window.setTimeout(() => {
      this.historySnapshotTimer = undefined;
      this.recordHistorySnapshotNow();
    }, 650);
  }

  private flushPendingHistorySnapshot(): void {
    if (this.historySnapshotTimer === undefined) return;
    this.clearHistorySnapshotTimer();
    this.recordHistorySnapshotNow();
  }

  private recordHistorySnapshotNow(): void {
    if (this.isRestoringHistory || this.activeView !== "visual") return;
    this.clearHistorySnapshotTimer();
    this.historyController.record(this.collectHistoryHtml());
  }

  private async restoreHistorySnapshot(snapshot: string): Promise<void> {
    if (this.isRestoringHistory) return;

    this.clearHistorySnapshotTimer();
    this.isRestoringHistory = true;
    try {
      await this.historyController.runSuspended(() => this.renderAndPaginate(snapshot));
      this.isDirty = true;
      this.sourceDirty = false;
      this.toolbar.updateActiveStates();
    } finally {
      this.isRestoringHistory = false;
    }
  }

  private clearHistorySnapshotTimer(): void {
    if (this.historySnapshotTimer === undefined) return;
    window.clearTimeout(this.historySnapshotTimer);
    this.historySnapshotTimer = undefined;
  }

  private collectHistoryHtml(): string {
    return this.documentSerializer.collectHistoryHtml(
      this.root,
      this.pages,
      this.pageSetup
    );
  }

  private collectHtml(): string {
    return this.documentSerializer.collectHtml(
      this.root,
      this.pages,
      this.pageSetup,
      this.paragraphStyleManager.cssText
    );
  }

  private collectPrintHtml(): string {
    return this.printHtmlExportService.createPrintHtml({
      root: this.root,
      pages: this.pages,
      pageSetup: this.pageSetup,
      additionalCss: this.paragraphStyleManager.cssText,
    });
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
    this.resetHistorySnapshot();
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
    this.clearHistorySnapshotTimer();
    this.resizeObserver?.disconnect();
    document.removeEventListener("selectionchange", this.handleSelectionChange);
    this.imageResizeController?.destroy();
    this.tableColumnResizeController?.destroy();
    this.tableCommandController?.destroy();
    this.paragraphStyleManager?.destroy();
    this.toolbar?.destroy();
    this.paginator?.destroy();
    this.revokeRuntimeHeaderLogoSrc();
    this.container.innerHTML = "";
  }

  private applyAllocatedSize(): void {
    const width = this.formatHostWidth(this.allocatedWidth);
    const height = this.formatHostHeight(this.allocatedHeight);

    this.container.style.width = width;
    this.container.style.height = height;
    this.container.style.minHeight = this.formatHostMinHeight(this.allocatedHeight);
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

  private formatHostWidth(value: number | undefined): string {
    if (this.isPcfHost) return "100%";

    return typeof value === "number" && Number.isFinite(value) && value > 0
      ? `${value}px`
      : "100%";
  }

  private formatHostHeight(value: number | undefined): string {
    if (this.isPcfHost) {
      return this.hasAllocatedSize(value) ? "100%" : `${DEFAULT_MODEL_DRIVEN_EDITOR_HEIGHT_PX}px`;
    }

    if (this.hasAllocatedSize(value)) {
      return `${value}px`;
    }

    return "100%";
  }

  private formatHostMinHeight(value: number | undefined): string {
    if (this.isPcfHost) {
      return this.hasAllocatedSize(value) ? "0" : `${DEFAULT_MODEL_DRIVEN_EDITOR_HEIGHT_PX}px`;
    }

    return "0";
  }

  private hasAllocatedSize(value: number | undefined): value is number {
    return typeof value === "number" && Number.isFinite(value) && value > 0;
  }
}

interface IInputs {
  htmlContent: ComponentFramework.PropertyTypes.StringProperty;
  printHtmlFieldName: ComponentFramework.PropertyTypes.StringProperty;
  styleEntitySetName: ComponentFramework.PropertyTypes.StringProperty;
  styleClassField: ComponentFramework.PropertyTypes.StringProperty;
  styleCssField: ComponentFramework.PropertyTypes.StringProperty;
  styleStateField: ComponentFramework.PropertyTypes.StringProperty;
  styleDropdownField: ComponentFramework.PropertyTypes.StringProperty;
  styleDocumentTypeField: ComponentFramework.PropertyTypes.StringProperty;
  styleDocumentTypeDropdownValue: ComponentFramework.PropertyTypes.StringProperty;
  styleTypeField: ComponentFramework.PropertyTypes.StringProperty;
  styleTypeStyleValue: ComponentFramework.PropertyTypes.StringProperty;
  styleTypeFontValue: ComponentFramework.PropertyTypes.StringProperty;
  runtimeHeaderLogoEntitySetName: ComponentFramework.PropertyTypes.StringProperty;
  runtimeHeaderLogoImageField: ComponentFramework.PropertyTypes.StringProperty;
  runtimeHeaderLogoRecordId: ComponentFramework.PropertyTypes.StringProperty;
  runtimeHeaderLogoIdField: ComponentFramework.PropertyTypes.StringProperty;
  runtimeHeaderLogoNameField: ComponentFramework.PropertyTypes.StringProperty;
  runtimeHeaderLogoNameValue: ComponentFramework.PropertyTypes.StringProperty;
}
