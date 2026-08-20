import { getTableFlowFragments } from "../dom/TableFlow";
import { getCellsForColumn, getColumnIndexAtClientX } from "../dom/TableGrid";

export interface TableSelectionControllerOptions {
  rootProvider: () => HTMLElement | null;
  onTableChanged: (table: HTMLTableElement) => void;
}

const SUPPORTED_COMMANDS = new Set(["bold", "underline"]);
const HANDLE_HIDE_DELAY_MS = 250;

type TableSelectionMode = "table" | "row" | "column";

export class TableSelectionController {
  private handle: HTMLButtonElement | null = null;
  private hoveredTable: HTMLTableElement | null = null;
  private activeCell: HTMLTableCellElement | null = null;
  private activeColumnIndex: number | null = null;
  private isCellTargetLocked = false;
  private selectedTable: HTMLTableElement | null = null;
  private selectedRow: HTMLTableRowElement | null = null;
  private selectedColumnIndex: number | null = null;
  private selectionMode: TableSelectionMode | null = null;
  private selectionToolbar: HTMLDivElement | null = null;
  private selectionToolbarHint: HTMLSpanElement | null = null;
  private selectionToolbarButtons = new Map<TableSelectionMode, HTMLButtonElement>();
  private selectionOverlays: HTMLDivElement[] = [];
  private handleHideTimer: number | null = null;
  private isStarted = false;

  private readonly handleRootPointerMove = (event: PointerEvent): void =>
    this.onRootPointerMove(event);
  private readonly handleRootPointerLeave = (): void => {
    if (!this.selectedTable) this.scheduleHandleHide();
  };
  private readonly handleDocumentPointerDown = (event: PointerEvent): void =>
    this.onDocumentPointerDown(event);
  private readonly handleScrollOrResize = (): void => this.refresh();
  private readonly handleDocumentKeyDown = (event: KeyboardEvent): void => {
    if (event.key === "Escape") this.clearSelection();
  };

  constructor(private readonly options: TableSelectionControllerOptions) {}

  start(): void {
    if (this.isStarted) return;
    const root = this.options.rootProvider();
    if (!root) return;

    this.isStarted = true;
    root.addEventListener("pointermove", this.handleRootPointerMove, true);
    root.addEventListener("pointerleave", this.handleRootPointerLeave);
    root.addEventListener("scroll", this.handleScrollOrResize, true);
    document.addEventListener("pointerdown", this.handleDocumentPointerDown, true);
    document.addEventListener("keydown", this.handleDocumentKeyDown);
    window.addEventListener("resize", this.handleScrollOrResize);
  }

  destroy(): void {
    if (!this.isStarted) return;
    const root = this.options.rootProvider();
    root?.removeEventListener("pointermove", this.handleRootPointerMove, true);
    root?.removeEventListener("pointerleave", this.handleRootPointerLeave);
    root?.removeEventListener("scroll", this.handleScrollOrResize, true);
    document.removeEventListener("pointerdown", this.handleDocumentPointerDown, true);
    document.removeEventListener("keydown", this.handleDocumentKeyDown);
    window.removeEventListener("resize", this.handleScrollOrResize);
    this.cancelHandleHide();
    this.isStarted = false;
    this.clearSelection();
    this.handle?.remove();
    this.handle = null;
    this.selectionToolbar?.remove();
    this.selectionToolbar = null;
    this.selectionToolbarHint = null;
    this.selectionToolbarButtons.clear();
  }

  handleToolbarCommand(command: string): boolean {
    if (!SUPPORTED_COMMANDS.has(command)) return false;

    const targets = this.getSelectedFormattingTargets();
    if (targets.length === 0) return false;
    const textContainers = targets.flatMap((target) => this.getTextContainers(target));

    if (command === "bold") {
      const isActive =
        textContainers.length > 0
          ? textContainers.every((element) => this.isBold(element))
          : targets.every((target) => this.isBold(target));
      targets.forEach((target) => {
        target.style.fontWeight = isActive ? "normal" : "700";
      });
      textContainers.forEach((element) => {
        element.style.fontWeight = isActive ? "normal" : "700";
      });
    } else {
      const isActive =
        textContainers.length > 0
          ? textContainers.every((element) => this.isUnderlined(element))
          : targets.every((target) => this.isUnderlined(target));
      targets.forEach((target) => {
        target.style.textDecoration = isActive ? "none" : "underline";
      });
      targets.forEach((target) => {
        this.getTextContainers(target).forEach((element) =>
          this.setUnderline(element, isActive ? "none" : "underline", target)
        );
      });
    }

    this.notifySelectedTableChanged();
    return true;
  }

  applyFontSize(fontSize: string): boolean {
    if (!fontSize) return false;

    const targets = this.getSelectedFormattingTargets();
    if (targets.length === 0) return false;

    targets.forEach((target) => {
      target.style.fontSize = fontSize;
      target.setAttribute("data-hwe-user-font-size", "true");
      this.getTextContainers(target).forEach((element) => {
        element.style.fontSize = fontSize;
        element.setAttribute("data-hwe-user-font-size", "true");
      });
    });
    this.notifySelectedTableChanged();
    return true;
  }

  refresh(): void {
    const table = this.getVisibleHandleTable();
    if (!table) {
      this.hideHandle();
    } else {
      this.positionHandle(table);
    }

    this.refreshSelectionVisuals();
  }

  clearSelection(): void {
    this.cancelHandleHide();
    this.getSelectedTables().forEach((table) => table.classList.remove("hwe-table-selected"));
    this.selectedTable = null;
    this.selectedRow = null;
    this.selectedColumnIndex = null;
    this.selectionMode = null;
    this.hoveredTable = null;
    this.activeCell = null;
    this.activeColumnIndex = null;
    this.isCellTargetLocked = false;
    this.clearSelectionOverlays();
    this.hideSelectionToolbar();
    this.hideHandle();
  }

  private onRootPointerMove(event: PointerEvent): void {
    const target = event.target as Element | null;
    if (!target) return;
    if (this.isSelectionControl(target)) {
      this.cancelHandleHide();
      return;
    }

    const table = target.closest<HTMLTableElement>(".hwe-page-inner table");
    const root = this.options.rootProvider();
    if (!table || !root?.contains(table)) {
      const selectedTable = this.getSelectedTables()[0];
      if (selectedTable) {
        this.hoveredTable = null;
        this.cancelHandleHide();
        this.positionHandle(selectedTable);
      } else {
        this.scheduleHandleHide();
      }
      return;
    }

    this.cancelHandleHide();
    this.hoveredTable = table;
    const cell = target.closest<HTMLTableCellElement>("td, th");
    if (cell && table.contains(cell) && !this.isCellTargetLocked) {
      this.activeCell = cell;
      this.activeColumnIndex = getColumnIndexAtClientX(table, cell, event.clientX);
      this.updateSelectionToolbarState();
    }
    this.positionHandle(table);
  }

  private updateSelectedScopeFromActiveCell(table: HTMLTableElement): void {
    if (!this.activeCell) return;

    if (this.selectionMode === "row") {
      const row = this.activeCell.closest<HTMLTableRowElement>("tr");
      if (!row || (this.selectedTable === table && this.selectedRow === row)) return;

      this.selectedTable = table;
      this.selectedRow = row;
      this.selectedColumnIndex = null;
      this.refreshSelectionVisuals();
      return;
    }

    if (this.selectionMode === "column" && this.activeColumnIndex !== null) {
      if (
        this.selectedTable === table &&
        this.selectedColumnIndex === this.activeColumnIndex
      ) {
        return;
      }

      this.selectedTable = table;
      this.selectedRow = null;
      this.selectedColumnIndex = this.activeColumnIndex;
      this.refreshSelectionVisuals();
    }
  }

  private onDocumentPointerDown(event: PointerEvent): void {
    const target = event.target as Element | null;
    if (!target || this.isSelectionControl(target)) return;
    if (target.closest(".hwe-editor-header")) return;

    const selectedTables = this.getSelectedTables();
    const clickedTable = target.closest<HTMLTableElement>(".hwe-page-inner table");
    const clickedCell = target.closest<HTMLTableCellElement>("td, th");
    const root = this.options.rootProvider();

    if (clickedTable && clickedCell && root?.contains(clickedTable)) {
      const belongsToSelection = selectedTables.includes(clickedTable);
      if (selectedTables.length > 0 && !belongsToSelection) this.clearSelection();

      this.hoveredTable = clickedTable;
      this.activeCell = clickedCell;
      this.activeColumnIndex = getColumnIndexAtClientX(
        clickedTable,
        clickedCell,
        event.clientX
      );
      this.isCellTargetLocked = true;
      this.updateSelectionToolbarState();

      if (belongsToSelection) this.updateSelectedScopeFromActiveCell(clickedTable);
      this.positionHandle(clickedTable);
      return;
    }

    this.clearSelection();
  }

  private selectTable(table: HTMLTableElement): void {
    this.getSelectedTables().forEach((fragment) =>
      fragment.classList.remove("hwe-table-selected")
    );
    this.selectedTable = table;
    this.selectedRow = null;
    this.selectedColumnIndex = null;
    this.selectionMode = "table";
    this.isCellTargetLocked = !!this.activeCell;
    this.showSelectionToolbar();
    this.positionHandle(table);
    this.refreshSelectionVisuals();
  }

  private selectActiveRow(): void {
    const row = this.activeCell?.closest<HTMLTableRowElement>("tr") ?? null;
    const table = row?.closest<HTMLTableElement>("table") ?? null;
    if (!row || !table) return;

    this.selectedTable = table;
    this.selectedRow = row;
    this.selectedColumnIndex = null;
    this.selectionMode = "row";
    this.showSelectionToolbar();
    this.refreshSelectionVisuals();
  }

  private selectActiveColumn(): void {
    const table = this.activeCell?.closest<HTMLTableElement>("table") ?? null;
    if (!table || this.activeColumnIndex === null) return;

    this.selectedTable = table;
    this.selectedRow = null;
    this.selectedColumnIndex = this.activeColumnIndex;
    this.selectionMode = "column";
    this.showSelectionToolbar();
    this.refreshSelectionVisuals();
  }

  private getSelectedTables(): HTMLTableElement[] {
    if (!this.selectedTable) return [];
    const root = this.options.rootProvider();
    return getTableFlowFragments(root, this.selectedTable).filter(
      (table) => table.isConnected && !!root?.contains(table)
    );
  }

  private getVisibleHandleTable(): HTMLTableElement | null {
    if (this.hoveredTable?.isConnected) return this.hoveredTable;
    return this.getSelectedTables()[0] ?? null;
  }

  private ensureHandle(): HTMLButtonElement | null {
    if (this.handle) return this.handle;
    const root = this.options.rootProvider();
    if (!root) return null;

    const handle = document.createElement("button");
    handle.type = "button";
    handle.className = "hwe-table-selector";
    handle.contentEditable = "false";
    handle.textContent = "\u2725";
    handle.title = "Seleccionar tabla";
    handle.setAttribute("aria-label", "Seleccionar tabla");
    handle.addEventListener("pointerenter", () => this.cancelHandleHide());
    handle.addEventListener("pointerleave", () => {
      if (!this.selectedTable) this.scheduleHandleHide();
    });
    handle.addEventListener("pointerdown", (event) => {
      this.cancelHandleHide();
      event.preventDefault();
      event.stopPropagation();
      const table = this.hoveredTable ?? this.getSelectedTables()[0];
      if (table) this.selectTable(table);
    });
    root.appendChild(handle);
    this.handle = handle;
    return handle;
  }

  private positionHandle(table: HTMLTableElement): void {
    const handle = this.ensureHandle();
    const root = this.options.rootProvider();
    if (!handle || !root) return;

    const tableRect = table.getBoundingClientRect();
    const rootRect = root.getBoundingClientRect();
    if (tableRect.width <= 0 || tableRect.height <= 0) {
      this.hideHandle();
      return;
    }

    handle.style.left = `${Math.max(0, tableRect.left - rootRect.left - 18)}px`;
    handle.style.top = `${Math.max(0, tableRect.top - rootRect.top - 18)}px`;
    handle.style.display = "flex";
  }

  private ensureSelectionToolbar(): HTMLDivElement | null {
    if (this.selectionToolbar) return this.selectionToolbar;
    const root = this.options.rootProvider();
    if (!root) return null;

    const toolbar = document.createElement("div");
    toolbar.className = "hwe-table-selection-toolbar";
    toolbar.contentEditable = "false";
    toolbar.setAttribute("role", "toolbar");
    toolbar.setAttribute("aria-label", "Seleccion de tabla");

    (["table", "row", "column"] as TableSelectionMode[]).forEach((mode) => {
      const button = document.createElement("button");
      button.type = "button";
      button.textContent = this.getSelectionModeLabel(mode);
      button.title = `Seleccionar ${this.getSelectionModeLabel(mode).toLowerCase()}`;
      button.setAttribute("aria-pressed", "false");
      button.addEventListener("mousedown", (event) => event.preventDefault());
      button.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        if (mode === "table") {
          const table = this.activeCell?.closest<HTMLTableElement>("table") ?? this.selectedTable;
          if (table) this.selectTable(table);
        } else if (mode === "row") {
          this.selectActiveRow();
        } else {
          this.selectActiveColumn();
        }
      });
      toolbar.appendChild(button);
      this.selectionToolbarButtons.set(mode, button);
    });

    const hint = document.createElement("span");
    hint.className = "hwe-table-selection-hint";
    hint.setAttribute("aria-live", "polite");
    toolbar.appendChild(hint);
    this.selectionToolbarHint = hint;

    root.appendChild(toolbar);
    this.selectionToolbar = toolbar;
    return toolbar;
  }

  private showSelectionToolbar(): void {
    const toolbar = this.ensureSelectionToolbar();
    if (!toolbar) return;
    toolbar.style.display = "flex";
    this.updateSelectionToolbarState();
    this.positionSelectionToolbar();
  }

  private hideSelectionToolbar(): void {
    if (this.selectionToolbar) this.selectionToolbar.style.display = "none";
  }

  private positionSelectionToolbar(): void {
    const toolbar = this.selectionToolbar;
    const table = this.getSelectionAnchorTable();
    const root = this.options.rootProvider();
    if (!toolbar || !table || !root || toolbar.style.display === "none") return;

    const tableRect = table.getBoundingClientRect();
    const rootRect = root.getBoundingClientRect();
    const toolbarRect = toolbar.getBoundingClientRect();
    const preferredLeft = tableRect.left - rootRect.left;
    const maxLeft = Math.max(0, rootRect.width - toolbarRect.width);
    const left = Math.max(0, Math.min(preferredLeft, maxLeft));
    const top = Math.max(0, tableRect.top - rootRect.top - toolbarRect.height - 4);

    toolbar.style.left = `${left}px`;
    toolbar.style.top = `${top}px`;
  }

  private updateSelectionToolbarState(): void {
    this.selectionToolbarButtons.forEach((button, mode) => {
      const requiresCell = mode !== "table";
      button.disabled = requiresCell && !this.activeCell;
      const isActive = this.selectionMode === mode;
      button.classList.toggle("hwe-active", isActive);
      button.setAttribute("aria-pressed", String(isActive));
    });

    if (this.selectionToolbarHint) {
      if (this.selectionMode === "row") {
        this.selectionToolbarHint.textContent = "Pulsa una celda de otra fila";
      } else if (this.selectionMode === "column") {
        this.selectionToolbarHint.textContent = "Pulsa una celda de otra columna";
      } else {
        this.selectionToolbarHint.textContent = "Apunta una celda y elige el alcance";
      }
    }
  }

  private getSelectionModeLabel(mode: TableSelectionMode): string {
    if (mode === "table") return "Tabla";
    if (mode === "row") return "Fila";
    return "Columna";
  }

  private getSelectionAnchorTable(): HTMLTableElement | null {
    if (this.selectionMode === "row" && this.selectedRow?.isConnected) {
      return this.selectedRow.closest<HTMLTableElement>("table");
    }
    return this.getSelectedTables()[0] ?? null;
  }

  private hideHandle(): void {
    if (this.handle) this.handle.style.display = "none";
  }

  private scheduleHandleHide(): void {
    if (this.handleHideTimer !== null) return;

    this.handleHideTimer = window.setTimeout(() => {
      this.handleHideTimer = null;
      if (this.selectedTable) return;
      this.hoveredTable = null;
      this.activeCell = null;
      this.activeColumnIndex = null;
      this.hideHandle();
    }, HANDLE_HIDE_DELAY_MS);
  }

  private cancelHandleHide(): void {
    if (this.handleHideTimer === null) return;
    window.clearTimeout(this.handleHideTimer);
    this.handleHideTimer = null;
  }

  private isSelectionControl(target: Element): boolean {
    return !!this.handle?.contains(target) || !!this.selectionToolbar?.contains(target);
  }

  private refreshSelectionVisuals(): void {
    this.clearSelectionOverlays();
    this.updateSelectionToolbarState();
    this.positionSelectionToolbar();

    if (!this.selectionMode) return;
    const elements = this.getSelectedVisualElements();
    elements.forEach((element) => this.addSelectionOverlay(element, this.selectionMode!));
  }

  private getSelectedVisualElements(): HTMLElement[] {
    if (this.selectionMode === "table") return this.getSelectedTables();
    if (this.selectionMode === "row") return this.getSelectedRows();
    if (this.selectionMode === "column") return this.getSelectedColumnCells();
    return [];
  }

  private getSelectedFormattingTargets(): HTMLElement[] {
    return this.getSelectedVisualElements();
  }

  private getSelectedRows(): HTMLTableRowElement[] {
    const row = this.selectedRow;
    if (!row) return [];

    const section = row.closest<HTMLTableSectionElement>("thead");
    if (!section) return row.isConnected ? [row] : [];

    const rowIndex = Array.from(section.rows).indexOf(row);
    if (rowIndex < 0) return [];

    return this.getSelectedTables()
      .map((table) => table.tHead?.rows[rowIndex] ?? null)
      .filter((candidate): candidate is HTMLTableRowElement => !!candidate);
  }

  private getSelectedColumnCells(): HTMLTableCellElement[] {
    if (this.selectedColumnIndex === null) return [];
    const cells = this.getSelectedTables().flatMap((table) =>
      getCellsForColumn(table, this.selectedColumnIndex!)
    );
    return Array.from(new Set(cells));
  }

  private addSelectionOverlay(element: HTMLElement, mode: TableSelectionMode): void {
    const root = this.options.rootProvider();
    if (!root || !element.isConnected) return;

    const rect = element.getBoundingClientRect();
    const rootRect = root.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) return;

    const overlay = document.createElement("div");
    overlay.className = `hwe-table-selection-overlay hwe-table-selection-overlay-${mode}`;
    overlay.contentEditable = "false";
    overlay.setAttribute("aria-hidden", "true");
    overlay.style.left = `${rect.left - rootRect.left}px`;
    overlay.style.top = `${rect.top - rootRect.top}px`;
    overlay.style.width = `${rect.width}px`;
    overlay.style.height = `${rect.height}px`;
    root.appendChild(overlay);
    this.selectionOverlays.push(overlay);
  }

  private clearSelectionOverlays(): void {
    this.selectionOverlays.forEach((overlay) => overlay.remove());
    this.selectionOverlays = [];
  }

  private getTextContainers(root: HTMLElement): HTMLElement[] {
    const containers = new Set<HTMLElement>();
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode: (node) => {
        const parent = node.parentElement;
        return node.textContent && parent && !parent.closest("[contenteditable='false']")
          ? NodeFilter.FILTER_ACCEPT
          : NodeFilter.FILTER_REJECT;
      },
    });

    let current = walker.nextNode();
    while (current) {
      if (current.parentElement) containers.add(current.parentElement);
      current = walker.nextNode();
    }
    return Array.from(containers);
  }

  private isBold(element: HTMLElement): boolean {
    const fontWeight = window.getComputedStyle(element).fontWeight;
    return fontWeight === "bold" || Number(fontWeight) >= 600;
  }

  private isUnderlined(element: HTMLElement): boolean {
    let current: HTMLElement | null = element;
    while (current) {
      if (window.getComputedStyle(current).textDecorationLine.includes("underline")) return true;
      if (current instanceof HTMLTableElement) break;
      current = current.parentElement;
    }
    return false;
  }

  private setUnderline(
    element: HTMLElement,
    value: "none" | "underline",
    boundary: HTMLElement
  ): void {
    let current: HTMLElement | null = element;
    while (current) {
      current.style.textDecoration = value;
      if (value === "underline" || current === boundary) break;
      current = current.parentElement;
    }
  }

  private notifySelectedTableChanged(): void {
    const table = this.getSelectedTables()[0];
    if (!table) return;
    this.options.onTableChanged(table);
    window.requestAnimationFrame(() => this.refresh());
  }
}
