import { getTableFlowFragments } from "../dom/TableFlow";

export interface TableColumnResizeControllerOptions {
  onColumnsChanged: (table: HTMLTableElement) => void;
  rootProvider: () => HTMLElement | null;
}

interface ColumnResizeHit {
  boundaryIndex: number;
  table: HTMLTableElement;
}

interface ColumnResizeDragState {
  affectedTables: HTMLTableElement[];
  boundaryIndex: number;
  columnCount: number;
  currentWidths: number[];
  frame: number | undefined;
  hasChanged: boolean;
  latestClientX: number;
  startClientX: number;
  startWidths: number[];
  table: HTMLTableElement;
  tableWidth: number;
}

const TABLE_SELECTOR = ".hwe-page-inner table";
const HIT_ZONE_PX = 6;
const MIN_COLUMN_WIDTH_PX = 24;
const MAX_MEASURED_ROWS = 12;

export class TableColumnResizeController {
  private dragState: ColumnResizeDragState | null = null;
  private guide: HTMLElement | null = null;

  private readonly handleRootPointerDown = (event: PointerEvent): void =>
    this.onRootPointerDown(event);
  private readonly handleRootPointerMove = (event: PointerEvent): void =>
    this.onRootPointerMove(event);
  private readonly handleRootPointerLeave = (): void => this.setResizeCursor(false);
  private readonly handleDocumentPointerMove = (event: PointerEvent): void =>
    this.onDocumentPointerMove(event);
  private readonly handleDocumentPointerUp = (): void => this.onDocumentPointerUp();
  private readonly handleSelectStart = (event: Event): void => {
    if (this.dragState) event.preventDefault();
  };
  private readonly handleScrollOrResize = (): void => this.positionGuide();

  constructor(private readonly options: TableColumnResizeControllerOptions) {}

  start(): void {
    const root = this.options.rootProvider();
    if (!root) return;

    root.addEventListener("pointerdown", this.handleRootPointerDown, true);
    root.addEventListener("pointermove", this.handleRootPointerMove, true);
    root.addEventListener("pointerleave", this.handleRootPointerLeave);
    root.addEventListener("scroll", this.handleScrollOrResize, true);
    document.addEventListener("pointermove", this.handleDocumentPointerMove);
    document.addEventListener("pointerup", this.handleDocumentPointerUp);
    document.addEventListener("pointercancel", this.handleDocumentPointerUp);
    document.addEventListener("selectstart", this.handleSelectStart, true);
    window.addEventListener("resize", this.handleScrollOrResize);
  }

  destroy(): void {
    const root = this.options.rootProvider();
    root?.removeEventListener("pointerdown", this.handleRootPointerDown, true);
    root?.removeEventListener("pointermove", this.handleRootPointerMove, true);
    root?.removeEventListener("pointerleave", this.handleRootPointerLeave);
    root?.removeEventListener("scroll", this.handleScrollOrResize, true);
    document.removeEventListener("pointermove", this.handleDocumentPointerMove);
    document.removeEventListener("pointerup", this.handleDocumentPointerUp);
    document.removeEventListener("pointercancel", this.handleDocumentPointerUp);
    document.removeEventListener("selectstart", this.handleSelectStart, true);
    window.removeEventListener("resize", this.handleScrollOrResize);
    this.clear();
  }

  clear(): void {
    if (this.dragState?.frame !== undefined) {
      window.cancelAnimationFrame(this.dragState.frame);
    }
    this.dragState = null;
    this.guide?.remove();
    this.guide = null;
    this.setResizeCursor(false);
  }

  private onRootPointerMove(event: PointerEvent): void {
    if (this.dragState) return;
    this.setResizeCursor(this.getResizeHit(event) !== null);
  }

  private onRootPointerDown(event: PointerEvent): void {
    if (event.button !== 0) return;

    const hit = this.getResizeHit(event);
    if (!hit) return;

    event.preventDefault();
    event.stopPropagation();

    const columnCount = this.getColumnCount(hit.table);
    if (columnCount < 2) return;

    const tableWidth = Math.max(1, hit.table.getBoundingClientRect().width);
    const startWidths = this.readColumnWidths(hit.table, columnCount, tableWidth);

    this.dragState = {
      affectedTables: getTableFlowFragments(this.options.rootProvider(), hit.table),
      boundaryIndex: hit.boundaryIndex,
      columnCount,
      currentWidths: startWidths,
      frame: undefined,
      hasChanged: false,
      latestClientX: event.clientX,
      startClientX: event.clientX,
      startWidths,
      table: hit.table,
      tableWidth,
    };
    this.setResizeCursor(true);
    this.ensureGuide();
    this.positionGuide();
  }

  private onDocumentPointerMove(event: PointerEvent): void {
    if (!this.dragState) return;

    event.preventDefault();
    this.dragState.latestClientX = event.clientX;
    if (this.dragState.frame !== undefined) return;

    this.dragState.frame = window.requestAnimationFrame(() => {
      if (!this.dragState) return;

      this.dragState.frame = undefined;
      this.applyDragPreview(this.dragState);
    });
  }

  private onDocumentPointerUp(): void {
    if (!this.dragState) return;

    if (this.dragState.frame !== undefined) {
      window.cancelAnimationFrame(this.dragState.frame);
      this.dragState.frame = undefined;
      this.applyDragPreview(this.dragState);
    }

    if (!this.dragState.hasChanged) {
      this.clear();
      return;
    }

    const { affectedTables, currentWidths, table } = this.dragState;
    this.applyColumnWidths(affectedTables, currentWidths);
    this.options.onColumnsChanged(table);
    this.clear();
  }

  private applyDragPreview(state: ColumnResizeDragState): void {
    const nextWidths = this.getDraggedWidths(state);
    state.hasChanged =
      state.hasChanged ||
      nextWidths.some((width, index) => Math.abs(width - state.startWidths[index]) > 0.5);
    state.currentWidths = nextWidths;
    if (!state.hasChanged) {
      this.positionGuide();
      return;
    }

    this.applyColumnWidths([state.table], nextWidths);
    this.positionGuide();
  }

  private getDraggedWidths(state: ColumnResizeDragState): number[] {
    const widths = [...state.startWidths];
    const leftIndex = state.boundaryIndex;
    const rightIndex = leftIndex + 1;
    const pairWidth = state.startWidths[leftIndex] + state.startWidths[rightIndex];
    const minWidth = Math.min(MIN_COLUMN_WIDTH_PX, Math.max(8, pairWidth * 0.15));
    const rawDelta = state.latestClientX - state.startClientX;
    const minDelta = minWidth - state.startWidths[leftIndex];
    const maxDelta = state.startWidths[rightIndex] - minWidth;
    const delta = this.clamp(rawDelta, minDelta, maxDelta);

    widths[leftIndex] = state.startWidths[leftIndex] + delta;
    widths[rightIndex] = state.startWidths[rightIndex] - delta;
    return widths;
  }

  private getResizeHit(event: PointerEvent): ColumnResizeHit | null {
    const target = event.target as Element | null;
    if (!target?.closest || target.closest(".hwe-editor-header")) return null;

    const cell = target.closest<HTMLTableCellElement>("td, th");
    if (!cell) return null;

    const table = cell.closest<HTMLTableElement>("table");
    const root = this.options.rootProvider();
    if (!table || !root?.contains(table) || !table.closest(".hwe-page-inner")) return null;

    const rect = cell.getBoundingClientRect();
    if (event.clientY < rect.top || event.clientY > rect.bottom) return null;
    if (Math.abs(event.clientX - rect.right) > HIT_ZONE_PX) return null;

    const columnCount = this.getColumnCount(table);
    const cellEndIndex = this.getCellEndColumnIndex(cell);
    if (cellEndIndex === null || cellEndIndex <= 0 || cellEndIndex >= columnCount) return null;

    return {
      boundaryIndex: cellEndIndex - 1,
      table,
    };
  }

  private getCellEndColumnIndex(cell: HTMLTableCellElement): number | null {
    const row = cell.parentElement as HTMLTableRowElement | null;
    if (!row) return null;

    let columnIndex = 0;
    for (const currentCell of Array.from(row.cells)) {
      columnIndex += this.getColSpan(currentCell);
      if (currentCell === cell) return columnIndex;
    }

    return null;
  }

  private getColumnCount(table: HTMLTableElement): number {
    const explicitColumns = table.querySelectorAll(":scope > colgroup > col").length;
    if (explicitColumns > 0) return explicitColumns;

    return Array.from(table.rows).reduce((max, row) => {
      const columns = Array.from(row.cells).reduce(
        (count, cell) => count + this.getColSpan(cell),
        0
      );
      return Math.max(max, columns);
    }, 0);
  }

  private getColSpan(cell: HTMLTableCellElement): number {
    const span = Number.parseInt(cell.getAttribute("colspan") ?? "1", 10);
    return Number.isFinite(span) && span > 0 ? span : 1;
  }

  private readColumnWidths(
    table: HTMLTableElement,
    columnCount: number,
    tableWidth: number
  ): number[] {
    const colgroupWidths = this.readColgroupWidths(table, columnCount, tableWidth);
    if (colgroupWidths) return colgroupWidths;

    const totals = Array.from({ length: columnCount }, () => 0);
    const counts = Array.from({ length: columnCount }, () => 0);
    const rows = Array.from(table.rows).slice(0, MAX_MEASURED_ROWS);

    rows.forEach((row) => {
      let columnIndex = 0;
      Array.from(row.cells).forEach((cell) => {
        const span = this.getColSpan(cell);
        const width = Math.max(0, cell.getBoundingClientRect().width / span);
        for (let offset = 0; offset < span && columnIndex + offset < columnCount; offset++) {
          totals[columnIndex + offset] += width;
          counts[columnIndex + offset]++;
        }
        columnIndex += span;
      });
    });

    const fallbackWidth = tableWidth / columnCount;
    const widths = totals.map((total, index) =>
      counts[index] > 0 && total > 0 ? total / counts[index] : fallbackWidth
    );
    return this.normalizeWidths(widths, tableWidth);
  }

  private readColgroupWidths(
    table: HTMLTableElement,
    columnCount: number,
    tableWidth: number
  ): number[] | null {
    const columns = Array.from(table.querySelectorAll<HTMLTableColElement>(":scope > colgroup > col"));
    if (columns.length !== columnCount) return null;

    const widths = columns.map((column) => {
      const value =
        column.style.getPropertyValue("width").trim() ||
        column.getAttribute("width")?.trim() ||
        "";
      return this.parseCssLength(value, tableWidth);
    });

    if (widths.some((width) => width === null)) return null;
    return this.normalizeWidths(widths as number[], tableWidth);
  }

  private normalizeWidths(widths: number[], tableWidth: number): number[] {
    const total = widths.reduce((sum, width) => sum + Math.max(0, width), 0);
    if (total <= 0) return widths.map(() => tableWidth / widths.length);

    const factor = tableWidth / total;
    return widths.map((width) => Math.max(1, width * factor));
  }

  private applyColumnWidths(tables: HTMLTableElement[], widths: number[]): void {
    tables.forEach((table) => this.ensureColgroup(table, widths));
  }

  private ensureColgroup(table: HTMLTableElement, widths: number[]): void {
    table.removeAttribute("width");
    table.style.setProperty("width", "100%", "important");
    table.style.setProperty("max-width", "100%", "important");
    table.style.setProperty("table-layout", "fixed", "important");
    table.setAttribute("data-hwe-user-column-widths", "true");

    let colgroup = this.getDirectColgroup(table);
    if (!colgroup) {
      colgroup = document.createElement("colgroup");
      table.insertBefore(colgroup, table.firstChild);
    }

    colgroup.innerHTML = "";
    const total = widths.reduce((sum, width) => sum + width, 0) || 1;
    widths.forEach((width) => {
      const column = document.createElement("col");
      column.style.setProperty("width", `${((width / total) * 100).toFixed(3)}%`, "important");
      colgroup.appendChild(column);
    });
  }

  private getDirectColgroup(table: HTMLTableElement): HTMLTableColElement | null {
    return Array.from(table.children).find(
      (child) => child.tagName === "COLGROUP"
    ) as HTMLTableColElement | null;
  }

  private ensureGuide(): void {
    if (this.guide) return;

    const root = this.options.rootProvider();
    if (!root) return;

    this.guide = document.createElement("div");
    this.guide.className = "hwe-table-column-resize-guide";
    root.appendChild(this.guide);
  }

  private positionGuide(): void {
    if (!this.dragState || !this.guide) return;

    const root = this.options.rootProvider();
    if (!root) return;

    const tableRect = this.dragState.table.getBoundingClientRect();
    const rootRect = root.getBoundingClientRect();
    const widthBeforeBoundary = this.dragState.currentWidths
      .slice(0, this.dragState.boundaryIndex + 1)
      .reduce((sum, width) => sum + width, 0);
    const totalWidth = this.dragState.currentWidths.reduce((sum, width) => sum + width, 0) || 1;
    const boundaryLeft = tableRect.left + (widthBeforeBoundary / totalWidth) * tableRect.width;

    this.guide.style.left = `${boundaryLeft - rootRect.left + root.scrollLeft}px`;
    this.guide.style.top = `${tableRect.top - rootRect.top + root.scrollTop}px`;
    this.guide.style.height = `${tableRect.height}px`;
  }

  private setResizeCursor(active: boolean): void {
    this.options.rootProvider()?.classList.toggle("hwe-table-column-resize-ready", active);
    this.options.rootProvider()?.classList.toggle("hwe-table-column-resizing", !!this.dragState);
  }

  private parseCssLength(value: string, referenceWidth: number): number | null {
    if (!value) return null;

    const match = /^([0-9.]+)\s*(px|pt|in|cm|mm|%)?$/i.exec(value);
    if (!match) return null;

    const amount = Number.parseFloat(match[1]);
    if (!Number.isFinite(amount) || amount <= 0) return null;

    const unit = (match[2] || "px").toLowerCase();
    if (unit === "%") return (amount / 100) * referenceWidth;
    if (unit === "pt") return amount * (96 / 72);
    if (unit === "in") return amount * 96;
    if (unit === "cm") return amount * (96 / 2.54);
    if (unit === "mm") return amount * (96 / 25.4);
    return amount;
  }

  private clamp(value: number, min: number, max: number): number {
    if (max < min) return min;
    return Math.max(min, Math.min(max, value));
  }
}
