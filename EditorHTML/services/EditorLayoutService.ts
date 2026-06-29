import { getMeaningfulChildren } from "../dom/EditableDom";

const TEXT_FLOW_SELECTOR = "p, h1, h2, h3, h4, h5, h6, blockquote, pre, ul, ol, li";
const LONG_TABLE_MIN_ROWS = 60;
const COMPACT_TABLE_MIN_COLUMNS = 6;
const DENSE_TABLE_MIN_COLUMNS = 8;
const ULTRA_DENSE_TABLE_MIN_COLUMNS = 10;
const COMPACT_TABLE_MAX_COLUMN_PERCENT = 10;
const DENSE_TABLE_MAX_COLUMN_PERCENT = 8;
const ULTRA_DENSE_TABLE_MAX_COLUMN_PERCENT = 6;
const TABLE_DENSITY_CLASSES = [
  "hwe-long-word-table",
  "hwe-table-compact",
  "hwe-table-dense",
  "hwe-table-ultra-dense",
];

type TableDensity = "compact" | "dense" | "ultra-dense" | null;

export class EditorLayoutService {
  applyOfficialTableWidths(root: HTMLElement): void {
    const inners = root.classList.contains("hwe-page-inner")
      ? [root]
      : Array.from(root.querySelectorAll<HTMLElement>(".hwe-page-inner"));

    inners.forEach((inner) => {
      this.refreshFlowClasses(inner);

      inner.querySelectorAll<HTMLElement>(".hwe-table-flow-wrapper").forEach((wrapper) => {
        if (!this.getDirectFlowTable(wrapper)) wrapper.classList.remove("hwe-table-flow-wrapper");
      });

      Array.from(inner.children).forEach((child) => {
        const table = this.getDirectFlowTable(child as HTMLElement);
        if (!table) return;

        if ((child as HTMLElement) !== table) {
          (child as HTMLElement).classList.add("hwe-table-flow-wrapper");
        }
        this.normalizeFlowTable(table);
      });
    });
  }

  pageOverflows(page: HTMLElement): boolean {
    const inner = page.querySelector<HTMLElement>(".hwe-page-inner");
    if (!inner) return false;

    const scrollOverflows = inner.scrollHeight > inner.clientHeight + 1;
    const contentBottom = this.getContentBottom(inner);
    if (contentBottom === null) return scrollOverflows;

    return contentBottom > this.getContentLimitBottom(inner) + 1 || scrollOverflows;
  }

  private getDirectFlowTable(element: HTMLElement): HTMLTableElement | null {
    if (element.tagName === "TABLE") return element as HTMLTableElement;

    const children = getMeaningfulChildren(element, true);
    if (children.length !== 1 || children[0].nodeType !== Node.ELEMENT_NODE) return null;

    const onlyChild = children[0] as HTMLElement;
    return this.getDirectFlowTable(onlyChild);
  }

  private normalizeFlowTable(table: HTMLTableElement): void {
    table.removeAttribute("width");
    table.style.setProperty("width", "100%", "important");
    table.style.setProperty("max-width", "100%", "important");
    table.style.setProperty("margin-left", "0", "important");
    table.style.setProperty("margin-right", "0", "important");
    this.normalizeColumnWidths(table);
    this.clearTableStructureHeights(table);
    this.normalizeTableDensity(table);
  }

  private normalizeColumnWidths(table: HTMLTableElement): void {
    const columns = Array.from(table.querySelectorAll<HTMLTableColElement>("col"));
    if (columns.length === 0) return;

    const widths = columns.map((column) => this.readColumnWidth(column));
    const numericWidths = widths.map((width) => width ?? 0);
    const total = numericWidths.reduce((sum, width) => sum + width, 0);
    if (total <= 0) return;

    columns.forEach((column, index) => {
      const percent = Math.max(1, (numericWidths[index] / total) * 100);
      column.removeAttribute("width");
      column.style.setProperty("width", `${percent.toFixed(3)}%`, "important");
    });
  }

  private readColumnWidth(column: HTMLTableColElement): number | null {
    const styleWidth = column.style.getPropertyValue("width").trim();
    const attrWidth = column.getAttribute("width")?.trim() ?? "";
    return this.parseCssLength(styleWidth) ?? this.parseCssLength(attrWidth);
  }

  private normalizeTableDensity(table: HTMLTableElement): void {
    const rowCount = table.querySelectorAll("tbody tr, tfoot tr").length;
    const columnCount = this.getColumnCount(table);
    const minColumnPercent = this.getMinimumColumnPercent(table);
    const density = this.getTableDensity(rowCount, columnCount, minColumnPercent);

    TABLE_DENSITY_CLASSES.forEach((className) => table.classList.remove(className));
    if (!density) return;

    table.classList.add(`hwe-table-${density}`);
    table.classList.add("hwe-long-word-table");
  }

  private getTableDensity(
    rowCount: number,
    columnCount: number,
    minColumnPercent: number | null
  ): TableDensity {
    if (
      columnCount >= ULTRA_DENSE_TABLE_MIN_COLUMNS ||
      (minColumnPercent !== null && minColumnPercent <= ULTRA_DENSE_TABLE_MAX_COLUMN_PERCENT)
    ) {
      return "ultra-dense";
    }
    if (
      columnCount >= DENSE_TABLE_MIN_COLUMNS ||
      (minColumnPercent !== null && minColumnPercent <= DENSE_TABLE_MAX_COLUMN_PERCENT)
    ) {
      return "dense";
    }
    if (
      rowCount >= LONG_TABLE_MIN_ROWS ||
      columnCount >= COMPACT_TABLE_MIN_COLUMNS ||
      (minColumnPercent !== null && minColumnPercent <= COMPACT_TABLE_MAX_COLUMN_PERCENT)
    ) {
      return "compact";
    }

    return null;
  }

  private getMinimumColumnPercent(table: HTMLTableElement): number | null {
    const columns = Array.from(table.querySelectorAll<HTMLTableColElement>("col"));
    if (columns.length === 0) return null;

    const widths = columns.map((column) => this.readColumnWidth(column) ?? 0);
    const total = widths.reduce((sum, width) => sum + width, 0);
    if (total <= 0) return null;

    return widths.reduce<number | null>((minimum, width) => {
      if (width <= 0) return minimum;
      const percent = (width / total) * 100;
      return minimum === null ? percent : Math.min(minimum, percent);
    }, null);
  }

  private getColumnCount(table: HTMLTableElement): number {
    const explicitColumns = table.querySelectorAll("col").length;
    if (explicitColumns > 0) return explicitColumns;

    const firstRow = table.querySelector("tr");
    if (!firstRow) return 0;

    return Array.from(firstRow.children).reduce((count, cell) => {
      const span = Number.parseInt(cell.getAttribute("colspan") ?? "1", 10);
      return count + (Number.isFinite(span) && span > 0 ? span : 1);
    }, 0);
  }

  private clearFixedHeight(element: HTMLElement): void {
    element.removeAttribute("height");
    element.style.removeProperty("height");
    element.style.removeProperty("min-height");
    element.style.removeProperty("max-height");
  }

  private clearTableStructureHeights(table: HTMLTableElement): void {
    table
      .querySelectorAll<HTMLElement>("tr, td, th")
      .forEach((element) => this.clearFixedHeight(element));
  }

  private parseCssLength(value: string): number | null {
    if (!value) return null;
    const match = /^([0-9.]+)\s*(px|pt|in|cm|mm|%)?$/i.exec(value);
    if (!match) return null;

    const amount = Number.parseFloat(match[1]);
    if (!Number.isFinite(amount) || amount <= 0) return null;

    const unit = (match[2] || "px").toLowerCase();
    if (unit === "pt") return amount * (96 / 72);
    if (unit === "in") return amount * 96;
    if (unit === "cm") return amount * (96 / 2.54);
    if (unit === "mm") return amount * (96 / 25.4);
    return amount;
  }

  private refreshFlowClasses(inner: HTMLElement): void {
    inner
      .querySelectorAll<HTMLElement>(".hwe-text-flow-block, .hwe-image-flow-block")
      .forEach((element) => {
        element.classList.remove("hwe-text-flow-block", "hwe-image-flow-block");
      });

    inner.querySelectorAll<HTMLElement>(TEXT_FLOW_SELECTOR).forEach((element) => {
      if (element.closest("td, th")) return;
      if (element.querySelector("img, table, tr, td, th, figure, video, canvas, svg")) return;
      element.classList.add("hwe-text-flow-block");
    });

    inner.querySelectorAll<HTMLImageElement>("img").forEach((image) => {
      if (image.closest("td, th")) return;
      image.classList.add("hwe-image-flow-block");
    });
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
}
