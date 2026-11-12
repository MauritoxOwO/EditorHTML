import { getMeaningfulChildren, unwrapElement } from "../dom/EditableDom";
import { getTableFlowFragments } from "../dom/TableFlow";
import { getPageContentBounds } from "../pagination/PageContentGeometry";

const TEXT_FLOW_SELECTOR = "p, h1, h2, h3, h4, h5, h6, blockquote, pre, ul, ol";
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
  applyOfficialTableWidthsToPages(pages: HTMLElement[]): void {
    // Density depends on the entire table. Count each flow once per layout
    // update rather than scanning every fragment again for every page.
    const rowCounts = new Map<string, number>();
    pages.forEach((page) => page.querySelectorAll<HTMLTableElement>("table[data-hwe-table-flow-id]").forEach((table) => {
      const id = table.getAttribute("data-hwe-table-flow-id")!;
      rowCounts.set(id, (rowCounts.get(id) ?? 0) + table.querySelectorAll("tbody tr, tfoot tr").length);
    }));
    pages.forEach((page) => this.applyOfficialTableWidths(page, rowCounts));
  }

  applyOfficialTableWidths(root: HTMLElement, rowCounts?: ReadonlyMap<string, number>): void {
    const inners = root.classList.contains("hwe-page-inner")
      ? [root]
      : Array.from(root.querySelectorAll<HTMLElement>(".hwe-page-inner"));

    inners.forEach((inner) => {
      this.applyTextFlowLayout(inner);

      inner.querySelectorAll<HTMLElement>(".hwe-table-flow-wrapper").forEach((wrapper) => {
        if (!this.getDirectFlowTable(wrapper)) wrapper.classList.remove("hwe-table-flow-wrapper");
      });

      Array.from(inner.children).forEach((child) => {
        const table = this.getDirectFlowTable(child as HTMLElement);
        if (!table) return;

        if ((child as HTMLElement) !== table) {
          (child as HTMLElement).classList.add("hwe-table-flow-wrapper");
        }
        this.normalizeFlowTable(table, rowCounts);
      });
    });
  }

  pageOverflows(page: HTMLElement): boolean {
    const inner = page.querySelector<HTMLElement>(".hwe-page-inner");
    if (!inner) return false;

    const scrollOverflows = inner.scrollHeight > inner.clientHeight + 1;
    const contentBottom = this.getContentBottom(inner);
    if (contentBottom === null) return scrollOverflows;

    const bounds = getPageContentBounds(inner);
    return contentBottom > bounds.bottom + bounds.scale || scrollOverflows;
  }

  private getDirectFlowTable(element: HTMLElement): HTMLTableElement | null {
    if (element.tagName === "TABLE") return element as HTMLTableElement;

    const children = getMeaningfulChildren(element, true);
    if (children.length !== 1 || children[0].nodeType !== Node.ELEMENT_NODE) return null;

    const onlyChild = children[0] as HTMLElement;
    return this.getDirectFlowTable(onlyChild);
  }

  private normalizeFlowTable(table: HTMLTableElement, rowCounts?: ReadonlyMap<string, number>): void {
    table.removeAttribute("width");
    if (!table.hasAttribute("data-hwe-user-table-width")) {
      table.style.setProperty("width", "100%", "important");
      table.style.setProperty("margin-left", "0", "important");
    }
    table.style.setProperty("max-width", "100%", "important");
    table.style.setProperty("margin-right", "0", "important");
    this.normalizeColumnWidths(table);
    this.clearTableStructureHeights(table);
    this.normalizeTableDensity(table, rowCounts);
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

  private normalizeTableDensity(table: HTMLTableElement, rowCounts?: ReadonlyMap<string, number>): void {
    // Density belongs to the complete table, not to each page fragment.
    // Otherwise splitting a long table removes its compact style and increases row heights.
    const scope = table.closest(".hwe-workspace, .pcf-html-editor-root");
    const id = table.getAttribute("data-hwe-table-flow-id");
    const rowCount = (id ? rowCounts?.get(id) : undefined) ?? getTableFlowFragments(scope, table).reduce(
      (count, fragment) => count + fragment.querySelectorAll("tbody tr, tfoot tr").length,
      0
    );
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

  applyTextFlowLayout(inner: HTMLElement): void {
    const textBlocks = new Set<HTMLElement>();
    inner.querySelectorAll<HTMLElement>(TEXT_FLOW_SELECTOR).forEach((element) => {
      if (element.closest("td, th")) return;
      if (element.closest("li")) return;
      if (element.querySelector("img, table, tr, td, th, figure, video, canvas, svg")) return;
      textBlocks.add(element);
    });
    const images = new Set<HTMLElement>();
    inner.querySelectorAll<HTMLImageElement>("img").forEach((image) => {
      if (image.closest("td, th")) return;
      images.add(image);
    });
    // Keep unchanged classes intact: removing and adding them invalidates style
    // and layout throughout a long document on every pagination callback.
    inner.querySelectorAll<HTMLElement>(".hwe-text-flow-block").forEach((element) => {
      if (!textBlocks.has(element)) element.classList.remove("hwe-text-flow-block");
    });
    inner.querySelectorAll<HTMLElement>(".hwe-image-flow-block").forEach((element) => {
      if (!images.has(element)) element.classList.remove("hwe-image-flow-block");
    });
    textBlocks.forEach((element) => {
      if (!element.classList.contains("hwe-text-flow-block")) element.classList.add("hwe-text-flow-block");
    });
    images.forEach((image) => {
      if (!image.classList.contains("hwe-image-flow-block")) image.classList.add("hwe-image-flow-block");
    });

    inner.querySelectorAll<HTMLElement>(".hwe-text-flow-column").forEach((column) => {
      if (!column.children.length || Array.from(column.children).some(
        (child) => !child.classList.contains("hwe-text-flow-block")
      )) unwrapElement(column);
    });
    inner.querySelectorAll<HTMLElement>(".hwe-text-flow-block").forEach((element) => {
      if (element.parentElement?.closest(".hwe-text-flow-block, .hwe-text-flow-column")) return;
      const column = document.createElement("div");
      column.className = "hwe-text-flow-column";
      column.setAttribute("data-hwe-generated-wrapper", "true");
      element.before(column);
      column.appendChild(element);
    });
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
