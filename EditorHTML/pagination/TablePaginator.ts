import { hweDebugLog } from "../debug/DebugLogger";

export interface TablePaginatorContext {
  getContentLimitBottom(inner: HTMLElement): number;
  getInner(page: HTMLElement): HTMLElement | null;
}

export class TablePaginator {
  private static tableFlowCounter = 0;

  splitTable(
    table: HTMLElement,
    targetInner: HTMLElement,
    page: HTMLElement,
    context: TablePaginatorContext
  ): boolean {
    const rows = this.getSplittableTableRows(table);
    if (rows.length <= 1) return false;

    const inner = context.getInner(page);
    const pageBottom = inner
      ? context.getContentLimitBottom(inner)
      : page.getBoundingClientRect().bottom;
    const splitRowIndex = this.findFirstOverflowingRow(rows, pageBottom);

    if (splitRowIndex <= 0) return false;

    const rowsForNext = rows.slice(splitRowIndex);
    if (rowsForNext.length === 0) return false;

    const flowId = this.ensureTableFlowId(table);
    table.setAttribute("data-hwe-table-fragment", "true");

    const newTable = this.cloneTableShell(table);
    newTable.setAttribute("data-hwe-table-flow-id", flowId);
    newTable.setAttribute("data-hwe-table-fragment", "true");

    const tbody = document.createElement("tbody");
    rowsForNext.forEach((row) => tbody.appendChild(row));
    newTable.appendChild(tbody);

    targetInner.insertBefore(newTable, targetInner.firstChild);
    hweDebugLog("tablePaginator.splitTable", {
      rows: rows.length,
      rowsForNext: rowsForNext.length,
      rowsRemaining: rows.length - rowsForNext.length,
      splitRowIndex,
    });

    const remainingRows = table.querySelectorAll("tbody tr, tfoot tr");
    if (remainingRows.length === 0) table.remove();

    return true;
  }

  private ensureTableFlowId(table: HTMLElement): string {
    const existingId = table.getAttribute("data-hwe-table-flow-id");
    if (existingId) return existingId;

    const id = `hwe-table-${Date.now().toString(36)}-${TablePaginator.tableFlowCounter++}`;
    table.setAttribute("data-hwe-table-flow-id", id);
    return id;
  }

  private cloneTableShell(table: HTMLElement): HTMLElement {
    const newTable = document.createElement("table");
    Array.from(table.attributes).forEach((attr) => {
      newTable.setAttribute(attr.name, attr.value);
    });

    const colgroup = table.querySelector("colgroup");
    if (colgroup) newTable.appendChild(colgroup.cloneNode(true));

    const thead = table.querySelector("thead");
    if (thead && this.shouldRepeatTableHeader(table)) {
      newTable.appendChild(thead.cloneNode(true));
    }

    return newTable;
  }

  private findFirstOverflowingRow(rows: HTMLElement[], pageBottom: number): number {
    return rows.findIndex((row) => row.getBoundingClientRect().bottom > pageBottom + 0.5);
  }

  private getSplittableTableRows(table: HTMLElement): HTMLElement[] {
    const thead = table.querySelector("thead");
    const allRows = Array.from(table.querySelectorAll("tr")) as HTMLElement[];

    return allRows.filter((row) => !thead?.contains(row));
  }

  private shouldRepeatTableHeader(table: HTMLElement): boolean {
    return table.getAttribute("data-hwe-repeat-header") === "true";
  }
}
