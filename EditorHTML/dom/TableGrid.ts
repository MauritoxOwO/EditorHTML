export interface TableGridCell {
  cell: HTMLTableCellElement;
  row: HTMLTableRowElement;
  rowIndex: number;
  columnStart: number;
  columnEnd: number;
}

export function getTableGridCells(table: HTMLTableElement): TableGridCell[] {
  const placements: TableGridCell[] = [];
  const occupiedUntilRow: number[] = [];

  Array.from(table.rows).forEach((row, rowIndex) => {
    let columnIndex = 0;

    Array.from(row.cells).forEach((cell) => {
      const columnSpan = readPositiveSpan(cell.getAttribute("colspan"));
      const rowSpan = readPositiveSpan(cell.getAttribute("rowspan"));

      columnIndex = findFirstAvailableColumn(
        occupiedUntilRow,
        rowIndex,
        columnIndex,
        columnSpan
      );

      const columnEnd = columnIndex + columnSpan;
      placements.push({
        cell,
        row,
        rowIndex,
        columnStart: columnIndex,
        columnEnd,
      });

      for (let index = columnIndex; index < columnEnd; index++) {
        occupiedUntilRow[index] = Math.max(
          occupiedUntilRow[index] ?? 0,
          rowIndex + rowSpan
        );
      }
      columnIndex = columnEnd;
    });
  });

  return placements;
}

export function getColumnIndexAtClientX(
  table: HTMLTableElement,
  cell: HTMLTableCellElement,
  clientX: number
): number | null {
  const placement = getTableGridCells(table).find((candidate) => candidate.cell === cell);
  if (!placement) return null;

  const span = placement.columnEnd - placement.columnStart;
  if (span <= 1) return placement.columnStart;

  const rect = cell.getBoundingClientRect();
  if (rect.width <= 0) return placement.columnStart;

  const ratio = Math.max(0, Math.min(0.999999, (clientX - rect.left) / rect.width));
  return placement.columnStart + Math.floor(ratio * span);
}

export function getCellsForColumn(
  table: HTMLTableElement,
  columnIndex: number
): HTMLTableCellElement[] {
  return getTableGridCells(table)
    .filter(
      (placement) =>
        placement.columnStart <= columnIndex && columnIndex < placement.columnEnd
    )
    .map((placement) => placement.cell);
}

function findFirstAvailableColumn(
  occupiedUntilRow: number[],
  rowIndex: number,
  initialColumn: number,
  columnSpan: number
): number {
  let candidate = initialColumn;

  while (
    Array.from({ length: columnSpan }, (_, offset) => candidate + offset).some(
      (columnIndex) => (occupiedUntilRow[columnIndex] ?? 0) > rowIndex
    )
  ) {
    candidate++;
  }

  return candidate;
}

function readPositiveSpan(value: string | null): number {
  const parsed = Number.parseInt(value ?? "1", 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 1;
}
