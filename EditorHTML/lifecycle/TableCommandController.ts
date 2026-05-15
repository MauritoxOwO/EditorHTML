export interface TableCommandContext {
  getActiveEditable: () => HTMLElement | null;
  getEditableForPageIndex: (pageIndex: number) => HTMLElement | null;
  markEdited: (element: HTMLElement) => void;
  rootProvider: () => HTMLElement;
}

export class TableCommandController {
  private lastSelectedTableRow: HTMLTableRowElement | null = null;

  constructor(private readonly context: TableCommandContext) {}

  insertTable(): void {
    const editable =
      this.context.getActiveEditable() ?? this.context.getEditableForPageIndex(0);
    if (!editable) return;

    editable.focus({ preventScroll: true });
    document.execCommand(
      "insertHTML",
      false,
      '<table class="hwe-word-table" data-hwe-source="manual" data-hwe-table="word" style="border-collapse:collapse;width:100%"><tbody><tr><td style="border:solid windowtext 1.0pt;padding:2.85pt 4.25pt"><p><br></p></td><td style="border:solid windowtext 1.0pt;padding:2.85pt 4.25pt"><p><br></p></td></tr><tr><td style="border:solid windowtext 1.0pt;padding:2.85pt 4.25pt"><p><br></p></td><td style="border:solid windowtext 1.0pt;padding:2.85pt 4.25pt"><p><br></p></td></tr></tbody></table><p><br></p>'
    );
    this.context.markEdited(editable);
  }

  insertTableRowAfter(): void {
    const selectionElement = this.getSelectionElement();
    const row = selectionElement?.closest("tr") ?? this.lastSelectedTableRow;
    if (!row || !this.context.rootProvider().contains(row)) {
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
    this.context.markEdited(row);
  }

  rememberSelectedTableRow(): void {
    const row = this.getSelectionElement()?.closest("tr");
    if (row && this.context.rootProvider().contains(row)) {
      this.lastSelectedTableRow = row as HTMLTableRowElement;
    }
  }

  rememberTableRowFromEvent(event: Event): void {
    const target = event.target;
    if (!(target instanceof Element)) return;

    const row = target.closest("tr");
    if (row && this.context.rootProvider().contains(row)) {
      this.lastSelectedTableRow = row as HTMLTableRowElement;
    }
  }

  private getSelectionElement(): HTMLElement | null {
    const selection = window.getSelection();
    const node = selection?.anchorNode;
    if (!node) return null;
    return node.nodeType === Node.ELEMENT_NODE
      ? (node as HTMLElement)
      : node.parentElement;
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
}
