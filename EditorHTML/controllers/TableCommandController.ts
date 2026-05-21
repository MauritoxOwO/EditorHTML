export interface TableCommandContext {
  getActiveEditable: () => HTMLElement | null;
  getEditableForPageIndex: (pageIndex: number) => HTMLElement | null;
  markEdited: (element: HTMLElement) => void;
  rootProvider: () => HTMLElement;
}

export class TableCommandController {
  private lastSelectedTableRow: HTMLTableRowElement | null = null;
  private contextMenu: HTMLDivElement | null = null;
  private isStarted = false;
  private readonly handleContextMenu = (event: MouseEvent): void => this.onContextMenu(event);
  private readonly handleDocumentPointerDown = (event: PointerEvent): void => {
    if (this.contextMenu?.contains(event.target as Node)) return;
    this.hideContextMenu();
  };
  private readonly handleDocumentKeyDown = (event: KeyboardEvent): void => {
    if (event.key === "Escape") this.hideContextMenu();
  };
  private readonly handleAnyScroll = (): void => this.hideContextMenu();

  constructor(private readonly context: TableCommandContext) {}

  start(): void {
    if (this.isStarted) return;
    this.isStarted = true;
    this.context.rootProvider().addEventListener("contextmenu", this.handleContextMenu, true);
    document.addEventListener("pointerdown", this.handleDocumentPointerDown, true);
    document.addEventListener("keydown", this.handleDocumentKeyDown, true);
    window.addEventListener("scroll", this.handleAnyScroll, true);
  }

  destroy(): void {
    if (!this.isStarted) return;
    this.isStarted = false;
    this.context.rootProvider().removeEventListener("contextmenu", this.handleContextMenu, true);
    document.removeEventListener("pointerdown", this.handleDocumentPointerDown, true);
    document.removeEventListener("keydown", this.handleDocumentKeyDown, true);
    window.removeEventListener("scroll", this.handleAnyScroll, true);
    this.hideContextMenu();
    this.contextMenu?.remove();
    this.contextMenu = null;
  }

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
    const row = this.getActiveTableRow();
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

  deleteTableRow(): void {
    const row = this.getActiveTableRow();
    if (!row || !this.context.rootProvider().contains(row)) return;

    const table = row.closest<HTMLTableElement>("table");
    const editable = row.closest<HTMLElement>("[contenteditable='true']");
    if (!table || !editable) return;

    const flowRows = this.getTableFlowRows(table);
    const adjacentRow =
      this.getSiblingTableRow(row.nextElementSibling) ??
      this.getSiblingTableRow(row.previousElementSibling);

    if (flowRows.length <= 1) {
      this.removeWholeTableFlow(table);
      this.context.markEdited(editable);
      return;
    }

    row.remove();
    if (table.rows.length === 0) {
      this.getTableFlowRoot(table).remove();
    }

    const nextRow = adjacentRow?.isConnected
      ? adjacentRow
      : this.getTableFlowRows(table)[0] ?? null;
    this.lastSelectedTableRow = nextRow;
    const cell = nextRow?.cells[0] as HTMLElement | undefined;
    if (cell) this.placeCaretInElement(cell);
    this.context.markEdited(editable);
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

  private getActiveTableRow(): HTMLTableRowElement | null {
    const row = this.getSelectionElement()?.closest("tr") ?? this.lastSelectedTableRow;
    return row instanceof HTMLTableRowElement && this.context.rootProvider().contains(row)
      ? row
      : null;
  }

  private onContextMenu(event: MouseEvent): void {
    const target = event.target;
    if (!(target instanceof Element)) return;

    const row = target.closest("tr");
    if (!(row instanceof HTMLTableRowElement) || !this.context.rootProvider().contains(row)) {
      this.hideContextMenu();
      return;
    }

    event.preventDefault();
    this.lastSelectedTableRow = row;
    this.showContextMenu(event.clientX, event.clientY);
  }

  private showContextMenu(clientX: number, clientY: number): void {
    const menu = this.ensureContextMenu();
    menu.style.display = "block";
    menu.style.left = "0px";
    menu.style.top = "0px";

    const rect = menu.getBoundingClientRect();
    const left = Math.max(8, Math.min(clientX, window.innerWidth - rect.width - 8));
    const top = Math.max(8, Math.min(clientY, window.innerHeight - rect.height - 8));
    menu.style.left = `${left}px`;
    menu.style.top = `${top}px`;
  }

  private hideContextMenu(): void {
    if (this.contextMenu) this.contextMenu.style.display = "none";
  }

  private ensureContextMenu(): HTMLDivElement {
    if (this.contextMenu) return this.contextMenu;

    const menu = document.createElement("div");
    menu.className = "hwe-table-context-menu";
    menu.setAttribute("role", "menu");
    menu.appendChild(
      this.makeContextMenuButton("Anadir fila debajo", () => this.insertTableRowAfter())
    );
    menu.appendChild(
      this.makeContextMenuButton("Eliminar fila", () => this.deleteTableRow())
    );
    document.body.appendChild(menu);
    this.contextMenu = menu;
    return menu;
  }

  private makeContextMenuButton(label: string, onAction: () => void): HTMLButtonElement {
    const button = document.createElement("button");
    button.type = "button";
    button.setAttribute("role", "menuitem");
    button.textContent = label;
    button.addEventListener("click", () => {
      this.hideContextMenu();
      onAction();
    });
    return button;
  }

  private getTableFlowRows(table: HTMLTableElement): HTMLTableRowElement[] {
    const flowId = table.getAttribute("data-hwe-table-flow-id");
    if (!flowId) return Array.from(table.rows);

    const rows: HTMLTableRowElement[] = [];
    Array.from(this.context.rootProvider().querySelectorAll<HTMLTableElement>("table"))
      .filter((candidate) => candidate.getAttribute("data-hwe-table-flow-id") === flowId)
      .forEach((candidate) => rows.push(...Array.from(candidate.rows)));
    return rows;
  }

  private removeWholeTableFlow(table: HTMLTableElement): void {
    const flowRoot = this.getTableFlowRoot(table);
    const blank = document.createElement("p");
    blank.appendChild(document.createElement("br"));
    flowRoot.parentNode?.insertBefore(blank, flowRoot.nextSibling);

    const flowId = table.getAttribute("data-hwe-table-flow-id");
    if (flowId) {
      Array.from(this.context.rootProvider().querySelectorAll<HTMLTableElement>("table"))
        .filter((tableInFlow) => tableInFlow.getAttribute("data-hwe-table-flow-id") === flowId)
        .forEach((tableInFlow) => this.getTableFlowRoot(tableInFlow).remove());
    } else {
      flowRoot.remove();
    }

    this.lastSelectedTableRow = null;
    this.placeCaretInElement(blank);
  }

  private getTableFlowRoot(table: HTMLTableElement): HTMLElement {
    return table.closest<HTMLElement>(".hwe-table-flow-wrapper") ?? table;
  }

  private getSiblingTableRow(element: Element | null): HTMLTableRowElement | null {
    return element instanceof HTMLTableRowElement ? element : null;
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
