const TABLE_STRUCTURE_SELECTOR = "table, thead, tbody, tfoot, tr, colgroup, col";
const TABLE_CELL_SELECTOR = "td, th";
const MANUAL_PAGE_BREAK_SELECTOR = "[data-hwe-manual-page-break='true']";

interface InvalidTableContext {
  table: HTMLTableElement;
  cell: HTMLElement | null;
}

export class TableDomIntegrityController {
  handleBeforeInput(event: InputEvent, root: HTMLElement): boolean {
    if (!this.shouldGuardInput(event.inputType)) return false;

    const context = this.getInvalidTableContext(root, event.target);
    if (!context) return false;

    event.preventDefault();
    if (!context.cell) return true;

    this.placeCaretAtEnd(context.cell);

    if (event.inputType === "insertParagraph") {
      document.execCommand("insertParagraph");
    } else if (event.inputType === "insertLineBreak") {
      document.execCommand("insertLineBreak");
    } else if (event.data) {
      this.insertPlainText(event.data);
    }

    return true;
  }

  handlePaste(event: ClipboardEvent, root: HTMLElement): boolean {
    const context = this.getInvalidTableContext(root, event.target);
    if (!context) return false;

    event.preventDefault();
    if (!context.cell) return true;

    this.placeCaretAtEnd(context.cell);
    this.insertPlainText(
      event.clipboardData?.getData("text/plain") ||
        this.htmlToPlainText(event.clipboardData?.getData("text/html") ?? "")
    );
    return true;
  }

  handleDrop(event: DragEvent, root: HTMLElement): boolean {
    const context = this.getInvalidTableContext(root, event.target);
    if (!context) return false;

    event.preventDefault();
    if (!context.cell) return true;

    this.placeCaretAtEnd(context.cell);
    this.insertPlainText(event.dataTransfer?.getData("text/plain") ?? "");
    return true;
  }

  normalize(root: HTMLElement): boolean {
    let changed = false;
    const containers = Array.from(
      root.querySelectorAll<HTMLElement>("table, thead, tbody, tfoot, tr, colgroup")
    );

    containers.forEach((container) => {
      Array.from(container.childNodes).forEach((child) => {
        if (this.isAllowedTableChild(container, child)) return;

        if (this.isWhitespaceText(child)) {
          child.remove();
          changed = true;
          return;
        }

        if (
          child.nodeType === Node.ELEMENT_NODE &&
          (child as HTMLElement).matches(MANUAL_PAGE_BREAK_SELECTOR)
        ) {
          changed = this.moveManualBreakAfterTable(child as HTMLElement, container) || changed;
          return;
        }

        changed = this.moveInvalidNodeIntoCell(child, container) || changed;
      });
    });

    return changed;
  }

  private shouldGuardInput(inputType: string): boolean {
    return (
      inputType.startsWith("insert") ||
      inputType.startsWith("delete") ||
      inputType === "historyUndo" ||
      inputType === "historyRedo"
    );
  }

  private getInvalidTableContext(
    root: HTMLElement,
    eventTarget: EventTarget | null
  ): InvalidTableContext | null {
    const selectionElement = this.getSelectionElement(root);
    if (selectionElement?.closest(TABLE_CELL_SELECTOR)) return null;

    const structureElement =
      this.getTableStructureElement(selectionElement, root) ??
      this.getTableStructureElement(this.eventTargetToElement(eventTarget), root);
    if (!structureElement) return null;

    const table =
      structureElement.tagName === "TABLE"
        ? (structureElement as HTMLTableElement)
        : structureElement.closest<HTMLTableElement>("table");
    if (!table || !root.contains(table)) return null;

    return {
      table,
      cell: this.findNearestCell(structureElement, table),
    };
  }

  private getSelectionElement(root: HTMLElement): HTMLElement | null {
    const selection = window.getSelection();
    const node = selection?.anchorNode ?? null;
    const element = this.nodeToElement(node);
    return element && root.contains(element) ? element : null;
  }

  private getTableStructureElement(
    element: HTMLElement | null,
    root: HTMLElement
  ): HTMLElement | null {
    if (!element || !root.contains(element)) return null;
    if (element.closest(TABLE_CELL_SELECTOR)) return null;
    return element.closest<HTMLElement>(TABLE_STRUCTURE_SELECTOR);
  }

  private eventTargetToElement(target: EventTarget | null): HTMLElement | null {
    return target instanceof HTMLElement ? target : null;
  }

  private nodeToElement(node: Node | null): HTMLElement | null {
    if (!node) return null;
    return node.nodeType === Node.ELEMENT_NODE
      ? (node as HTMLElement)
      : node.parentElement;
  }

  private findNearestCell(origin: HTMLElement, table: HTMLTableElement): HTMLElement | null {
    const row = origin.closest("tr");
    const rowCell = row?.querySelector<HTMLElement>(TABLE_CELL_SELECTOR) ?? null;
    if (rowCell) return rowCell;

    return this.findSiblingCell(origin, true) ?? this.findSiblingCell(origin, false) ??
      table.querySelector<HTMLElement>(TABLE_CELL_SELECTOR);
  }

  private findSiblingCell(origin: ChildNode, previous: boolean): HTMLElement | null {
    let sibling = previous ? origin.previousSibling : origin.nextSibling;
    while (sibling) {
      const cell = this.findCellInNode(sibling, previous);
      if (cell) return cell;
      sibling = previous ? sibling.previousSibling : sibling.nextSibling;
    }

    return null;
  }

  private findCellInNode(node: ChildNode, last = false): HTMLElement | null {
    if (node.nodeType !== Node.ELEMENT_NODE) return null;

    const element = node as HTMLElement;
    if (element.matches(TABLE_CELL_SELECTOR)) return element;

    const cells = Array.from(element.querySelectorAll<HTMLElement>(TABLE_CELL_SELECTOR));
    return last ? cells[cells.length - 1] ?? null : cells[0] ?? null;
  }

  private placeCaretAtEnd(cell: HTMLElement): void {
    if (!cell.firstChild) cell.appendChild(document.createElement("br"));

    const editable = cell.closest<HTMLElement>("[contenteditable='true']");
    editable?.focus({ preventScroll: true });

    const selection = window.getSelection();
    if (!selection) return;

    const range = document.createRange();
    range.selectNodeContents(cell);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
  }

  private insertPlainText(text: string): void {
    if (!text) return;
    if (!document.execCommand("insertText", false, text)) {
      const selection = window.getSelection();
      const range = selection?.rangeCount ? selection.getRangeAt(0) : null;
      if (!range) return;

      range.deleteContents();
      range.insertNode(document.createTextNode(text));
      range.collapse(false);
      selection?.removeAllRanges();
      selection?.addRange(range);
    }
  }

  private htmlToPlainText(html: string): string {
    if (!html) return "";
    const template = document.createElement("template");
    template.innerHTML = html;
    return template.content.textContent ?? "";
  }

  private isAllowedTableChild(container: HTMLElement, child: ChildNode): boolean {
    if (child.nodeType === Node.TEXT_NODE) return this.isWhitespaceText(child);
    if (child.nodeType !== Node.ELEMENT_NODE) return false;

    const tagName = (child as HTMLElement).tagName;
    switch (container.tagName) {
      case "TABLE":
        return ["CAPTION", "COLGROUP", "THEAD", "TBODY", "TFOOT", "TR"].includes(tagName);
      case "THEAD":
      case "TBODY":
      case "TFOOT":
        return tagName === "TR";
      case "TR":
        return tagName === "TD" || tagName === "TH";
      case "COLGROUP":
        return tagName === "COL";
      default:
        return true;
    }
  }

  private isWhitespaceText(node: ChildNode): boolean {
    return (
      node.nodeType === Node.TEXT_NODE &&
      (node.textContent ?? "").replace(/\u00a0/g, " ").trim() === ""
    );
  }

  private moveManualBreakAfterTable(marker: HTMLElement, container: HTMLElement): boolean {
    const table =
      container.tagName === "TABLE"
        ? (container as HTMLTableElement)
        : container.closest<HTMLTableElement>("table");
    const flowRoot = table?.closest<HTMLElement>(".hwe-table-flow-wrapper") ?? table;
    const parent = flowRoot?.parentNode;
    if (!flowRoot || !parent) return false;

    parent.insertBefore(marker, flowRoot.nextSibling);
    return true;
  }

  private moveInvalidNodeIntoCell(node: ChildNode, container: HTMLElement): boolean {
    const table =
      container.tagName === "TABLE"
        ? (container as HTMLTableElement)
        : container.closest<HTMLTableElement>("table");
    const cell =
      this.findSiblingCell(node, true) ??
      this.findSiblingCell(node, false) ??
      container.closest("tr")?.querySelector<HTMLElement>(TABLE_CELL_SELECTOR) ??
      table?.querySelector<HTMLElement>(TABLE_CELL_SELECTOR) ??
      null;

    if (!cell) {
      node.remove();
      return true;
    }

    if (node.nodeType === Node.TEXT_NODE) {
      const paragraph = document.createElement("p");
      paragraph.appendChild(node);
      cell.appendChild(paragraph);
      return true;
    }

    cell.appendChild(node);
    return true;
  }
}
