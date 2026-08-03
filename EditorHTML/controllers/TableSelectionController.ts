import { getTableFlowFragments } from "../dom/TableFlow";

export interface TableSelectionControllerOptions {
  rootProvider: () => HTMLElement | null;
  onTableChanged: (table: HTMLTableElement) => void;
}

const SUPPORTED_COMMANDS = new Set(["bold", "underline"]);

export class TableSelectionController {
  private handle: HTMLButtonElement | null = null;
  private hoveredTable: HTMLTableElement | null = null;
  private selectedTable: HTMLTableElement | null = null;
  private isStarted = false;

  private readonly handleRootPointerMove = (event: PointerEvent): void =>
    this.onRootPointerMove(event);
  private readonly handleRootPointerLeave = (): void => {
    if (!this.selectedTable) this.hideHandle();
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
    this.isStarted = false;
    this.clearSelection();
    this.handle?.remove();
    this.handle = null;
  }

  handleToolbarCommand(command: string): boolean {
    if (!SUPPORTED_COMMANDS.has(command)) return false;

    const tables = this.getSelectedTables();
    if (tables.length === 0) return false;

    if (command === "bold") {
      const textContainers = tables.flatMap((table) => this.getTextContainers(table));
      const isActive =
        textContainers.length > 0
          ? textContainers.every((element) => this.isBold(element))
          : tables.every((table) => this.isBold(table));
      tables.forEach((table) => {
        table.style.fontWeight = isActive ? "normal" : "700";
      });
      textContainers.forEach((element) => {
        element.style.fontWeight = isActive ? "normal" : "700";
      });
    } else {
      const textContainers = tables.flatMap((table) => this.getTextContainers(table));
      const isActive =
        textContainers.length > 0
          ? textContainers.every((element) => this.isUnderlined(element))
          : tables.every((table) => this.isUnderlined(table));
      tables.forEach((table) => {
        table.style.textDecoration = isActive ? "none" : "underline";
      });
      textContainers.forEach((element) =>
        this.setUnderline(element, isActive ? "none" : "underline")
      );
    }

    this.notifyTableChanged(tables[0]);
    return true;
  }

  applyFontSize(fontSize: string): boolean {
    if (!fontSize) return false;

    const tables = this.getSelectedTables();
    if (tables.length === 0) return false;

    tables.forEach((table) => {
      table.style.fontSize = fontSize;
      table.setAttribute("data-hwe-user-font-size", "true");
      this.getTextContainers(table).forEach((element) => {
        element.style.fontSize = fontSize;
        element.setAttribute("data-hwe-user-font-size", "true");
      });
    });
    this.notifyTableChanged(tables[0]);
    return true;
  }

  refresh(): void {
    const table = this.getVisibleHandleTable();
    if (!table) {
      this.hideHandle();
      return;
    }
    this.positionHandle(table);
  }

  clearSelection(): void {
    this.getSelectedTables().forEach((table) => table.classList.remove("hwe-table-selected"));
    this.selectedTable = null;
    this.hoveredTable = null;
    this.hideHandle();
  }

  private onRootPointerMove(event: PointerEvent): void {
    const target = event.target as Element | null;
    if (!target || this.handle?.contains(target)) return;

    const table = target.closest<HTMLTableElement>(".hwe-page-inner table");
    const root = this.options.rootProvider();
    if (!table || !root?.contains(table)) {
      this.hoveredTable = null;
      const selectedTable = this.getSelectedTables()[0];
      if (selectedTable) this.positionHandle(selectedTable);
      else this.hideHandle();
      return;
    }

    this.hoveredTable = table;
    this.positionHandle(table);
  }

  private onDocumentPointerDown(event: PointerEvent): void {
    const target = event.target as Element | null;
    if (!target || this.handle?.contains(target)) return;
    if (target.closest(".hwe-editor-header")) return;

    const selectedTables = this.getSelectedTables();
    if (selectedTables.some((table) => table.contains(target))) return;
    this.clearSelection();
  }

  private selectTable(table: HTMLTableElement): void {
    this.getSelectedTables().forEach((fragment) =>
      fragment.classList.remove("hwe-table-selected")
    );
    this.selectedTable = table;
    this.getSelectedTables().forEach((fragment) => fragment.classList.add("hwe-table-selected"));
    this.positionHandle(table);
  }

  private getSelectedTables(): HTMLTableElement[] {
    if (!this.selectedTable?.isConnected) return [];
    return getTableFlowFragments(this.options.rootProvider(), this.selectedTable);
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
    handle.addEventListener("pointerdown", (event) => {
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

  private hideHandle(): void {
    if (this.handle) this.handle.style.display = "none";
  }

  private getTextContainers(table: HTMLTableElement): HTMLElement[] {
    const containers = new Set<HTMLElement>();
    const walker = document.createTreeWalker(table, NodeFilter.SHOW_TEXT, {
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

  private setUnderline(element: HTMLElement, value: "none" | "underline"): void {
    let current: HTMLElement | null = element;
    while (current) {
      current.style.textDecoration = value;
      if (value === "underline" || current instanceof HTMLTableElement) break;
      current = current.parentElement;
    }
  }

  private notifyTableChanged(table: HTMLTableElement): void {
    this.options.onTableChanged(table);
    window.requestAnimationFrame(() => this.refresh());
  }
}
