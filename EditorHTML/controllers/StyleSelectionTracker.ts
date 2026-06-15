const PARAGRAPH_STYLE_BLOCK_SELECTOR = "p, div, li, h1, h2, h3, h4, h5, h6, blockquote, pre";
const EDITOR_CONTAINER_CLASS_NAMES = new Set([
  "hwe-page",
  "hwe-page-inner",
  "hwe-root",
  "hwe-workspace",
  "hwe-table-flow-wrapper",
  "hwe-keep-together",
]);

export class StyleSelectionTracker {
  private lastTextSelection: Range | null = null;
  private lastStyleBlock: HTMLElement | null = null;

  constructor(
    private readonly rootProvider: () => HTMLElement,
    private readonly activeEditableProvider: () => HTMLElement | null
  ) {}

  rememberTextSelection(): void {
    const root = this.rootProvider();
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;

    const range = selection.getRangeAt(0);
    if (!this.isRangeInsideRoot(range, root)) return;
    if (!this.isEditableSelection(selection, root)) return;

    const block =
      this.closestStyleBlockFromNode(range.startContainer) ??
      this.closestStyleBlockFromNode(selection.anchorNode) ??
      this.closestStyleBlockFromNode(selection.focusNode) ??
      this.closestStyleBlockFromNode(range.endContainer);
    this.lastTextSelection = range.cloneRange();
    if (block) this.lastStyleBlock = block;
  }

  rememberStyleBlockFromEvent(event: Event): void {
    const root = this.rootProvider();
    const target = event.target as HTMLElement | null;
    const block = target ? this.closestStyleBlock(target) : null;
    if (block && root.contains(block)) this.lastStyleBlock = block;
  }

  restoreTextSelection(): void {
    const root = this.rootProvider();
    const range = this.getStoredRange(root);
    if (!range) return;

    const selection = window.getSelection();
    try {
      selection?.removeAllRanges();
      selection?.addRange(range);
    } catch {
      this.lastTextSelection = null;
    }
  }

  getSelectedStyleBlocks(): HTMLElement[] {
    const root = this.rootProvider();
    const range = this.getUsableRange(root);
    if (!range) return this.getLastStyleBlock(root);

    const editable = this.closestEditableFromNode(range.commonAncestorContainer, root);
    const scope = range.collapsed ? this.activeEditableProvider() ?? editable ?? root : editable ?? root;
    const blocks = Array.from(
      scope.querySelectorAll<HTMLElement>(PARAGRAPH_STYLE_BLOCK_SELECTOR)
    ).filter((block) => this.isStyleBlock(block) && this.rangeOverlapsBlock(range, block));

    if (blocks.length > 0) {
      this.lastStyleBlock = blocks[0];
      return blocks;
    }

    const element =
      range.startContainer.nodeType === Node.ELEMENT_NODE
        ? (range.startContainer as HTMLElement)
        : range.startContainer.parentElement;
    const currentBlock = element ? this.closestStyleBlock(element) : null;
    if (currentBlock && root.contains(currentBlock)) {
      this.lastStyleBlock = currentBlock;
      return [currentBlock];
    }

    return this.getLastStyleBlock(root);
  }

  private getUsableRange(root: HTMLElement): Range | null {
    const selection = window.getSelection();
    const activeRange =
      selection && selection.rangeCount > 0 ? selection.getRangeAt(0) : null;
    const storedRange = this.getStoredRange(root);

    if (!activeRange || !this.isRangeInsideRoot(activeRange, root)) return storedRange;
    if (activeRange.collapsed && storedRange && !storedRange.collapsed) return storedRange;

    return activeRange;
  }

  private getStoredRange(root: HTMLElement): Range | null {
    if (!this.lastTextSelection) return null;
    return this.isRangeInsideRoot(this.lastTextSelection, root) ? this.lastTextSelection : null;
  }

  private getLastStyleBlock(root: HTMLElement): HTMLElement[] {
    return this.lastStyleBlock && root.contains(this.lastStyleBlock)
      ? [this.lastStyleBlock]
      : [];
  }

  private isRangeInsideRoot(range: Range, root: HTMLElement): boolean {
    return root.contains(range.commonAncestorContainer);
  }

  private isEditableSelection(selection: Selection, root: HTMLElement): boolean {
    return Boolean(
      this.closestEditableFromNode(selection.anchorNode, root) ||
        this.closestEditableFromNode(selection.focusNode, root)
    );
  }

  private rangeOverlapsBlock(range: Range, block: HTMLElement): boolean {
    if (range.collapsed) return block.contains(range.startContainer);

    const blockRange = document.createRange();
    blockRange.selectNodeContents(block);

    const startsBeforeBlockEnds =
      range.compareBoundaryPoints(Range.START_TO_END, blockRange) < 0;
    const endsAfterBlockStarts =
      range.compareBoundaryPoints(Range.END_TO_START, blockRange) > 0;
    return startsBeforeBlockEnds && endsAfterBlockStarts;
  }

  private closestStyleBlock(element: HTMLElement): HTMLElement | null {
    const root = this.rootProvider();
    let current: HTMLElement | null = element;

    while (current && root.contains(current)) {
      if (current.matches(PARAGRAPH_STYLE_BLOCK_SELECTOR) && this.isStyleBlock(current)) {
        return current;
      }
      current = current.parentElement;
    }

    return null;
  }

  private closestStyleBlockFromNode(node: Node | null): HTMLElement | null {
    if (!node) return null;
    const element =
      node.nodeType === Node.ELEMENT_NODE ? (node as HTMLElement) : node.parentElement;
    return element ? this.closestStyleBlock(element) : null;
  }

  private closestEditableFromNode(node: Node | null, root: HTMLElement): HTMLElement | null {
    if (!node) return null;
    const element =
      node.nodeType === Node.ELEMENT_NODE ? (node as HTMLElement) : node.parentElement;
    const editable = element?.closest<HTMLElement>("[contenteditable='true']") ?? null;
    return editable && root.contains(editable) ? editable : null;
  }

  private isStyleBlock(block: HTMLElement): boolean {
    if (block.closest("[data-hwe-api-header='true'], [data-hwe-dynamic-header='true']")) return false;
    if (Array.from(block.classList).some((className) => EDITOR_CONTAINER_CLASS_NAMES.has(className))) {
      return false;
    }

    if (block.tagName !== "DIV") return true;
    if (block.hasAttribute("contenteditable")) return false;
    if (Array.from(block.attributes).some((attr) => attr.name.startsWith("data-hwe-"))) {
      return false;
    }

    return true;
  }
}
