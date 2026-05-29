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
    if (!root.contains(range.commonAncestorContainer)) return;

    const element =
      range.commonAncestorContainer.nodeType === Node.ELEMENT_NODE
        ? (range.commonAncestorContainer as HTMLElement)
        : range.commonAncestorContainer.parentElement;
    if (!element?.closest("[contenteditable='true']")) return;

    const block = this.closestStyleBlock(element);
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
    if (!this.lastTextSelection) return;

    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(this.lastTextSelection);
  }

  getSelectedStyleBlocks(): HTMLElement[] {
    const root = this.rootProvider();
    const selection = window.getSelection();
    const range = selection && selection.rangeCount > 0 ? selection.getRangeAt(0) : null;
    if (!range || !root.contains(range.commonAncestorContainer)) {
      return this.lastStyleBlock && root.contains(this.lastStyleBlock)
        ? [this.lastStyleBlock]
        : [];
    }

    const editable = this.activeEditableProvider();
    const scope = editable ?? root;
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

    return this.lastStyleBlock && root.contains(this.lastStyleBlock)
      ? [this.lastStyleBlock]
      : [];
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

  private isStyleBlock(block: HTMLElement): boolean {
    if (block.closest("[data-hwe-dynamic-header='true']")) return false;
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
