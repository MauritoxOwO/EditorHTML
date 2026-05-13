const PARAGRAPH_STYLE_BLOCK_SELECTOR = "p, li, h1, h2, h3, h4, h5, h6, blockquote, pre";

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

    this.lastTextSelection = range.cloneRange();
    this.lastStyleBlock = element.closest<HTMLElement>(PARAGRAPH_STYLE_BLOCK_SELECTOR);
  }

  rememberStyleBlockFromEvent(event: Event): void {
    const root = this.rootProvider();
    const target = event.target as HTMLElement | null;
    const block = target?.closest<HTMLElement>(PARAGRAPH_STYLE_BLOCK_SELECTOR);
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
    ).filter((block) => this.rangeOverlapsBlock(range, block));

    if (blocks.length > 0) {
      this.lastStyleBlock = blocks[0];
      return blocks;
    }

    const element =
      range.startContainer.nodeType === Node.ELEMENT_NODE
        ? (range.startContainer as HTMLElement)
        : range.startContainer.parentElement;
    const currentBlock = element?.closest<HTMLElement>(PARAGRAPH_STYLE_BLOCK_SELECTOR);
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
}
