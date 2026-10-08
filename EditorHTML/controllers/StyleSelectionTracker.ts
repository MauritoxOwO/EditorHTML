import { EDITOR_ROOT_CLASS } from "../dom/EditorCssScope";
import type { HistoryViewState } from "./HistoryViewState";

const PARAGRAPH_STYLE_BLOCK_SELECTOR = "p, div, li, h1, h2, h3, h4, h5, h6, blockquote, pre";
const EDITOR_CONTAINER_CLASS_NAMES = new Set([
  "hwe-page",
  "hwe-page-inner",
  EDITOR_ROOT_CLASS,
  "hwe-workspace",
  "hwe-table-flow-wrapper",
  "hwe-text-flow-column",
  "hwe-keep-together",
]);

export class StyleSelectionTracker {
  private lastTextSelection: StoredSelection | null = null;
  private lastBookmark: HistoryViewState | null = null;
  private lastStyleBlock: HTMLElement | null = null;

  constructor(
    private readonly rootProvider: () => HTMLElement,
    _activeEditableProvider: () => HTMLElement | null,
    private readonly documentSelection: { getSelectedParagraphs(range: Range): HTMLElement[] },
    private readonly bookmarkProvider?: {
      capture(workspace: HTMLElement): HistoryViewState;
      restore(workspace: HTMLElement, state: HistoryViewState): void;
    }
  ) {}

  rememberTextSelection(): void {
    const root = this.rootProvider();
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;

    const range = selection.getRangeAt(0);
    if (!this.isRangeInsideRoot(range, root) || !this.isEditableSelection(selection, root)) {
      this.clear();
      return;
    }

    const block =
      this.closestStyleBlockFromNode(range.startContainer) ??
      this.closestStyleBlockFromNode(selection.anchorNode) ??
      this.closestStyleBlockFromNode(selection.focusNode) ??
      this.closestStyleBlockFromNode(range.endContainer);
    this.lastTextSelection = {
      anchorNode: selection.anchorNode!,
      anchorOffset: selection.anchorOffset,
      focusNode: selection.focusNode!,
      focusOffset: selection.focusOffset,
      collapsed: selection.isCollapsed,
      backward: !selection.isCollapsed &&
        !(range.startContainer === selection.anchorNode && range.startOffset === selection.anchorOffset),
    };
    this.lastBookmark = this.bookmarkProvider?.capture(this.getBookmarkWorkspace(root)) ?? null;
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
    if (this.hasTextBookmark()) {
      this.bookmarkProvider!.restore(this.getBookmarkWorkspace(root), this.lastBookmark!);
      return;
    }
    const stored = this.getStoredSelection(root);
    if (!stored) return;

    const selection = window.getSelection();
    try {
      const range = document.createRange();
      selection?.removeAllRanges();
      if (stored.collapsed) {
        range.setStart(stored.anchorNode, stored.anchorOffset);
        range.collapse(true);
        selection?.addRange(range);
      } else if (stored.backward) {
        selection?.setBaseAndExtent(stored.anchorNode, stored.anchorOffset, stored.focusNode, stored.focusOffset);
      } else {
        range.setStart(stored.anchorNode, stored.anchorOffset);
        range.setEnd(stored.focusNode, stored.focusOffset);
        selection?.addRange(range);
      }
    } catch {
      this.lastTextSelection = null;
    }
  }

  clear(): void {
    this.lastTextSelection = null;
    this.lastBookmark = null;
    this.lastStyleBlock = null;
  }

  getSelectedStyleBlocks(): HTMLElement[] {
    const root = this.rootProvider();
    const range = this.getUsableRange(root);
    if (!range) return this.getLastStyleBlock(root);

    if (range.collapsed) {
      const block = this.closestStyleBlockFromNode(range.startContainer);
      if (block) this.lastStyleBlock = block;
      return block ? this.expandLogicalFlow([block], root) : [];
    }

    const blocks = this.documentSelection.getSelectedParagraphs(range)
      .filter((block) => this.isStyleBlock(block));

    if (blocks.length > 0) {
      // No aplicar tambien al padre: sus otros parrafos heredarian el estilo.
      const paragraphs = blocks.filter((block) =>
        !blocks.some((other) => other !== block && block.contains(other))
      );
      const logicalParagraphs = this.expandLogicalFlow(paragraphs, root);
      this.lastStyleBlock = logicalParagraphs[0];
      return logicalParagraphs;
    }

    const element =
      range.startContainer.nodeType === Node.ELEMENT_NODE
        ? (range.startContainer as HTMLElement)
        : range.startContainer.parentElement;
    const currentBlock = element ? this.closestStyleBlock(element) : null;
    if (currentBlock && root.contains(currentBlock)) {
      this.lastStyleBlock = currentBlock;
      return this.expandLogicalFlow([currentBlock], root);
    }

    return this.getLastStyleBlock(root);
  }

  private getUsableRange(root: HTMLElement): Range | null {
    const selection = window.getSelection();
    const activeRange =
      selection && selection.rangeCount > 0 ? selection.getRangeAt(0) : null;
    if (activeRange && this.isRangeInsideRoot(activeRange, root)) return activeRange;
    if (this.hasTextBookmark()) {
      this.bookmarkProvider!.restore(this.getBookmarkWorkspace(root), this.lastBookmark!);
      const selectionAfterRestore = window.getSelection();
      return selectionAfterRestore?.rangeCount ? selectionAfterRestore.getRangeAt(0) : null;
    }
    const stored = this.getStoredSelection(root);
    return stored ? this.createRange(stored) : null;
  }

  private getStoredSelection(root: HTMLElement): StoredSelection | null {
    const stored = this.lastTextSelection;
    if (!stored) return null;
    if (!root.contains(stored.anchorNode) || !root.contains(stored.focusNode) ||
        !this.isEditableNode(stored.anchorNode, root) || !this.isEditableNode(stored.focusNode, root)) {
      if (!this.hasTextBookmark()) this.clear();
      return null;
    }
    return stored;
  }

  private getBookmarkWorkspace(root: HTMLElement): HTMLElement {
    return root.querySelector<HTMLElement>(".hwe-workspace") ?? root;
  }

  private hasTextBookmark(): boolean {
    return Boolean(this.lastBookmark?.anchor && this.lastBookmark.focus && this.bookmarkProvider);
  }

  private createRange(stored: StoredSelection): Range {
    const range = document.createRange();
    if (stored.collapsed) {
      range.setStart(stored.anchorNode, stored.anchorOffset);
      range.collapse(true);
    } else if (stored.backward) {
      range.setStart(stored.focusNode, stored.focusOffset);
      range.setEnd(stored.anchorNode, stored.anchorOffset);
    } else {
      range.setStart(stored.anchorNode, stored.anchorOffset);
      range.setEnd(stored.focusNode, stored.focusOffset);
    }
    return range;
  }

  private isEditableNode(node: Node, root: HTMLElement): boolean {
    const element = node.nodeType === Node.ELEMENT_NODE ? node as HTMLElement : node.parentElement;
    return Boolean(element?.closest(".hwe-page-inner") && !element.closest(
      "[contenteditable='false'], [hidden], [data-hwe-api-header], [data-hwe-dynamic-header], [data-hwe-runtime-page-header]"
    ) && root.contains(element));
  }

  private getLastStyleBlock(root: HTMLElement): HTMLElement[] {
    return this.lastStyleBlock && root.contains(this.lastStyleBlock)
      ? [this.lastStyleBlock]
      : [];
  }

  private expandLogicalFlow(blocks: HTMLElement[], root: HTMLElement): HTMLElement[] {
    const flowIds = new Set(blocks.map((block) => block.getAttribute("data-hwe-text-flow-id")).filter((id): id is string => Boolean(id)));
    if (!flowIds.size) return blocks;
    const expanded = new Set(blocks);
    for (const block of Array.from(root.querySelectorAll<HTMLElement>(PARAGRAPH_STYLE_BLOCK_SELECTOR))) {
      if (!block.closest(".hwe-page-inner") || block.closest("[contenteditable='false'], [hidden], [data-hwe-api-header], [data-hwe-dynamic-header], [data-hwe-runtime-page-header]")) continue;
      if (!this.isStyleBlock(block)) continue;
      if (block.tagName === "DIV" && block.querySelector(PARAGRAPH_STYLE_BLOCK_SELECTOR)) continue;
      const id = block.getAttribute("data-hwe-text-flow-id");
      if (id && flowIds.has(id)) expanded.add(block);
    }
    return [...expanded];
  }

  private isRangeInsideRoot(range: Range, root: HTMLElement): boolean {
    return root.contains(range.startContainer) && root.contains(range.endContainer);
  }

  private isEditableSelection(selection: Selection, root: HTMLElement): boolean {
    return Boolean(
      this.closestEditableFromNode(selection.anchorNode, root) ||
        this.closestEditableFromNode(selection.focusNode, root)
    );
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
    if (block.closest("[contenteditable='false'], [data-hwe-api-header='true'], [data-hwe-dynamic-header='true']")) return false;
    if (Array.from(block.classList).some((className) => EDITOR_CONTAINER_CLASS_NAMES.has(className))) {
      return false;
    }

    if (block.tagName !== "DIV") return true;
    if (block.hasAttribute("contenteditable")) return false;
    if (block.matches("[data-hwe-page-break], [data-hwe-manual-page-break], [data-hwe-generated-wrapper], [data-hwe-keep-together]")) {
      return false;
    }

    return true;
  }
}

interface StoredSelection {
  anchorNode: Node;
  anchorOffset: number;
  focusNode: Node;
  focusOffset: number;
  collapsed: boolean;
  backward?: boolean;
}
