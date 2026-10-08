import {
  getMeaningfulChildren,
  isEditableBlankBlock,
  isEmptyNode,
} from "../dom/EditableDom";

export interface PageBackspaceContext {
  pages: HTMLElement[];
  onContentChanged: (page: HTMLElement) => void;
}

export class PageBackspaceController {
  handleBoundaryDelete(
    direction: "backward" | "forward",
    range: Range,
    context: PageBackspaceContext
  ): HTMLElement | null {
    const inner = this.getInnerForPoint(range.startContainer);
    const page = inner?.closest<HTMLElement>(".hwe-page");
    if (!inner || !page) return null;
    const index = context.pages.indexOf(page);
    const currentAtBoundary = direction === "backward"
      ? this.isCaretAtStartOfEditable(inner)
      : this.isCaretAtEndOfEditable(inner);
    const point = range.startContainer.nodeType === Node.ELEMENT_NODE
      ? range.startContainer as HTMLElement : range.startContainer.parentElement;
    const block = point?.closest<HTMLElement>("p");
    const column = block?.parentElement;
    const edge = range.cloneRange();
    if (block) {
      edge.selectNodeContents(block);
      if (direction === "backward") edge.setEnd(range.startContainer, range.startOffset);
      else edge.setStart(range.startContainer, range.startOffset);
    }
    const sibling = direction === "backward" ? block?.previousElementSibling : block?.nextElementSibling;
    const nextColumn = direction === "backward" ? column?.previousElementSibling : column?.nextElementSibling;
    const localBlock = sibling ?? (nextColumn?.classList.contains("hwe-text-flow-column")
      ? direction === "backward" ? nextColumn.lastElementChild : nextColumn.firstElementChild : null);
    const localDelete = !!block && column?.classList.contains("hwe-text-flow-column") &&
      !this.fragmentHasVisibleContent(edge.cloneContents()) && localBlock?.matches("p");
    if (!currentAtBoundary && !localDelete) return null;
    const adjacent = localDelete ? page : direction === "backward" ? context.pages[index - 1] : context.pages[index + 1];
    const adjacentInner = adjacent?.querySelector<HTMLElement>(".hwe-page-inner");
    if (!adjacentInner) return null;
    const currentBlock = localDelete ? block : this.getBoundaryBlock(inner, direction === "backward" ? "first" : "last");
    const adjacentBlock = localDelete ? localBlock as HTMLElement : this.getBoundaryBlock(adjacentInner, direction === "backward" ? "last" : "first");
    if (!currentBlock || !adjacentBlock) {
      this.placeCaretAtEnd(direction === "backward" ? adjacentInner : inner);
      return direction === "backward" ? adjacent : page;
    }

    if (currentBlock.closest("td, th") || adjacentBlock.closest("td, th") ||
        !this.areMergeableBlocks(currentBlock, adjacentBlock)) {
      this.placeCaretAtEnd(direction === "backward" ? adjacentInner : inner);
      return direction === "backward" ? adjacent : page;
    }

    let before = direction === "backward" ? adjacentBlock : currentBlock;
    const after = direction === "backward" ? currentBlock : adjacentBlock;
    const afterIsBlank = isEditableBlankBlock(after);
    const flowId = before.getAttribute("data-hwe-text-flow-id");
    const sameFlow = !!flowId && flowId === after.getAttribute("data-hwe-text-flow-id");
    let deletionPoint: { node: Text; offset: number } | null = null;
    if (sameFlow) {
      const walker = document.createTreeWalker(direction === "backward" ? before : after, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const text = walker.currentNode as Text;
        if (!text.length) continue;
        deletionPoint = { node: text, offset: direction === "backward" ? text.length : 0 };
        if (direction === "forward") break;
      }
    }
    if (deletionPoint?.node?.nodeType === Node.TEXT_NODE) {
      const text = deletionPoint.node as Text;
      const characters = Array.from(text.data);
      const length = (direction === "backward" ? characters.pop() : characters[0])?.length ?? 0;
      if (direction === "backward") deletionPoint.offset -= length;
      text.deleteData(deletionPoint.offset, length);
    }
    if (!afterIsBlank && isEditableBlankBlock(before)) {
      before.replaceWith(after);
      before = after;
    }
    const insertionOffset = before.childNodes.length;
    if (!afterIsBlank && before !== after) {
      const beforeId = before.getAttribute("data-hwe-text-flow-id") ?? after.getAttribute("data-hwe-text-flow-id");
      const afterId = after.getAttribute("data-hwe-text-flow-id");
      if (beforeId) before.setAttribute("data-hwe-text-flow-id", beforeId);
      if (beforeId && afterId && beforeId !== afterId) {
        context.pages.forEach((item) => item.querySelectorAll<HTMLElement>("[data-hwe-text-flow-id]").forEach((fragment) => {
          if (fragment.getAttribute("data-hwe-text-flow-id") === afterId) fragment.setAttribute("data-hwe-text-flow-id", beforeId);
        }));
      }
      while (after.firstChild) before.appendChild(after.firstChild);
      before.removeAttribute("data-hwe-user-blank");
    }
    if (before !== after) after.remove();
    const selection = window.getSelection();
    const insertionPoint = document.createRange();
    if (deletionPoint?.node?.nodeType === Node.TEXT_NODE) {
      insertionPoint.setStart(deletionPoint.node, deletionPoint.offset);
    } else insertionPoint.setStart(before, before === after ? 0 : insertionOffset);
    insertionPoint.collapse(true);
    selection?.removeAllRanges();
    selection?.addRange(insertionPoint);
    return direction === "backward" ? adjacent : page;
  }

  handleBackspaceAtPageStart(event: KeyboardEvent, context: PageBackspaceContext): boolean {
    const selection = window.getSelection();
    const selectionNode = selection?.focusNode;
    const selectionElement = selectionNode?.nodeType === Node.ELEMENT_NODE
      ? selectionNode as HTMLElement
      : selectionNode?.parentElement;
    const inner = selectionElement?.closest<HTMLElement>(".hwe-page-inner") ??
      ((event.currentTarget as HTMLElement).matches(".hwe-page-inner")
        ? event.currentTarget as HTMLElement
        : null);
    if (!inner) return false;
    const page = inner.closest<HTMLElement>(".hwe-page");
    if (!page) return false;

    const caretAtPageStart = this.isCaretAtStartOfEditable(inner);
    const caretAtTableContinuationStart =
      !caretAtPageStart && this.isCaretAtStartOfFirstTableFragment(inner);
    if (!caretAtPageStart && !caretAtTableContinuationStart) return false;

    const pageIndex = context.pages.indexOf(page);
    if (pageIndex <= 0) return false;

    const previousPage = context.pages[pageIndex - 1];
    const previousInner = previousPage.querySelector<HTMLElement>(".hwe-page-inner");
    if (!previousInner) return false;

    event.preventDefault();
    if (caretAtTableContinuationStart) {
      this.removeTrailingPaginationBlanks(previousInner);
    } else {
      this.removeLeadingPaginationBlanks(inner);
    }
    this.placeCaretAtEnd(previousInner);
    context.onContentChanged(previousPage);
    return true;
  }

  private isCaretAtStartOfFirstTableFragment(inner: HTMLElement): boolean {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0 || !selection.isCollapsed) return false;

    const range = selection.getRangeAt(0);
    if (!inner.contains(range.startContainer)) return false;

    const first = getMeaningfulChildren(inner)[0];
    if (!first || first.nodeType !== Node.ELEMENT_NODE) return false;

    const firstElement = first as HTMLElement;
    if (firstElement.tagName !== "TABLE" || !firstElement.contains(range.startContainer)) {
      return false;
    }

    const beforeRange = document.createRange();
    beforeRange.selectNodeContents(firstElement);
    beforeRange.setEnd(range.startContainer, range.startOffset);

    return !this.fragmentHasTextOrMediaContent(beforeRange.cloneContents());
  }

  private fragmentHasTextOrMediaContent(fragment: DocumentFragment): boolean {
    return Array.from(fragment.childNodes).some((node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        return (node.textContent ?? "").replace(/\u00a0/g, " ").trim().length > 0;
      }

      if (node.nodeType !== Node.ELEMENT_NODE) return false;

      const element = node as HTMLElement;
      if (element.tagName === "BR") return false;
      if (element.matches("img, hr, video, canvas, svg, [data-hwe-manual-page-break='true'], [contenteditable='false'], [hidden]")) return true;
      if (element.querySelector("img, hr, video, canvas, svg, [data-hwe-manual-page-break='true']")) return true;
      return (element.textContent ?? "").replace(/\u00a0/g, " ").trim().length > 0;
    });
  }

  private removeTrailingPaginationBlanks(container: HTMLElement): boolean {
    let removedAny = false;
    let child = container.lastChild;

    while (child) {
      const previous = child.previousSibling;
      if (!this.isRemovablePaginationBlank(child)) break;

      child.remove();
      removedAny = true;
      child = previous;
    }

    return removedAny;
  }

  private removeLeadingPaginationBlanks(container: HTMLElement): boolean {
    let removedAny = false;
    let child = container.firstChild;

    while (child) {
      const next = child.nextSibling;
      if (!this.isRemovablePaginationBlank(child)) break;

      child.remove();
      removedAny = true;
      child = next;
    }

    return removedAny;
  }

  private isRemovablePaginationBlank(node: ChildNode): boolean {
    if (node.nodeType === Node.TEXT_NODE) {
      return (node.textContent ?? "").replace(/\u00a0/g, " ").trim() === "";
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return false;

    const element = node as HTMLElement;
    if (element.hasAttribute("data-hwe-user-blank")) return false;
    return isEmptyNode(element, false) || isEditableBlankBlock(element);
  }

  private isCaretAtStartOfEditable(inner: HTMLElement): boolean {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0 || !selection.isCollapsed) return false;

    const range = selection.getRangeAt(0);
    if (!inner.contains(range.startContainer)) return false;

    const beforeRange = document.createRange();
    beforeRange.selectNodeContents(inner);
    beforeRange.setEnd(range.startContainer, range.startOffset);

    const fragment = beforeRange.cloneContents();
    return !this.fragmentHasVisibleContent(fragment);
  }

  private isCaretAtEndOfEditable(inner: HTMLElement): boolean {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0 || !selection.isCollapsed) return false;
    const range = selection.getRangeAt(0);
    if (!inner.contains(range.startContainer)) return false;
    const afterRange = document.createRange();
    afterRange.selectNodeContents(inner);
    afterRange.setStart(range.startContainer, range.startOffset);
    return !this.fragmentHasVisibleContent(afterRange.cloneContents());
  }

  private getBoundaryBlock(inner: HTMLElement, direction: "first" | "last"): HTMLElement | null {
    const blocks = Array.from(inner.querySelectorAll<HTMLElement>("p, li, h1, h2, h3, h4, h5, h6, blockquote, pre"))
      .filter((block) => !block.closest("[contenteditable='false'], [hidden], [data-hwe-api-header], [data-hwe-dynamic-header]"));
    const block = direction === "first" ? blocks[0] : blocks[blocks.length - 1];
    if (!block) return null;
    const barriers = Array.from(inner.querySelectorAll<HTMLElement>(
      "img, table, hr, [data-hwe-manual-page-break='true'], [contenteditable='false'], [hidden]"
    ));
    if (barriers.some((barrier) => {
      const relation = barrier.compareDocumentPosition(block);
      return direction === "first"
        ? !!(relation & Node.DOCUMENT_POSITION_FOLLOWING)
        : !!(relation & Node.DOCUMENT_POSITION_PRECEDING);
    })) return null;
    return block;
  }

  private areMergeableBlocks(left: HTMLElement, right: HTMLElement): boolean {
    return left.tagName === right.tagName &&
      left.getAttribute("class") === right.getAttribute("class") &&
      left.getAttribute("style") === right.getAttribute("style") &&
      left.matches("p, li, h1, h2, h3, h4, h5, h6, blockquote, pre") &&
      right.matches("p, li, h1, h2, h3, h4, h5, h6, blockquote, pre") &&
      (left.tagName !== "LI" || left.parentElement === right.parentElement);
  }

  private getInnerForPoint(node: Node): HTMLElement | null {
    const element = node.nodeType === Node.ELEMENT_NODE ? node as HTMLElement : node.parentElement;
    return element?.closest<HTMLElement>(".hwe-page-inner") ?? null;
  }

  private fragmentHasVisibleContent(fragment: DocumentFragment): boolean {
    return Array.from(fragment.childNodes).some((node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        return (node.textContent ?? "").replace(/\u00a0/g, " ").length > 0;
      }

      if (node.nodeType !== Node.ELEMENT_NODE) return false;

      const element = node as HTMLElement;
      if (element.tagName === "BR") return true;
      if (element.matches("img, table, hr, [data-hwe-manual-page-break='true'], [contenteditable='false'], [hidden]")) return true;
      if (element.querySelector("br, img, table, tr, td, th, hr, video, canvas, svg, [data-hwe-manual-page-break='true']")) return true;
      return (element.textContent ?? "").replace(/\u00a0/g, " ").length > 0;
    });
  }

  private placeCaretAtEnd(inner: HTMLElement): void {
    (inner.closest<HTMLElement>(".hwe-workspace") ?? inner).focus({ preventScroll: true });

    const range = document.createRange();
    const endPosition = this.getLastCaretPosition(inner);
    if (endPosition) {
      range.setStart(endPosition.node, endPosition.offset);
    } else {
      range.selectNodeContents(inner);
      range.collapse(false);
    }

    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
  }

  private getLastCaretPosition(root: Node): { node: Node; offset: number } | null {
    for (let index = root.childNodes.length - 1; index >= 0; index--) {
      const child = root.childNodes[index];

      if (child.nodeType === Node.TEXT_NODE) {
        return { node: child, offset: child.textContent?.length ?? 0 };
      }

      if (child.nodeType !== Node.ELEMENT_NODE) continue;

      const element = child as HTMLElement;
      if (element.hasAttribute("data-hwe-caret")) continue;
      if (element.tagName === "BR") {
        return { node: root, offset: index };
      }

      const nested = this.getLastCaretPosition(element);
      if (nested) return nested;

      if (!isEmptyNode(element, false)) {
        return { node: root, offset: index + 1 };
      }
    }

    return null;
  }
}
