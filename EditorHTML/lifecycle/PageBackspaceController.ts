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
  handleBackspaceAtPageStart(event: KeyboardEvent, context: PageBackspaceContext): boolean {
    const inner = event.currentTarget as HTMLElement;
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
      this.deleteLastContent(previousInner);
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
      if (element.querySelector("img, video, canvas, svg")) return true;
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

  private fragmentHasVisibleContent(fragment: DocumentFragment): boolean {
    return Array.from(fragment.childNodes).some((node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        return (node.textContent ?? "").replace(/\u00a0/g, " ").length > 0;
      }

      if (node.nodeType !== Node.ELEMENT_NODE) return false;

      const element = node as HTMLElement;
      if (element.tagName === "BR") return false;
      if (element.querySelector("img, table, tr, td, th, video, canvas, svg")) return true;
      return (element.textContent ?? "").replace(/\u00a0/g, " ").length > 0;
    });
  }

  private deleteLastContent(container: HTMLElement): void {
    const removed = this.deleteLastContentFromNode(container);
    if (!removed && getMeaningfulChildren(container, true).length === 0) {
      container.innerHTML = "<p><br></p>";
    }
  }

  private deleteLastContentFromNode(node: Node): boolean {
    for (let index = node.childNodes.length - 1; index >= 0; index--) {
      const child = node.childNodes[index];

      if (child.nodeType === Node.TEXT_NODE) {
        const text = child.textContent ?? "";
        if (text.length === 0) {
          child.remove();
          continue;
        }

        child.textContent = text.slice(0, -1);
        if (child.textContent.length === 0) child.remove();
        return true;
      }

      if (child.nodeType !== Node.ELEMENT_NODE) {
        child.remove();
        return true;
      }

      const element = child as HTMLElement;
      if (element.hasAttribute("data-hwe-caret")) {
        element.remove();
        continue;
      }

      if (element.tagName === "BR" || this.isAtomicEditableElement(element)) {
        element.remove();
        return true;
      }

      if (this.deleteLastContentFromNode(element)) {
        if (isEmptyNode(element, false)) element.remove();
        return true;
      }

      if (isEditableBlankBlock(element)) {
        element.remove();
        return true;
      }
    }

    return false;
  }

  private isAtomicEditableElement(element: HTMLElement): boolean {
    return ["IMG", "TABLE", "VIDEO", "CANVAS", "SVG"].includes(element.tagName);
  }

  private placeCaretAtEnd(inner: HTMLElement): void {
    inner.focus({ preventScroll: true });

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