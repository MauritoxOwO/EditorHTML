import {
  EDITABLE_BLANK_BLOCK_SELECTOR,
  isEditableBlankBlock,
} from "../dom/EditableDom";

export class BlankLineController {
  syncEditableBlankBlocks(root: HTMLElement, markNewBlanks: boolean): void {
    root.querySelectorAll<HTMLElement>(EDITABLE_BLANK_BLOCK_SELECTOR).forEach((element) => {
      if (!isEditableBlankBlock(element)) {
        element.removeAttribute("data-hwe-user-blank");
        return;
      }

      if (markNewBlanks || element.hasAttribute("data-hwe-user-blank")) {
        element.setAttribute("data-hwe-user-blank", "true");
        this.ensureBlankBlockHasCaretStop(element);
      }
    });
  }

  private ensureBlankBlockHasCaretStop(element: HTMLElement): void {
    if (element.childNodes.length === 0) {
      element.appendChild(document.createElement("br"));
      return;
    }

    const hasBreak = Array.from(element.childNodes).some(
      (child) => child.nodeType === Node.ELEMENT_NODE && (child as HTMLElement).tagName === "BR"
    );
    if (!hasBreak && (element.textContent ?? "").length === 0) {
      element.appendChild(document.createElement("br"));
    }
  }
}
