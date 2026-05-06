export class CaretManager {
  static createMarker(root: HTMLElement): HTMLElement | null {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0 || !selection.isCollapsed) return null;

    const anchorNode = selection.anchorNode;
    if (!anchorNode || !root.contains(anchorNode)) return null;

    const range = selection.getRangeAt(0).cloneRange();
    const marker = document.createElement("span");
    marker.setAttribute("data-hwe-caret", "true");
    marker.style.cssText =
      "display:inline-block;width:0;height:0;overflow:hidden;line-height:0;";

    range.insertNode(marker);
    return marker;
  }

  static restoreMarker(marker: HTMLElement | null, fallbackFocus?: HTMLElement | null): void {
    if (!marker || !marker.parentNode) {
      fallbackFocus?.focus({ preventScroll: true });
      return;
    }

    const range = document.createRange();
    range.setStartAfter(marker);
    range.collapse(true);

    const editable = marker.closest<HTMLElement>("[contenteditable='true']") ?? fallbackFocus;

    const selection = window.getSelection();
    if (!selection) {
      marker.remove();
      editable?.focus({ preventScroll: true });
      return;
    }

    editable?.focus({ preventScroll: true });
    selection.removeAllRanges();
    selection.addRange(range);
    marker.remove();
  }

  static removeMarkers(root: HTMLElement): void {
    root.querySelectorAll("[data-hwe-caret]").forEach((marker) => marker.remove());
  }
}
