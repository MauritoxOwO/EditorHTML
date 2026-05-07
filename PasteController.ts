import { WordPasteImporter } from "../import/WordPasteImporter";
import { PageSetup } from "../Orquestador/PageGeometry";

export interface PasteResult {
  handled: boolean;
  affectedPage?: HTMLElement;
  pageSetup?: PageSetup;
}

export class PasteController {
  private readonly wordPasteImporter = new WordPasteImporter();

  handlePaste(event: ClipboardEvent, fallbackPage: HTMLElement): PasteResult {
    const html = event.clipboardData?.getData("text/html") ?? "";
    if (!html || !this.wordPasteImporter.isWordHtml(html)) return { handled: false };

    event.preventDefault();

    const targetEditable = event.currentTarget as HTMLElement;
    const insertionMarker = this.createPasteInsertionMarker(targetEditable);
    const imported = this.wordPasteImporter.importFromHtml(html);
    const affectedPage =
      this.insertHtmlAtPasteMarker(imported.html, insertionMarker, targetEditable) ??
      fallbackPage;

    return {
      handled: true,
      affectedPage,
      pageSetup: imported.pageSetup,
    };
  }

  private createPasteInsertionMarker(targetEditable: HTMLElement): HTMLElement {
    const marker = document.createElement("span");
    marker.setAttribute("data-hwe-paste-marker", "true");
    marker.style.cssText = "display:inline-block;width:0;height:0;overflow:hidden;line-height:0;";

    const selection = window.getSelection();
    targetEditable.focus({ preventScroll: true });

    if (!selection || selection.rangeCount === 0) {
      targetEditable.appendChild(marker);
      return marker;
    }

    const range = selection.getRangeAt(0);
    if (!targetEditable.contains(range.commonAncestorContainer)) {
      targetEditable.appendChild(marker);
      return marker;
    }

    range.deleteContents();
    range.insertNode(marker);
    return marker;
  }

  private insertHtmlAtPasteMarker(
    html: string,
    marker: HTMLElement,
    targetEditable: HTMLElement
  ): HTMLElement | null {
    const template = document.createElement("template");
    template.innerHTML = html;
    const fragment = template.content;
    const insertedNodes = Array.from(fragment.childNodes);

    if (!marker.parentNode) {
      targetEditable.appendChild(fragment);
    } else {
      marker.replaceWith(fragment);
    }

    const selection = window.getSelection();
    const lastInserted = insertedNodes[insertedNodes.length - 1];
    if (selection && lastInserted?.parentNode) {
      const nextRange = document.createRange();
      nextRange.setStartAfter(lastInserted);
      nextRange.collapse(true);
      targetEditable.focus({ preventScroll: true });
      selection.removeAllRanges();
      selection.addRange(nextRange);
    }

    return targetEditable.closest<HTMLElement>(".hwe-page");
  }
}
