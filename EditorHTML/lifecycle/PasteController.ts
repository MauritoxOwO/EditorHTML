import { WordPasteImporter } from "../import/WordPasteImporter";
import { PageSetup } from "../Orquestador/PageGeometry";
import { hweDebugStart } from "../debug/DebugLogger";

export interface PasteResult {
  handled: boolean;
  affectedPage?: HTMLElement;
  pageSetup?: PageSetup;
}

export class PasteController {
  private readonly wordPasteImporter = new WordPasteImporter();

  handlePaste(event: ClipboardEvent, fallbackPage: HTMLElement): PasteResult {
    const done = hweDebugStart("paste.handlePaste");
    const html = event.clipboardData?.getData("text/html") ?? "";
    if (!html || !this.wordPasteImporter.isWordHtml(html)) {
      done({ handled: false, htmlLength: html.length });
      return { handled: false };
    }

    event.preventDefault();

    const targetEditable = event.currentTarget as HTMLElement;
    const imported = this.wordPasteImporter.importFromHtml(html);
    const insertAsBlock = this.shouldInsertAsBlock(imported.html);
    const insertionMarker = this.createPasteInsertionMarker(targetEditable, insertAsBlock);
    const affectedPage =
      this.insertHtmlAtPasteMarker(imported.html, insertionMarker, targetEditable) ??
      fallbackPage;

    done({
      affectedPage: !!affectedPage,
      handled: true,
      importedLength: imported.html.length,
      insertAsBlock,
      rows: this.countMatches(imported.html, /<tr\b/gi),
      tables: this.countMatches(imported.html, /<table\b/gi),
    });

    return {
      handled: true,
      affectedPage,
      pageSetup: imported.pageSetup,
    };
  }

  private createPasteInsertionMarker(
    targetEditable: HTMLElement,
    insertAsBlock: boolean
  ): HTMLElement {
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
    if (insertAsBlock && this.insertMarkerAtBlockBoundary(marker, range, targetEditable)) {
      return marker;
    }

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

  private shouldInsertAsBlock(html: string): boolean {
    return /<(?:table|h[1-6]|p|div|section|article|ul|ol|blockquote)\b/i.test(html);
  }

  private insertMarkerAtBlockBoundary(
    marker: HTMLElement,
    range: Range,
    targetEditable: HTMLElement
  ): boolean {
    const container =
      range.startContainer.nodeType === Node.ELEMENT_NODE
        ? (range.startContainer as Element)
        : range.startContainer.parentElement;
    const block = container?.closest<HTMLElement>(
      "p, li, h1, h2, h3, h4, h5, h6, blockquote, pre, div"
    );
    if (!block || block === targetEditable || !targetEditable.contains(block)) return false;
    if (block.closest("td, th")) return false;

    block.parentNode?.insertBefore(marker, block.nextSibling);
    return !!marker.parentNode;
  }

  private countMatches(value: string, pattern: RegExp): number {
    return value.match(pattern)?.length ?? 0;
  }
}
