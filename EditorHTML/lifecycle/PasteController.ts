import { WordPasteImporter } from "../import/WordPasteImporter";
import { PageSetup } from "../Orquestador/PageGeometry";
import { hweDebugLog, hweDebugStart } from "../debug/DebugLogger";

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

    this.normalizeInvalidTableAncestors(targetEditable);

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

  private normalizeInvalidTableAncestors(targetEditable: HTMLElement): void {
    let moved = 0;
    Array.from(targetEditable.querySelectorAll<HTMLTableElement>("p table, li table, span table"))
      .forEach((table) => {
        if (table.closest("td, th")) return;

        const invalidAncestor = this.getOuterInvalidTableAncestor(table, targetEditable);
        const reference = invalidAncestor?.tagName === "LI"
          ? invalidAncestor.closest<HTMLElement>("ul, ol") ?? invalidAncestor
          : invalidAncestor;
        const parent = reference?.parentNode;
        if (!invalidAncestor || !reference || !parent || invalidAncestor === targetEditable) {
          return;
        }
        if (!targetEditable.contains(invalidAncestor)) return;

        parent.insertBefore(table, reference.nextSibling);
        moved++;
        if (this.isVisuallyEmpty(invalidAncestor)) invalidAncestor.remove();
      });

    if (moved > 0) {
      hweDebugLog("paste.normalizeInvalidTableAncestors", { moved });
    }
  }

  private getOuterInvalidTableAncestor(
    table: HTMLTableElement,
    targetEditable: HTMLElement
  ): HTMLElement | null {
    let invalid: HTMLElement | null = null;
    let current = table.parentElement;

    while (current && current !== targetEditable && targetEditable.contains(current)) {
      if (current.matches("p, li, span")) invalid = current;
      current = current.parentElement;
    }

    return invalid;
  }

  private isVisuallyEmpty(element: HTMLElement): boolean {
    const text = (element.textContent ?? "").replace(/\u00a0/g, " ").trim();
    if (text) return false;
    return !element.querySelector("img, table, tr, td, th, video, canvas, svg");
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
