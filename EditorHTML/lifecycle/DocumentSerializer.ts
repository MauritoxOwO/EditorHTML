import { WordPasteImporter } from "../import/WordPasteImporter";
import {
  makePageSetupWrapper,
  PageSetup,
  readPageSetupFromElement,
} from "../Orquestador/PageGeometry";
import { CaretManager } from "../Orquestador/CaretManager";
import {
  getMeaningfulChildren,
  isEmptyNode,
  isSplittableContainer,
  removeComments,
  unwrapElement,
} from "../dom/EditableDom";
import { unwrapGeneratedKeepTogetherGroups } from "../Orquestador/KeepTogetherController";
import { addOcrTextLayers } from "../ocr/PdfOcrLayer";
import { hweDebugLog, hweDebugStart } from "../debug/DebugLogger";

export interface NormalizedDocument {
  html: string;
  pageSetup?: PageSetup;
}

const LARGE_DATA_IMAGE_URL_THRESHOLD = 80_000;
const LARGE_DATA_IMAGE_PIXEL_THRESHOLD = 1_500_000;
const LARGE_IMAGE_PLACEHOLDER_SRC =
  "data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%221850%22%20height%3D%222420%22%20viewBox%3D%220%200%201850%202420%22%3E%3Crect%20width%3D%221850%22%20height%3D%222420%22%20fill%3D%22%23f7f7f7%22%2F%3E%3Crect%20x%3D%2250%22%20y%3D%2250%22%20width%3D%221750%22%20height%3D%222320%22%20fill%3D%22none%22%20stroke%3D%22%23bdbdbd%22%20stroke-width%3D%226%22%20stroke-dasharray%3D%2224%2024%22%2F%3E%3Ctext%20x%3D%22925%22%20y%3D%221210%22%20text-anchor%3D%22middle%22%20font-family%3D%22Arial%2C%20sans-serif%22%20font-size%3D%2272%22%20fill%3D%22%23666%22%3EImagen%20grande%20en%20pausa%3C%2Ftext%3E%3C%2Fsvg%3E";

export class DocumentSerializer {
  private readonly wordPasteImporter = new WordPasteImporter();
  private readonly detachedLargeImages = new Map<string, string>();
  private largeImageCounter = 0;

  normalizeHtmlForPagination(html: string): NormalizedDocument {
    const done = hweDebugStart("serializer.normalizeHtmlForPagination", {
      htmlLength: html.length,
    });
    const safeHtml = this.detachLargeDataImages(html || "<p><br></p>");
    const doc = new DOMParser().parseFromString(safeHtml, "text/html");
    const temp = document.createElement("div");
    temp.innerHTML = doc.body?.innerHTML || safeHtml || "<p><br></p>";

    const savedDocument = this.getSavedDocumentWrapper(temp);
    const pageSetup = savedDocument ? readPageSetupFromElement(savedDocument) ?? undefined : undefined;
    if (savedDocument) {
      temp.innerHTML = savedDocument.innerHTML;
      this.removeRuntimeOnlyState(temp);
      const result = {
        html: temp.innerHTML || "<p><br></p>",
        pageSetup,
      };
      done({
        normalizedLength: result.html.length,
        pageSetup: result.pageSetup ?? null,
        source: "saved-document",
      });
      return result;
    }

    if (this.wordPasteImporter.isWordHtml(safeHtml)) {
      const imported = this.wordPasteImporter.importFromHtml(safeHtml);
      const result = {
        html: imported.html,
        pageSetup: imported.pageSetup,
      };
      done({
        normalizedLength: result.html.length,
        pageSetup: result.pageSetup ?? null,
        source: "word-html",
      });
      return result;
    }

    this.preserveHeadStyles(doc, temp);
    this.applyBodyFormattingWrapper(doc, temp);
    this.cleanImportedHtml(temp);

    let safety = 0;
    while (safety++ < 20) {
      const children = getMeaningfulChildren(temp);
      if (children.length !== 1 || children[0].nodeType !== Node.ELEMENT_NODE) break;

      const onlyChild = children[0] as HTMLElement;
      if (!isSplittableContainer(onlyChild)) break;
      if (this.hasFormattingShell(onlyChild)) break;

      temp.innerHTML = onlyChild.innerHTML;
    }

    const result = {
      html: temp.innerHTML || "<p><br></p>",
      pageSetup,
    };
    done({
      normalizedLength: result.html.length,
      pageSetup: result.pageSetup ?? null,
      source: "generic-html",
    });
    return result;
  }

  collectHtml(
    root: HTMLElement,
    pages: HTMLElement[],
    pageSetup: PageSetup
  ): string {
    CaretManager.removeMarkers(root);

    const html = pages
      .map((page, index) => {
        const inner = page.querySelector(".hwe-page-inner") as HTMLElement | null;
        const content = this.prepareContentForSave(inner ?? page);
        if (index === 0) return content;
        return `<div data-hwe-page-break="before" style="page-break-before:always">${content}</div>`;
      })
      .join("\n");

    return makePageSetupWrapper(html, pageSetup);
  }

  collectPdfHtml(
    root: HTMLElement,
    pages: HTMLElement[],
    pageSetup: PageSetup,
    additionalCss = ""
  ): string {
    CaretManager.removeMarkers(root);

    const content = pages
      .map((page) => this.preparePageForPdf(page))
      .join("\n");
    const preservedCss = Array.from(root.querySelectorAll("style"))
      .map((style) => style.textContent ?? "")
      .filter(Boolean)
      .join("\n");
    const css = this.sanitizeStyleText([preservedCss, additionalCss].filter(Boolean).join("\n"));

    return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<style>
${this.makePdfPageCss(pageSetup)}
${css}
</style>
</head>
<body>
${content}
</body>
</html>`;
  }

  private preparePageForPdf(page: HTMLElement): string {
    const clone = page.cloneNode(true) as HTMLElement;
    unwrapGeneratedKeepTogetherGroups(clone);
    this.removeRuntimeOnlyState(clone);
    clone.querySelectorAll<HTMLElement>("[contenteditable]").forEach((element) => {
      element.removeAttribute("contenteditable");
    });
    clone.querySelectorAll<HTMLElement>("[spellcheck]").forEach((element) => {
      element.removeAttribute("spellcheck");
    });
    addOcrTextLayers(clone);
    return clone.outerHTML;
  }

  private getSavedDocumentWrapper(root: HTMLElement): HTMLElement | null {
    const directChildren = getMeaningfulChildren(root, true);
    if (directChildren.length === 1 && directChildren[0].nodeType === Node.ELEMENT_NODE) {
      const onlyChild = directChildren[0] as HTMLElement;
      if (onlyChild.matches("[data-hwe-document='true']")) return onlyChild;
    }

    return root.querySelector<HTMLElement>("[data-hwe-document='true']");
  }

  private cleanImportedHtml(root: HTMLElement): void {
    removeComments(root);
    root.querySelectorAll("meta, link, xml, script, object, hr").forEach((node) => {
      node.remove();
    });
    this.removeOfficeNamespacedNodes(root);

    root.querySelectorAll<HTMLElement>("[style]").forEach((element) => {
      element.style.cssText = this.sanitizeInlineStyle(element.style.cssText);
      if (!element.getAttribute("style")) element.removeAttribute("style");
    });

    root.querySelectorAll<HTMLElement>("[width]").forEach((element) => {
      if (!this.canKeepDimensionAttribute(element)) element.removeAttribute("width");
    });
    root.querySelectorAll<HTMLElement>("[height]").forEach((element) => {
      if (!this.canKeepDimensionAttribute(element)) element.removeAttribute("height");
    });

    this.unwrapKnownWordContainers(root);
    this.removeVisuallyEmptyNodes(root);
    this.removeBorderOnlyBlocks(root);
  }

  private preserveHeadStyles(doc: Document, root: HTMLElement): void {
    const css = Array.from(doc.head?.querySelectorAll("style") ?? [])
      .map((style) => style.textContent ?? "")
      .filter(Boolean)
      .join("\n");
    const sanitizedCss = this.sanitizeStyleText(css);
    if (!sanitizedCss) return;

    const style = document.createElement("style");
    style.setAttribute("data-hwe-preserved-style", "true");
    style.textContent = sanitizedCss;
    root.insertBefore(style, root.firstChild);
  }

  private prepareContentForSave(element: HTMLElement): string {
    const clone = element.cloneNode(true) as HTMLElement;
    unwrapGeneratedKeepTogetherGroups(clone);
    this.removeRuntimeOnlyState(clone);
    this.restoreDetachedLargeImages(clone);
    return clone.innerHTML;
  }

  private detachLargeDataImages(html: string): string {
    const dataImages = this.countDataImageSources(html);
    let detached = 0;
    const reasons: Record<string, number> = {};
    const result = html.replace(
      /<img\b[^>]*\bsrc\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))[^>]*>/gi,
      (tag: string, doubleQuotedSrc: string, singleQuotedSrc: string, unquotedSrc: string) => {
        const src = doubleQuotedSrc ?? singleQuotedSrc ?? unquotedSrc ?? "";
        if (!/^data:image\//i.test(src)) return tag;

        const reason = this.getLargeDataImageReason(tag, src);
        if (!reason) return tag;

        detached++;
        reasons[reason] = (reasons[reason] ?? 0) + 1;
        const id = this.rememberDetachedLargeImage(src);
        const quote = singleQuotedSrc !== undefined ? "'" : "\"";
        return tag
          .replace(
            /\s+src\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/i,
            ` src=${quote}${LARGE_IMAGE_PLACEHOLDER_SRC}${quote}`
          )
          .replace(
            /<img\b/i,
            `<img data-hwe-large-image-id=${quote}${id}${quote} data-hwe-large-image-placeholder=${quote}true${quote} data-hwe-large-image-reason=${quote}${reason}${quote}`
          );
      }
    );

    if (detached > 0) {
      hweDebugLog("serializer.detachLargeDataImages", {
        dataImages,
        detached,
        htmlLength: html.length,
        reasons,
      });
    } else if (dataImages > 0) {
      hweDebugLog("serializer.detachLargeDataImages.none", {
        dataImages,
        htmlLength: html.length,
      });
    }

    return result;
  }

  private countDataImageSources(html: string): number {
    return html.match(/\bsrc\s*=\s*(?:"data:image\/|'data:image\/|data:image\/)/gi)?.length ?? 0;
  }

  private getLargeDataImageReason(tag: string, src: string): string | null {
    const declaredPixels = this.getDeclaredImagePixels(tag);
    if (declaredPixels > LARGE_DATA_IMAGE_PIXEL_THRESHOLD) {
      return "declared-pixels";
    }

    const bytes = this.estimateDataUrlBytes(src);
    if (bytes > LARGE_DATA_IMAGE_URL_THRESHOLD) {
      return "data-url-bytes";
    }

    return null;
  }

  private rememberDetachedLargeImage(src: string): string {
    const id = `hwe-large-image-${Date.now().toString(36)}-${this.largeImageCounter++}`;
    this.detachedLargeImages.set(id, src);
    return id;
  }

  private restoreDetachedLargeImages(root: HTMLElement): void {
    let restored = 0;
    root.querySelectorAll<HTMLImageElement>("img[data-hwe-large-image-id]").forEach((image) => {
      const id = image.getAttribute("data-hwe-large-image-id") ?? "";
      const src = this.detachedLargeImages.get(id);
      if (src) {
        image.setAttribute("src", src);
        restored++;
      }
      image.removeAttribute("data-hwe-large-image-id");
      image.removeAttribute("data-hwe-large-image-placeholder");
      image.removeAttribute("data-hwe-large-image-reason");
    });
    if (restored > 0) hweDebugLog("serializer.restoreDetachedLargeImages", { restored });
  }

  private getDeclaredImagePixels(tag: string): number {
    const width = this.getNumericAttribute(tag, "width");
    const height = this.getNumericAttribute(tag, "height");
    return width > 0 && height > 0 ? width * height : 0;
  }

  private getNumericAttribute(tag: string, name: string): number {
    const match = new RegExp(`\\s${name}\\s*=\\s*["']?([0-9.]+)`, "i").exec(tag);
    return match ? Number.parseFloat(match[1]) || 0 : 0;
  }

  private estimateDataUrlBytes(src: string): number {
    const commaIndex = src.indexOf(",");
    const payload = commaIndex >= 0 ? src.slice(commaIndex + 1) : src;
    if (/^data:[^;,]+;base64,/i.test(src)) {
      return Math.floor(payload.length * 0.75);
    }

    return payload.length;
  }

  private applyBodyFormattingWrapper(doc: Document, root: HTMLElement): void {
    const body = doc.body;
    if (!body) return;

    const style = body.getAttribute("style");
    const className = body.getAttribute("class");
    if (!style && !className) return;

    const wrapper = document.createElement("div");
    if (style) wrapper.setAttribute("style", this.sanitizeInlineStyle(style));
    if (className) wrapper.setAttribute("class", className);
    if (!wrapper.getAttribute("style") && !wrapper.getAttribute("class")) return;

    while (root.firstChild) {
      wrapper.appendChild(root.firstChild);
    }
    root.appendChild(wrapper);
  }

  private canKeepDimensionAttribute(element: HTMLElement): boolean {
    return ["IMG", "TABLE", "COL", "COLGROUP", "TD", "TH"].includes(element.tagName);
  }

  private removeRuntimeOnlyState(root: HTMLElement): void {
    CaretManager.removeMarkers(root);
    root.querySelectorAll<HTMLElement>("[data-hwe-caret], [data-hwe-paste-marker]").forEach(
      (element) => element.remove()
    );
    root.querySelectorAll<HTMLElement>("[data-hwe-ocr-state]").forEach((element) => {
      element.removeAttribute("data-hwe-ocr-state");
    });
  }

  private hasFormattingShell(element: HTMLElement): boolean {
    return Array.from(element.attributes).some((attr) => {
      const name = attr.name.toLowerCase();
      const value = attr.value.trim();
      if (!value) return false;
      if (name.startsWith("data-hwe-")) return false;
      if (name === "style" && /^page-break-before\s*:\s*always\s*;?$/i.test(value)) {
        return false;
      }

      return true;
    });
  }

  private sanitizeInlineStyle(style: string): string {
    return style
      .replace(/mso-[^:;]+:[^;]+;?/gi, "")
      .replace(/page-break-before\s*:\s*always\s*;?/gi, "")
      .replace(/page-break-after\s*:\s*always\s*;?/gi, "")
      .replace(/break-before\s*:\s*page\s*;?/gi, "")
      .replace(/break-after\s*:\s*page\s*;?/gi, "")
      .replace(/position\s*:[^;]+;?/gi, "")
      .replace(/left\s*:[^;]+;?/gi, "")
      .replace(/right\s*:[^;]+;?/gi, "")
      .replace(/transform\s*:[^;]+;?/gi, "")
      .replace(/overflow\s*:[^;]+;?/gi, "")
      .replace(/overflow-x\s*:[^;]+;?/gi, "")
      .replace(/overflow-y\s*:[^;]+;?/gi, "")
      .replace(/tab-stops\s*:[^;]+;?/gi, "")
      .replace(/behavior\s*:[^;]+;?/gi, "")
      .replace(/-moz-binding\s*:[^;]+;?/gi, "")
      .replace(/url\s*\([^)]*\)\s*;?/gi, "")
      .replace(/expression\s*\([^)]*\)\s*;?/gi, "")
      .trim();
  }

  private removeOfficeNamespacedNodes(root: HTMLElement): void {
    Array.from(root.querySelectorAll("*")).forEach((node) => {
      const tagName = node.tagName.toLowerCase();
      if (tagName.includes(":") && /^(o|v|w|m):/i.test(tagName)) node.remove();
    });
  }

  private unwrapKnownWordContainers(root: HTMLElement): void {
    root.querySelectorAll<HTMLElement>("div.WordSection1, div[class*='WordSection']").forEach(
      (element) => unwrapElement(element)
    );
  }

  private removeVisuallyEmptyNodes(root: HTMLElement): void {
    Array.from(root.childNodes).forEach((node) => {
      if (node.nodeType === Node.ELEMENT_NODE && (node as HTMLElement).tagName === "STYLE") {
        return;
      }

      if (node.nodeType === Node.ELEMENT_NODE) {
        this.removeVisuallyEmptyNodes(node as HTMLElement);
      }

      if (isEmptyNode(node)) {
        node.parentNode?.removeChild(node);
      }
    });
  }

  private removeBorderOnlyBlocks(root: HTMLElement): void {
    Array.from(root.querySelectorAll<HTMLElement>("p, div, section, article, span")).forEach(
      (element) => {
        const text = (element.textContent ?? "").replace(/\u00a0/g, " ").trim();
        const hasMedia = !!element.querySelector("img, table, tr, td, th, video, canvas, svg");
        const hasBorder = /border/i.test(element.getAttribute("style") ?? "");

        if (!text && !hasMedia && hasBorder) element.remove();
      }
    );
  }

  private makePdfPageCss(setup: PageSetup): string {
    return `
@page {
  size: ${setup.width} ${setup.height};
  margin: 0;
}
html,
body {
  margin: 0;
  padding: 0;
  background: #fff;
}
body {
  width: ${setup.width};
  color: #000;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
}
.hwe-page {
  width: ${setup.width};
  height: ${setup.height};
  margin: 0;
  padding: 0;
  position: relative;
  overflow: hidden;
  box-sizing: border-box;
  background: #fff;
  box-shadow: none;
  break-after: page;
  page-break-after: always;
}
.hwe-page:last-child {
  break-after: auto;
  page-break-after: auto;
}
.hwe-page-inner {
  width: 100%;
  height: 100%;
  padding: ${setup.marginTop} ${setup.marginRight} ${setup.marginBottom} ${setup.marginLeft};
  box-sizing: border-box;
  outline: none;
  overflow: visible;
  color: #000;
  font-family: Calibri, "Segoe UI", Arial, sans-serif;
  font-size: 11pt;
  line-height: 1.15;
  overflow-wrap: break-word;
  word-wrap: break-word;
  white-space: normal;
}
.hwe-page-inner *,
.hwe-page-inner *::before,
.hwe-page-inner *::after {
  box-sizing: border-box;
  max-width: 100%;
}
.hwe-page p {
  min-height: 1em;
  margin: 0 0 8pt;
}
.hwe-page h1 {
  margin: 12pt 0 6pt;
  font-size: 20pt;
  font-weight: 700;
  line-height: 1.2;
}
.hwe-page h2 {
  margin: 10pt 0 4pt;
  font-size: 16pt;
  font-weight: 700;
  line-height: 1.2;
}
.hwe-page h3 {
  margin: 8pt 0 4pt;
  font-size: 13pt;
  font-weight: 700;
  line-height: 1.2;
}
.hwe-page ul,
.hwe-page ol {
  margin: 0 0 8pt;
  padding-left: 24pt;
}
.hwe-page li {
  margin-bottom: 2pt;
}
.hwe-page-inner
  > :where(p, h1, h2, h3, h4, h5, h6, blockquote, pre, ul, ol, li) {
  display: block !important;
  width: 140mm !important;
  max-width: 100%;
  margin-left: auto !important;
  margin-right: auto !important;
}
table {
  border-collapse: collapse;
  max-width: 100% !important;
  break-inside: auto;
}
.hwe-page-inner > table,
.hwe-table-flow-wrapper {
  width: 100% !important;
  max-width: 100% !important;
  margin-left: 0;
  margin-right: 0;
}
.hwe-table-flow-wrapper > table,
.hwe-page table.hwe-word-table {
  width: 100% !important;
}
.hwe-page table:not(.hwe-word-table) {
  width: 100% !important;
  table-layout: fixed;
}
.hwe-page table:not(.hwe-word-table) td,
.hwe-page table:not(.hwe-word-table) th {
  padding: 4pt 6pt;
  border: 1px solid #999;
  vertical-align: top;
  word-break: break-word;
  overflow-wrap: anywhere;
}
.hwe-page table:not(.hwe-word-table) th {
  background: #f0f0f0;
  font-weight: 700;
}
.hwe-page table.hwe-word-table {
  table-layout: auto;
  border-collapse: collapse;
  margin-top: 0;
  margin-bottom: 0;
  page-break-inside: auto;
}
.hwe-page table.hwe-word-table td,
.hwe-page table.hwe-word-table th {
  vertical-align: top;
  word-break: normal;
  overflow-wrap: break-word;
}
.hwe-page table.hwe-word-table p {
  min-height: 0;
  margin: 0;
  line-height: inherit;
  width: auto !important;
}
.hwe-page table.hwe-word-table thead,
.hwe-page table.hwe-word-table tbody,
.hwe-page table.hwe-word-table tfoot {
  break-inside: auto;
  page-break-inside: auto;
}
.hwe-page table.hwe-word-table tr {
  break-inside: avoid;
  page-break-inside: avoid;
}
img {
  display: block;
  max-width: 100%;
  height: auto;
  break-inside: avoid;
  page-break-inside: avoid;
}
.hwe-ocr-wrapper {
  display: block;
  position: relative;
  max-width: 100%;
}
.hwe-ocr-layer {
  position: absolute;
  inset: 0;
  overflow: hidden;
  color: rgba(0, 0, 0, 0.01);
  font-size: 1px;
  line-height: 1;
  white-space: pre-wrap;
  pointer-events: none;
}
`;
  }

  private sanitizeStyleText(css: string): string {
    return css
      .replace(/<\/style/gi, "<\\/style")
      .replace(/expression\s*\([^)]*\)/gi, "")
      .replace(/url\s*\(\s*(['"]?)javascript:[^)]*\)/gi, "url()");
  }
}