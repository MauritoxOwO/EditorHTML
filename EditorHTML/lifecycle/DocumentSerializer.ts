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

export interface NormalizedDocument {
  html: string;
  pageSetup?: PageSetup;
}

export class DocumentSerializer {
  private readonly wordPasteImporter = new WordPasteImporter();

  normalizeHtmlForPagination(html: string): NormalizedDocument {
    const doc = new DOMParser().parseFromString(html || "<p><br></p>", "text/html");
    const temp = document.createElement("div");
    temp.innerHTML = doc.body?.innerHTML || html || "<p><br></p>";

    const savedDocument = this.getSavedDocumentWrapper(temp);
    const pageSetup = savedDocument ? readPageSetupFromElement(savedDocument) ?? undefined : undefined;
    if (savedDocument) {
      temp.innerHTML = savedDocument.innerHTML;
      this.removeRuntimeOnlyState(temp);
      return {
        html: temp.innerHTML || "<p><br></p>",
        pageSetup,
      };
    }

    if (this.wordPasteImporter.isWordHtml(html)) {
      const imported = this.wordPasteImporter.importFromHtml(html);
      return {
        html: imported.html,
        pageSetup: imported.pageSetup,
      };
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

    return {
      html: temp.innerHTML || "<p><br></p>",
      pageSetup,
    };
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
    return clone.innerHTML;
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
.hwe-page-inner > div:has(> table:only-child),
.hwe-page-inner > section:has(> table:only-child),
.hwe-page-inner > article:has(> table:only-child) {
  width: 100% !important;
  max-width: 100% !important;
  margin-left: 0;
  margin-right: 0;
}
.hwe-page-inner > div:has(> table:only-child) > table,
.hwe-page-inner > section:has(> table:only-child) > table,
.hwe-page-inner > article:has(> table:only-child) > table,
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
