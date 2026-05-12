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
    const savedHtml = this.collectHtml(root, pages, pageSetup);
    const temp = document.createElement("div");
    temp.innerHTML = savedHtml;

    const savedDocument = this.getSavedDocumentWrapper(temp);
    const content = savedDocument?.innerHTML ?? temp.innerHTML;
    const preservedCss = Array.from(temp.querySelectorAll("style"))
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
  margin: ${setup.marginTop} ${setup.marginRight} ${setup.marginBottom} ${setup.marginLeft};
}
html,
body {
  margin: 0;
  padding: 0;
  background: #fff;
}
body {
  color: #000;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
}
[data-hwe-page-break="before"] {
  break-before: page;
  page-break-before: always;
}
table {
  border-collapse: collapse;
}
img {
  max-width: 100%;
  height: auto;
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