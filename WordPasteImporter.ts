import {
  normalizePageSetup,
  PageSetup,
  sanitizeCssLength,
} from "../Orquestador/PageGeometry";

export interface WordPasteResult {
  html: string;
  isWordHtml: boolean;
  pageSetup?: PageSetup;
}

const WORD_MARKERS = [
  /class="?Mso/i,
  /mso-/i,
  /WordSection/i,
  /urn:schemas-microsoft-com:office/i,
  /xmlns:w=/i,
  /<!--\s*StartFragment\s*-->/i,
];

const SAFE_STYLE_PROPERTIES = new Set([
  "background",
  "background-color",
  "border",
  "border-bottom",
  "border-bottom-color",
  "border-bottom-style",
  "border-bottom-width",
  "border-collapse",
  "border-color",
  "border-left",
  "border-left-color",
  "border-left-style",
  "border-left-width",
  "border-right",
  "border-right-color",
  "border-right-style",
  "border-right-width",
  "border-spacing",
  "border-style",
  "border-top",
  "border-top-color",
  "border-top-style",
  "border-top-width",
  "border-width",
  "break-inside",
  "color",
  "font",
  "font-family",
  "font-size",
  "font-style",
  "font-weight",
  "height",
  "line-height",
  "margin",
  "margin-bottom",
  "margin-left",
  "margin-right",
  "margin-top",
  "max-height",
  "max-width",
  "min-height",
  "min-width",
  "padding",
  "padding-bottom",
  "padding-left",
  "padding-right",
  "padding-top",
  "page-break-inside",
  "table-layout",
  "text-align",
  "text-decoration",
  "text-decoration-color",
  "text-decoration-line",
  "text-decoration-style",
  "text-indent",
  "vertical-align",
  "white-space",
  "width",
]);

const SAFE_ATTRS = new Set([
  "alt",
  "border",
  "cellpadding",
  "cellspacing",
  "colspan",
  "height",
  "href",
  "rowspan",
  "src",
  "title",
  "width",
]);

export class WordPasteImporter {
  isWordHtml(html: string): boolean {
    return WORD_MARKERS.some((marker) => marker.test(html));
  }

  importFromHtml(html: string): WordPasteResult {
    const isWordHtml = this.isWordHtml(html);
    const doc = new DOMParser().parseFromString(html || "", "text/html");
    const pageSetup = this.extractPageSetup(doc);
    const fragment = this.getClipboardFragment(doc, html);

    this.cleanFragment(fragment);

    return {
      html: fragment.innerHTML.trim() || "<p><br></p>",
      isWordHtml,
      pageSetup,
    };
  }

  private getClipboardFragment(doc: Document, rawHtml: string): HTMLElement {
    const bodyHtml = doc.body?.innerHTML ?? rawHtml;
    const fragmentMatch = /<!--\s*StartFragment\s*-->([\s\S]*?)<!--\s*EndFragment\s*-->/i.exec(
      bodyHtml
    );

    const container = document.createElement("div");
    container.innerHTML = fragmentMatch?.[1] ?? bodyHtml;
    return container;
  }

  private extractPageSetup(doc: Document): PageSetup | undefined {
    const css = Array.from(doc.querySelectorAll("style"))
      .map((style) => style.textContent ?? "")
      .join("\n");

    const pageBlock = /@page[^{]*\{([\s\S]*?)\}/i.exec(css)?.[1];
    if (!pageBlock) return undefined;

    const pageSetup: Partial<PageSetup> = {};
    const size = this.getCssDeclaration(pageBlock, "size");
    if (size) {
      const parts = size
        .replace(/\s*!important/gi, "")
        .trim()
        .split(/\s+/)
        .filter(Boolean);

      if (parts.length >= 2) {
        pageSetup.width = sanitizeCssLength(parts[0]) ?? undefined;
        pageSetup.height = sanitizeCssLength(parts[1]) ?? undefined;
      } else if (/^a4$/i.test(parts[0] ?? "")) {
        pageSetup.width = "210mm";
        pageSetup.height = "297mm";
      } else if (/^letter$/i.test(parts[0] ?? "")) {
        pageSetup.width = "8.5in";
        pageSetup.height = "11in";
      }
    }

    const margin = this.getCssDeclaration(pageBlock, "margin");
    if (margin) {
      const margins = this.expandMarginShorthand(margin);
      pageSetup.marginTop = margins[0] ?? pageSetup.marginTop;
      pageSetup.marginRight = margins[1] ?? pageSetup.marginRight;
      pageSetup.marginBottom = margins[2] ?? pageSetup.marginBottom;
      pageSetup.marginLeft = margins[3] ?? pageSetup.marginLeft;
    }

    pageSetup.marginTop =
      sanitizeCssLength(this.getCssDeclaration(pageBlock, "margin-top")) ??
      pageSetup.marginTop;
    pageSetup.marginRight =
      sanitizeCssLength(this.getCssDeclaration(pageBlock, "margin-right")) ??
      pageSetup.marginRight;
    pageSetup.marginBottom =
      sanitizeCssLength(this.getCssDeclaration(pageBlock, "margin-bottom")) ??
      pageSetup.marginBottom;
    pageSetup.marginLeft =
      sanitizeCssLength(this.getCssDeclaration(pageBlock, "margin-left")) ??
      pageSetup.marginLeft;

    return Object.keys(pageSetup).length > 0 ? normalizePageSetup(pageSetup) : undefined;
  }

  private getCssDeclaration(cssBlock: string, property: string): string | null {
    const match = new RegExp(`${property}\\s*:\\s*([^;]+)`, "i").exec(cssBlock);
    return match?.[1]?.trim() ?? null;
  }

  private expandMarginShorthand(value: string): Array<string | undefined> {
    const parts = value
      .replace(/\s*!important/gi, "")
      .trim()
      .split(/\s+/)
      .map((part) => sanitizeCssLength(part) ?? undefined)
      .filter((part): part is string => !!part);

    if (parts.length === 0) return [];
    if (parts.length === 1) return [parts[0], parts[0], parts[0], parts[0]];
    if (parts.length === 2) return [parts[0], parts[1], parts[0], parts[1]];
    if (parts.length === 3) return [parts[0], parts[1], parts[2], parts[1]];
    return [parts[0], parts[1], parts[2], parts[3]];
  }

  private cleanFragment(root: HTMLElement): void {
    this.removeComments(root);
    root.querySelectorAll("style, meta, link, xml, script, object, iframe, form").forEach((node) => {
      node.remove();
    });
    root.querySelectorAll<HTMLElement>("div.WordSection1, div[class*='WordSection']").forEach(
      (element) => this.unwrapElement(element)
    );

    Array.from(root.querySelectorAll("*")).forEach((node) => {
      const element = node as HTMLElement;
      const tagName = element.tagName.toLowerCase();

      if (tagName.includes(":") && /^(o|v|w|m):/i.test(tagName)) {
        element.remove();
        return;
      }

      this.sanitizeAttributes(element);
      this.sanitizeElementStyle(element);
      this.normalizeWordElement(element);
    });

    this.removeVisuallyEmptyChrome(root);
  }

  private sanitizeAttributes(element: HTMLElement): void {
    Array.from(element.attributes).forEach((attr) => {
      const name = attr.name.toLowerCase();
      const value = attr.value;

      if (name === "style") return;
      if (name.startsWith("data-hwe-")) return;
      if (name === "class") {
        element.removeAttribute(attr.name);
        return;
      }
      if (name === "align") return;
      if (name.startsWith("on") || name.includes(":") || name.startsWith("xmlns")) {
        element.removeAttribute(attr.name);
        return;
      }
      if (!SAFE_ATTRS.has(name)) {
        element.removeAttribute(attr.name);
        return;
      }
      if ((name === "href" || name === "src") && !this.isSafeUrl(value)) {
        element.removeAttribute(attr.name);
      }
    });
  }

  private sanitizeElementStyle(element: HTMLElement): void {
    const style = element.getAttribute("style");
    if (!style) return;

    const sanitized = style
      .split(";")
      .map((declaration) => declaration.trim())
      .filter(Boolean)
      .map((declaration) => {
        const separator = declaration.indexOf(":");
        if (separator <= 0) return "";

        const property = declaration.slice(0, separator).trim().toLowerCase();
        const value = declaration.slice(separator + 1).trim();
        if (!SAFE_STYLE_PROPERTIES.has(property)) return "";
        if (property.startsWith("mso-")) return "";
        if (this.isUnsafeCssValue(value)) return "";
        if (/^(page-break-before|page-break-after|break-before|break-after)$/i.test(property)) {
          return "";
        }

        return `${property}: ${value}`;
      })
      .filter(Boolean)
      .join("; ");

    if (sanitized) {
      element.setAttribute("style", sanitized);
    } else {
      element.removeAttribute("style");
    }
  }

  private normalizeWordElement(element: HTMLElement): void {
    const tagName = element.tagName;

    if (tagName === "TABLE") {
      element.classList.add("hwe-word-table");
      element.setAttribute("data-hwe-source", "word");
      element.setAttribute("data-hwe-table", "word");

      const align = element.getAttribute("align")?.toLowerCase();
      if (align === "center") {
        element.style.marginLeft = "auto";
        element.style.marginRight = "auto";
        element.setAttribute("data-hwe-align", "center");
      }

      if (!element.style.borderCollapse) element.style.borderCollapse = "collapse";
      element.style.removeProperty("height");
      element.style.removeProperty("min-height");
      element.style.removeProperty("max-height");
      element.removeAttribute("height");
      this.moveLengthAttributeToStyle(element, "width");
    }

    if (["THEAD", "TBODY", "TFOOT"].includes(tagName)) {
      element.style.removeProperty("height");
      element.style.removeProperty("min-height");
      element.style.removeProperty("max-height");
      element.removeAttribute("height");
    }

    if (tagName === "COL" || tagName === "TD" || tagName === "TH") {
      this.moveLengthAttributeToStyle(element, "width");
      this.moveLengthAttributeToStyle(element, "height");
    }

    if (tagName === "P" || tagName === "DIV") {
      const align = element.getAttribute("align")?.toLowerCase();
      if (align && ["left", "center", "right", "justify"].includes(align)) {
        element.style.textAlign = align;
      }
    }

    element.removeAttribute("align");
  }

  private moveLengthAttributeToStyle(element: HTMLElement, attrName: "width" | "height"): void {
    const value = element.getAttribute(attrName);
    if (!value || element.style.getPropertyValue(attrName)) return;

    const cssValue = /^\d+$/.test(value.trim()) ? `${value.trim()}px` : value.trim();
    if (sanitizeCssLength(cssValue) || /^(?:\d+|\d*\.\d+)%$/.test(cssValue)) {
      element.style.setProperty(attrName, cssValue);
    }
  }

  private removeComments(root: HTMLElement): void {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_COMMENT);
    const comments: Node[] = [];

    while (walker.nextNode()) {
      comments.push(walker.currentNode);
    }

    comments.forEach((comment) => comment.parentNode?.removeChild(comment));
  }

  private removeVisuallyEmptyChrome(root: HTMLElement): void {
    Array.from(root.querySelectorAll<HTMLElement>("span, div")).forEach((element) => {
      const text = (element.textContent ?? "").replace(/\u00a0/g, " ").trim();
      const hasContent = !!element.querySelector("img, table, tr, td, th, br");
      const style = element.getAttribute("style") ?? "";

      if (!text && !hasContent && /^(?:border|padding|margin|font|color|background)/i.test(style)) {
        element.remove();
      }
    });
  }

  private unwrapElement(element: HTMLElement): void {
    if (!element.parentNode) return;

    const parent = element.parentNode;
    while (element.firstChild) {
      parent.insertBefore(element.firstChild, element);
    }
    parent.removeChild(element);
  }

  private isSafeUrl(value: string): boolean {
    const trimmed = value.trim();
    return /^(https?:|data:image\/|blob:|cid:|#|\/|\.\/|\.\.\/)/i.test(trimmed);
  }

  private isUnsafeCssValue(value: string): boolean {
    return /(expression\s*\(|javascript:|behavior\s*:|-moz-binding|url\s*\()/i.test(value);
  }
}
