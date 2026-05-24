import { ParagraphStyleOption, Toolbar } from "../ui/Toolbar";
import {
  ParagraphFontFaceDefinition,
  ParagraphStyleCatalog,
  ParagraphStyleDefinition,
} from "./dataverse/styleApi";

const GENERIC_FONT_FAMILIES = new Set([
  "serif",
  "sans-serif",
  "monospace",
  "cursive",
  "fantasy",
  "system-ui",
  "ui-serif",
  "ui-sans-serif",
  "ui-monospace",
  "ui-rounded",
  "emoji",
  "math",
  "fangsong",
  "inherit",
  "initial",
  "revert",
  "unset",
]);

export class ParagraphStyleManager {
  private readonly classNames = new Set<string>();
  private readonly fontFamilyClassNames = new Set<string>();
  private fontStyleElement: HTMLStyleElement | null = null;
  private styleElement: HTMLStyleElement | null = null;

  constructor(
    private readonly rootProvider: () => HTMLElement,
    private readonly toolbar: Toolbar
  ) {}

  get cssText(): string {
    return [this.fontStyleElement?.textContent ?? "", this.styleElement?.textContent ?? ""]
      .filter(Boolean)
      .join("\n\n");
  }

  hasClass(className: string): boolean {
    return this.classNames.has(className);
  }

  setStyles(styles: ParagraphStyleDefinition[]): void {
    this.setCatalog({
      styles,
      fonts: [],
    });
  }

  setCatalog(catalog: ParagraphStyleCatalog): void {
    const styles = catalog.styles;
    const validStyles = styles.filter(
      (style) => this.isValidCssClassName(style.className) && !this.isFontFaceCss(style.cssText)
    );
    const referencedFontFamilies = this.getReferencedFontFamilies(validStyles);
    const fonts = this.dedupeFonts(
      this.filterFontsForStyles(catalog.fonts, referencedFontFamilies)
    );

    this.classNames.clear();
    this.fontFamilyClassNames.clear();
    validStyles.forEach((style) => {
      this.classNames.add(style.className);
      if (this.styleDefinesFontFamily(style)) {
        this.fontFamilyClassNames.add(style.className);
      }
    });
    this.toolbar.setParagraphStyles(
      validStyles.map<ParagraphStyleOption>((style) => ({
        label: style.label,
        className: style.className,
      }))
    );
    this.injectFontFaceCss(fonts);
    this.injectParagraphStyleCss(validStyles);
  }

  applyToBlock(block: HTMLElement, className: string | null): void {
    this.classNames.forEach((styleClass) => block.classList.remove(styleClass));
    if (!className) return;

    if (this.fontFamilyClassNames.has(className)) {
      this.clearInlineFontFamily(block);
    }
    block.classList.add(className);
  }

  destroy(): void {
    this.fontStyleElement?.remove();
    this.styleElement?.remove();
    this.fontStyleElement = null;
    this.styleElement = null;
  }

  private injectFontFaceCss(fonts: ParagraphFontFaceDefinition[]): void {
    if (!this.fontStyleElement) {
      this.fontStyleElement = document.createElement("style");
      this.fontStyleElement.setAttribute("data-hwe-font-catalog", "true");
      document.head.appendChild(this.fontStyleElement);
    }

    this.fontStyleElement.textContent = fonts
      .map((font) => this.formatFontFaceCss(font))
      .filter(Boolean)
      .join("\n\n");
  }

  private injectParagraphStyleCss(styles: ParagraphStyleDefinition[]): void {
    if (!this.styleElement) {
      this.styleElement = document.createElement("style");
      this.styleElement.setAttribute("data-hwe-style-catalog", "true");
      this.rootProvider().insertBefore(this.styleElement, this.rootProvider().firstChild);
    }

    this.styleElement.textContent = styles
      .map((style) => this.formatParagraphStyleCss(style))
      .filter(Boolean)
      .join("\n\n");
  }

  private formatParagraphStyleCss(style: ParagraphStyleDefinition): string {
    const declarations = this.extractCssDeclarations(style.cssText);
    return declarations ? `.hwe-page-inner .${style.className} { ${declarations} }` : "";
  }

  private extractCssDeclarations(cssText: string): string {
    const css = cssText.trim();
    if (!css) return "";

    if (css.startsWith("{") && css.endsWith("}")) {
      return css.slice(1, -1).trim();
    }

    if (css.includes("{")) {
      const ruleMatch = /[^{]+\{([\s\S]*?)\}/.exec(css);
      return ruleMatch?.[1]?.trim() ?? "";
    }

    return css;
  }

  private styleDefinesFontFamily(style: ParagraphStyleDefinition): boolean {
    return this.extractFontFamilyNames(style.cssText).length > 0;
  }

  private clearInlineFontFamily(block: HTMLElement): void {
    [block, ...Array.from(block.querySelectorAll<HTMLElement>("*"))].forEach((element) => {
      if (element instanceof HTMLFontElement) {
        element.removeAttribute("face");
      }
      if (!element.hasAttribute("style")) return;

      element.style.removeProperty("font-family");
      element.style.removeProperty("font");
      if (!element.getAttribute("style")?.trim()) {
        element.removeAttribute("style");
      }
    });
  }

  private formatFontFaceCss(font: ParagraphFontFaceDefinition): string {
    const familyName = font.label.trim();
    const css = this.decodeCssEntities(font.cssText.trim());
    if (!familyName || !css) return "";

    return this.normalizeFontFaceCss(css, familyName);
  }

  private decodeCssEntities(css: string): string {
    return css
      .replace(/&quot;/g, "\"")
      .replace(/&#34;/g, "\"")
      .replace(/&#39;/g, "'")
      .replace(/&amp;/g, "&");
  }

  private isValidCssClassName(className: string): boolean {
    return /^-?[_a-zA-Z]+[_a-zA-Z0-9-]*$/.test(className);
  }

  private isFontFaceCss(cssText: string): boolean {
    return /@font-face\b/i.test(cssText);
  }

  private filterFontsForStyles(
    fonts: ParagraphFontFaceDefinition[],
    referencedFontFamilies: Set<string>
  ): ParagraphFontFaceDefinition[] {
    if (referencedFontFamilies.size === 0) return [];

    return fonts.filter((font) =>
      referencedFontFamilies.has(this.normalizeFontFamilyName(font.label))
    );
  }

  private getReferencedFontFamilies(styles: ParagraphStyleDefinition[]): Set<string> {
    const families = new Set<string>();
    styles.forEach((style) => {
      this.extractFontFamilyNames(style.cssText).forEach((family) => families.add(family));
    });
    return families;
  }

  private extractFontFamilyNames(cssText: string): string[] {
    const declarations = this.extractCssDeclarations(cssText);
    if (!declarations) return [];

    const families: string[] = [];
    const fontFamilyRegex = /\bfont-family\s*:\s*([^;{}]+)/gi;
    let match: RegExpExecArray | null;

    while ((match = fontFamilyRegex.exec(declarations)) !== null) {
      this.splitFontFamilyList(match[1]).forEach((family) => {
        const normalized = this.normalizeFontFamilyName(family);
        if (normalized && !GENERIC_FONT_FAMILIES.has(normalized)) {
          families.push(normalized);
        }
      });
    }

    return families;
  }

  private splitFontFamilyList(value: string): string[] {
    const families: string[] = [];
    let current = "";
    let quote: "\"" | "'" | null = null;
    let escaped = false;

    for (const char of value) {
      if (escaped) {
        current += char;
        escaped = false;
        continue;
      }

      if (char === "\\") {
        current += char;
        escaped = true;
        continue;
      }

      if (quote) {
        current += char;
        if (char === quote) quote = null;
        continue;
      }

      if (char === "\"" || char === "'") {
        quote = char;
        current += char;
        continue;
      }

      if (char === ",") {
        families.push(current);
        current = "";
        continue;
      }

      current += char;
    }

    families.push(current);
    return families;
  }

  private normalizeFontFamilyName(value: string): string {
    const trimmed = value.replace(/\s*!important\s*$/i, "").trim();
    if (!trimmed) return "";

    const quote = trimmed[0];
    const unquoted =
      (quote === "\"" || quote === "'") && trimmed.endsWith(quote)
        ? trimmed.slice(1, -1)
        : trimmed;

    return unquoted.replace(/\\(["'\\])/g, "$1").trim().toLowerCase();
  }

  private normalizeFontFaceCss(css: string, familyName: string): string {
    const familyDeclaration = `font-family: "${this.escapeCssString(familyName)}";`;
    const trimmed = css.trim();

    if (!this.isFontFaceCss(trimmed)) {
      return `@font-face { ${familyDeclaration} ${this.unwrapCssBlock(trimmed)} }`;
    }

    return trimmed.replace(/(@font-face\s*\{)([\s\S]*?)(\})/i, (_match, start, body, end) => {
      const normalizedBody = this.setFontFaceFamily(body, familyDeclaration);
      return `${start} ${normalizedBody} ${end}`;
    });
  }

  private setFontFaceFamily(body: string, familyDeclaration: string): string {
    const trimmed = body.trim();
    if (!trimmed) return familyDeclaration;

    if (/\bfont-family\s*:/i.test(trimmed)) {
      return trimmed.replace(/\bfont-family\s*:\s*[^;{}]+;?/i, `${familyDeclaration} `).trim();
    }

    return `${familyDeclaration} ${trimmed}`;
  }

  private unwrapCssBlock(css: string): string {
    return css.startsWith("{") && css.endsWith("}") ? css.slice(1, -1).trim() : css;
  }

  private escapeCssString(value: string): string {
    return value.replace(/\\/g, "\\\\").replace(/"/g, "\\\"");
  }

  private dedupeFonts(fonts: ParagraphFontFaceDefinition[]): ParagraphFontFaceDefinition[] {
    const seenLabels = new Set<string>();
    const seenCss = new Set<string>();
    return fonts.filter((font) => {
      const cssText = font.cssText.trim();
      if (!cssText) return false;
      const label = font.label.trim().toLowerCase();
      if ((label && seenLabels.has(label)) || seenCss.has(cssText)) return false;
      if (label) seenLabels.add(label);
      seenCss.add(cssText);
      return true;
    });
  }
}
