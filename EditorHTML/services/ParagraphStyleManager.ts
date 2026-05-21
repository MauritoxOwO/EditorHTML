import { ParagraphStyleOption, Toolbar } from "../ui/Toolbar";
import {
  ParagraphFontFaceDefinition,
  ParagraphStyleCatalog,
  ParagraphStyleDefinition,
} from "./dataverse/styleApi";

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
    const fonts = this.dedupeFonts(catalog.fonts);
    const styles = catalog.styles;
    const validStyles = styles.filter((style) => this.isValidCssClassName(style.className));

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
    return /\bfont-family\s*:/i.test(this.extractCssDeclarations(style.cssText));
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
    return this.decodeCssEntities(font.cssText.trim());
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
