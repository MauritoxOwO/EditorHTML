import { ParagraphStyleOption, Toolbar } from "../Resize/Toolbar";
import { ParagraphStyleDefinition } from "../execCommand/styleApi";

export class ParagraphStyleManager {
  private readonly classNames = new Set<string>();
  private styleElement: HTMLStyleElement | null = null;

  constructor(
    private readonly rootProvider: () => HTMLElement,
    private readonly toolbar: Toolbar
  ) {}

  get cssText(): string {
    return this.styleElement?.textContent ?? "";
  }

  hasClass(className: string): boolean {
    return this.classNames.has(className);
  }

  setStyles(styles: ParagraphStyleDefinition[]): void {
    const validStyles = styles.filter((style) => this.isValidCssClassName(style.className));

    this.classNames.clear();
    validStyles.forEach((style) => this.classNames.add(style.className));
    this.toolbar.setParagraphStyles(
      validStyles.map<ParagraphStyleOption>((style) => ({
        label: style.label,
        className: style.className,
      }))
    );
    this.injectParagraphStyleCss(validStyles);
  }

  applyToBlock(block: HTMLElement, className: string | null): void {
    this.classNames.forEach((styleClass) => block.classList.remove(styleClass));
    if (className) block.classList.add(className);
  }

  private injectParagraphStyleCss(styles: ParagraphStyleDefinition[]): void {
    if (!this.styleElement) {
      this.styleElement = document.createElement("style");
      this.styleElement.setAttribute("data-hwe-style-catalog", "true");
      this.rootProvider().insertBefore(this.styleElement, this.rootProvider().firstChild);
    }

    this.styleElement.textContent = styles
      .map((style) => this.formatParagraphStyleCss(style))
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

  private isValidCssClassName(className: string): boolean {
    return /^-?[_a-zA-Z]+[_a-zA-Z0-9-]*$/.test(className);
  }
}