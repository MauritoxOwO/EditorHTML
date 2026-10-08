import { SelectionFormattingInterface } from "../ui/Toolbar";

interface SelectionFormattingResolverOptions {
  rootProvider: () => HTMLElement;
  isParagraphStyleClass: (className: string) => boolean;
  documentSelection: {
    readSelection(): { range: Range } | null;
    getSelectedTextTargets(range: Range): Array<{ node: Text }>;
    getSelectedParagraphs(range: Range): HTMLElement[];
  };
}

export class SelectionFormattingResolver {
  constructor(private readonly options: SelectionFormattingResolverOptions) {}

  resolve(): SelectionFormattingInterface {
    const range = this.getRange();
    if (!range) return {};

    const element = this.getElementAtRangeStart(range);
    if (!element) return {};
    const textElements = range.collapsed ? [element] : this.getSelectedTextElements(range);
    const styles = textElements.map((item) => window.getComputedStyle(item));
    const uniform = (values: string[]): string | undefined =>
      values.length > 0 && values.every((value) => value === values[0]) ? values[0] : undefined;
    const paragraphs = this.getSelectedParagraphs(range, element);
    const paragraphStyles = paragraphs.map((paragraph) => this.getParagraphStyle(paragraph));
    return {
      fontSize: uniform(styles.map((style) => style.fontSize)),
      fontFamily: uniform(styles.map((style) => this.getPrimaryFontFamily(style.fontFamily))),
      paragraphStyle: uniform(paragraphStyles),
      commandStates: {
        bold: this.resolveMixedState(styles.map((style) => Number.parseInt(style.fontWeight, 10) >= 600 || style.fontWeight === "bold")),
        italic: this.resolveMixedState(styles.map((style) => style.fontStyle === "italic" || style.fontStyle === "oblique")),
        underline: this.resolveMixedState(textElements.map((item) => this.isUnderlined(item))),
      },
    };
  }

  private resolveMixedState(values: boolean[]): boolean | "mixed" | undefined {
    if (!values.length) return undefined;
    return values.every((value) => value === values[0]) ? values[0] : "mixed";
  }

  private isUnderlined(element: HTMLElement): boolean {
    const inner = element.closest(".hwe-page-inner");
    for (let current: HTMLElement | null = element; current && current !== inner; current = current.parentElement) {
      if (current.tagName === "U" || current.tagName === "INS") return true;
      const style = window.getComputedStyle(current);
      const decoration = `${current.style.textDecoration} ${current.style.textDecorationLine} ${style.textDecoration} ${style.textDecorationLine}`.toLowerCase();
      if (decoration.includes("underline")) return true;
    }
    return false;
  }

  private getSelectedTextElements(range: Range): HTMLElement[] {
    return this.options.documentSelection.getSelectedTextTargets(range)
      .map(({ node }) => node.parentElement)
      .filter((element): element is HTMLElement => Boolean(element));
  }

  private getSelectedParagraphs(range: Range, fallback: HTMLElement): HTMLElement[] {
    if (range.collapsed) return [this.getParagraph(fallback) ?? fallback];
    return this.options.documentSelection.getSelectedParagraphs(range);
  }

  private getParagraph(element: HTMLElement): HTMLElement | null {
    return element.closest<HTMLElement>("p, div, li, h1, h2, h3, h4, h5, h6, blockquote, pre");
  }

  private getRange(): Range | null {
    const root = this.options.rootProvider();
    const sharedRange = this.options.documentSelection.readSelection()?.range;
    return sharedRange && root.contains(sharedRange.startContainer) && root.contains(sharedRange.endContainer)
      ? sharedRange : null;
  }

  private getElementAtRangeStart(range: Range): HTMLElement | null {
    let node = range.startContainer;

    if (node.nodeType === Node.ELEMENT_NODE) {
      node =
        node.childNodes[range.startOffset] ??
        node.childNodes[Math.max(0, range.startOffset - 1)] ??
        node;
    }

    return node.nodeType === Node.ELEMENT_NODE
      ? (node as HTMLElement)
      : node.parentElement;
  }

  private getPrimaryFontFamily(fontFamily: string): string {
    return (fontFamily.split(",")[0] ?? "")
      .trim()
      .replace(/^(['"])(.*)\1$/, "$2");
  }

  private getParagraphStyle(element: HTMLElement): string {
    const root = this.options.rootProvider();
    let current: HTMLElement | null = element;

    while (current && root.contains(current)) {
      const paragraphStyle = Array.from(current.classList).find((className) =>
        this.options.isParagraphStyleClass(className)
      );
      if (paragraphStyle) return paragraphStyle;
      current = current.parentElement;
    }

    return "";
  }
}
