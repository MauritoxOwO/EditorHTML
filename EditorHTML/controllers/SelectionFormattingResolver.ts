import { SelectionFormattingInterface } from "../ui/Toolbar";

interface SelectionFormattingResolverOptions {
  rootProvider: () => HTMLElement;
  isParagraphStyleClass: (className: string) => boolean;
}

export class SelectionFormattingResolver {
  constructor(private readonly options: SelectionFormattingResolverOptions) {}

  resolve(): SelectionFormattingInterface {
    const range = this.getRange();
    if (!range) return {};

    const element = this.getElementAtRangeStart(range);
    if (!element) return {};

    const computedStyle = window.getComputedStyle(element);
    return {
      fontSize: computedStyle.fontSize,
      fontFamily: this.getPrimaryFontFamily(computedStyle.fontFamily),
      paragraphStyle: this.getParagraphStyle(element),
    };
  }

  private getRange(): Range | null {
    const root = this.options.rootProvider();
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return null;

    const range = selection.getRangeAt(0);
    return root.contains(range.commonAncestorContainer) ? range : null;
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
