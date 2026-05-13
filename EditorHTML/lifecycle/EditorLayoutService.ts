import { getMeaningfulChildren } from "../dom/EditableDom";

export class EditorLayoutService {
  applyOfficialTableWidths(root: HTMLElement): void {
    const inners = root.classList.contains("hwe-page-inner")
      ? [root]
      : Array.from(root.querySelectorAll<HTMLElement>(".hwe-page-inner"));

    inners.forEach((inner) => {
      inner.querySelectorAll<HTMLElement>(".hwe-table-flow-wrapper").forEach((wrapper) => {
        if (!this.getDirectFlowTable(wrapper)) wrapper.classList.remove("hwe-table-flow-wrapper");
      });

      Array.from(inner.children).forEach((child) => {
        const table = this.getDirectFlowTable(child as HTMLElement);
        if (!table) return;

        if ((child as HTMLElement) !== table) {
          (child as HTMLElement).classList.add("hwe-table-flow-wrapper");
        }
        table.style.setProperty("width", "100%", "important");
        table.style.setProperty("max-width", "100%", "important");
        table.style.setProperty("margin-left", "0", "important");
        table.style.setProperty("margin-right", "0", "important");
      });
    });
  }

  pageOverflows(page: HTMLElement): boolean {
    const inner = page.querySelector<HTMLElement>(".hwe-page-inner");
    if (!inner) return false;

    const scrollOverflows = inner.scrollHeight > inner.clientHeight + 1;
    const contentBottom = this.getContentBottom(inner);
    if (contentBottom === null) return scrollOverflows;

    return contentBottom > this.getContentLimitBottom(inner) + 1 || scrollOverflows;
  }

  private getDirectFlowTable(element: HTMLElement): HTMLTableElement | null {
    if (element.tagName === "TABLE") return element as HTMLTableElement;

    const children = getMeaningfulChildren(element, true);
    if (children.length !== 1 || children[0].nodeType !== Node.ELEMENT_NODE) return null;

    const onlyChild = children[0] as HTMLElement;
    return onlyChild.tagName === "TABLE" ? (onlyChild as HTMLTableElement) : null;
  }

  private getContentLimitBottom(inner: HTMLElement): number {
    const styles = getComputedStyle(inner);
    const paddingBottom = parseFloat(styles.paddingBottom) || 0;
    return inner.getBoundingClientRect().bottom - Math.min(paddingBottom, 12);
  }

  private getContentBottom(inner: HTMLElement): number | null {
    const children = getMeaningfulChildren(inner);
    if (children.length === 0) return null;

    return children.reduce<number | null>((bottom, child) => {
      const childBottom = this.getNodeBottom(child);
      if (childBottom === null) return bottom;
      return bottom === null ? childBottom : Math.max(bottom, childBottom);
    }, null);
  }

  private getNodeBottom(node: ChildNode): number | null {
    if (node.nodeType === Node.ELEMENT_NODE) {
      return (node as HTMLElement).getBoundingClientRect().bottom;
    }

    const range = document.createRange();
    range.selectNodeContents(node);
    const rect = range.getBoundingClientRect();
    return rect.width > 0 || rect.height > 0 ? rect.bottom : null;
  }
}