const PAGE_CONTENT_HEIGHT_PX = 929;
const PAGE_MERGE_THRESHOLD_PX = PAGE_CONTENT_HEIGHT_PX * 0.85;

export type PageFactory = (html?: string) => HTMLElement;
export type OnPagesChanged = (pages: HTMLElement[]) => void;

export class Paginator {
  private pages: HTMLElement[] = [];
  private readonly observer: ResizeObserver;
  private readonly pageFactory: PageFactory;
  private readonly onPagesChanged: OnPagesChanged;
  private rebalancing = false;

  constructor(pageFactory: PageFactory, onPagesChanged: OnPagesChanged) {
    this.pageFactory = pageFactory;
    this.onPagesChanged = onPagesChanged;

    this.observer = new ResizeObserver((entries) => {
      if (this.rebalancing) return;

      for (const entry of entries) {
        const inner = entry.target as HTMLElement;
        const page = inner.parentElement;
        if (page) this.rebalanceFromPage(page);
      }
    });
  }

  setPages(pages: HTMLElement[]): void {
    this.pages.forEach((page) => {
      const inner = this.getInner(page);
      if (inner) this.observer.unobserve(inner);
    });

    this.pages = [...pages];

    this.pages.forEach((page) => {
      const inner = this.getInner(page);
      if (inner) this.observer.observe(inner);
    });

    this.onPagesChanged(this.pages);
  }

  getPages(): HTMLElement[] {
    return [...this.pages];
  }

  rebalanceFromPage(page: HTMLElement): void {
    const index = this.pages.indexOf(page);
    if (index === -1 || this.rebalancing) return;

    this.rebalancing = true;
    try {
      this.resolveOverflow(index);
      this.resolveMerge(index);
      this.onPagesChanged(this.pages);
    } finally {
      this.rebalancing = false;
    }
  }

  destroy(): void {
    this.observer.disconnect();
  }

  private resolveOverflow(index: number): void {
    let safety = 0;

    while (safety++ < 200) {
      const page = this.pages[index];
      const inner = page ? this.getInner(page) : null;
      if (!inner || inner.scrollHeight <= PAGE_CONTENT_HEIGHT_PX + 1) break;

      const lastChild = this.getLastMeaningfulChild(inner);
      if (!lastChild) break;

      const nextPage = this.getOrCreateNextPage(index);
      const nextInner = this.getInner(nextPage);
      if (!nextInner) break;

      if (
        lastChild.nodeType === Node.ELEMENT_NODE &&
        this.unwrapIfSplittableContainer(lastChild as HTMLElement)
      ) {
        continue;
      }

      if (
        lastChild.nodeType === Node.ELEMENT_NODE &&
        (lastChild as HTMLElement).tagName === "TABLE"
      ) {
        const split = this.splitTable(lastChild as HTMLElement, nextInner, page);
        if (split) continue;
      }

      nextInner.insertBefore(lastChild, nextInner.firstChild);
    }
  }

  private resolveMerge(index: number): void {
    let safety = 0;

    while (safety++ < 200) {
      const page = this.pages[index];
      const nextPage = this.pages[index + 1];
      const inner = page ? this.getInner(page) : null;
      const nextInner = nextPage ? this.getInner(nextPage) : null;

      if (!inner || !nextInner) break;
      if (inner.scrollHeight >= PAGE_MERGE_THRESHOLD_PX) break;

      const firstChild = this.getFirstMeaningfulChild(nextInner);
      if (!firstChild) {
        this.removePage(index + 1);
        break;
      }

      inner.appendChild(firstChild);

      if (inner.scrollHeight > PAGE_CONTENT_HEIGHT_PX + 1) {
        nextInner.insertBefore(firstChild, nextInner.firstChild);
        break;
      }

      if (!this.getFirstMeaningfulChild(nextInner)) {
        this.removePage(index + 1);
        break;
      }
    }
  }

  private getOrCreateNextPage(afterIndex: number): HTMLElement {
    const existingPage = this.pages[afterIndex + 1];
    if (existingPage) return existingPage;

    const newPage = this.pageFactory();
    this.pages.splice(afterIndex + 1, 0, newPage);

    const inner = this.getInner(newPage);
    if (inner) this.observer.observe(inner);

    return newPage;
  }

  private removePage(index: number): void {
    const page = this.pages[index];
    if (!page) return;

    const inner = this.getInner(page);
    if (inner) this.observer.unobserve(inner);

    const previous = page.previousElementSibling;
    if (previous?.classList.contains("hwe-page-divider")) previous.remove();

    page.remove();
    this.pages.splice(index, 1);
  }

  private getInner(page: HTMLElement): HTMLElement | null {
    return page.querySelector(".hwe-page-inner");
  }

  private splitTable(
    table: HTMLElement,
    targetInner: HTMLElement,
    page: HTMLElement
  ): boolean {
    const pageBottom = page.getBoundingClientRect().bottom;
    const rows = Array.from(table.querySelectorAll("tr")) as HTMLElement[];

    let splitRowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].getBoundingClientRect().bottom > pageBottom) {
        splitRowIndex = i;
        break;
      }
    }

    if (splitRowIndex <= 0) return false;

    const rowsForNext = rows.slice(splitRowIndex);
    if (rowsForNext.length === 0) return false;

    const newTable = document.createElement("table");
    Array.from(table.attributes).forEach((attr) => {
      newTable.setAttribute(attr.name, attr.value);
    });

    const colgroup = table.querySelector("colgroup");
    if (colgroup) newTable.appendChild(colgroup.cloneNode(true));

    const thead = table.querySelector("thead");
    if (thead) newTable.appendChild(thead.cloneNode(true));

    const tbody = document.createElement("tbody");
    rowsForNext.forEach((row) => tbody.appendChild(row));
    newTable.appendChild(tbody);

    targetInner.insertBefore(newTable, targetInner.firstChild);

    const remainingRows = table.querySelectorAll("tbody tr, tfoot tr");
    if (remainingRows.length === 0) table.remove();

    return true;
  }

  private getLastMeaningfulChild(container: HTMLElement): ChildNode | null {
    const children = this.getMeaningfulChildren(container);
    if (children.length > 1) return children[children.length - 1];

    const onlyChild = children[0];
    if (
      onlyChild?.nodeType === Node.ELEMENT_NODE &&
      this.isSplittableContainer(onlyChild as HTMLElement)
    ) {
      return onlyChild;
    }

    return null;
  }

  private getFirstMeaningfulChild(container: HTMLElement): ChildNode | null {
    return this.getMeaningfulChildren(container)[0] ?? null;
  }

  private getMeaningfulChildren(container: HTMLElement): ChildNode[] {
    return Array.from(container.childNodes).filter((node) => !this.isEmptyNode(node));
  }

  private unwrapIfSplittableContainer(element: HTMLElement): boolean {
    if (!this.isSplittableContainer(element) || !element.parentNode) return false;

    this.unwrapElement(element);
    return true;
  }

  private unwrapElement(element: HTMLElement): void {
    if (!element.parentNode) return;

    const parent = element.parentNode;
    while (element.firstChild) {
      parent.insertBefore(element.firstChild, element);
    }
    parent.removeChild(element);
  }

  private isSplittableContainer(element: HTMLElement): boolean {
    const splittableTags = new Set(["DIV", "SECTION", "ARTICLE", "MAIN", "BODY"]);
    if (!splittableTags.has(element.tagName)) return false;
    if (element.classList.contains("hwe-page") || element.classList.contains("hwe-page-inner")) {
      return false;
    }

    return this.getMeaningfulChildren(element).length > 0;
  }

  private isEmptyNode(node: ChildNode): boolean {
    if (node.nodeType === Node.COMMENT_NODE) {
      return true;
    }

    if (node.nodeType === Node.TEXT_NODE) {
      return (node.textContent ?? "").replace(/\u00a0/g, " ").trim() === "";
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return false;

    const el = node as HTMLElement;
    if (el.tagName === "BR") return true;
    if (["META", "LINK", "STYLE", "SCRIPT", "XML"].includes(el.tagName)) return true;
    if (el.querySelector("img, table, tr, td, th, video, canvas, svg")) return false;

    return (
      ["P", "DIV", "SECTION", "ARTICLE", "SPAN"].includes(el.tagName) &&
      (el.textContent ?? "").replace(/\u00a0/g, " ").trim() === "" &&
      Array.from(el.childNodes).every((child) => this.isEmptyNode(child))
    );
  }
}
