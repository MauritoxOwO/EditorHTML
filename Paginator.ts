


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

  private getLastMeaningfulChild(container: HTMLElement): ChildNode | null {
    const children = Array.from(container.childNodes).filter(
      (node) => !this.isEmptyNode(node)
    );
    if (children.length <= 1) return null;
    return children[children.length - 1];
  }

  private getFirstMeaningfulChild(container: HTMLElement): ChildNode | null {
    return Array.from(container.childNodes).find((node) => !this.isEmptyNode(node)) ?? null;
  }

  private isEmptyNode(node: ChildNode): boolean {
    if (node.nodeType === Node.TEXT_NODE) {
      return (node.textContent ?? "").trim() === "";
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return false;

    const el = node as HTMLElement;
    if (el.tagName === "BR") return true;

    return (
      el.tagName === "P" &&
      el.childNodes.length <= 1 &&
      (el.textContent ?? "").trim() === ""
    );
  }
}
