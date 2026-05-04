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
