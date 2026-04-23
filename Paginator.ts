/**
 * Paginator.ts
 * 
 * Estrategia: cada .hwe-page tiene height fija (297mm) y overflow:hidden.
 * Dentro hay un div.hwe-page-inner sin restricción de altura.
 * ResizeObserver vigila ese inner div. Cuando su offsetHeight supera
 * el área de contenido disponible, se mueve el último hijo al inicio
 * de la página siguiente.
 */

// Área de contenido A4 a 96dpi menos márgenes 2×25.4mm
// 297mm * (96/25.4) = 1122px - 2 * 96.38px ≈ 929px
const PAGE_CONTENT_HEIGHT_PX = 929;
const PAGE_MERGE_THRESHOLD_PX = PAGE_CONTENT_HEIGHT_PX * 0.85;

export type PageFactory = (html?: string) => HTMLElement;
export type OnPagesChanged = (pages: HTMLElement[]) => void;

export class Paginator {
  private pages: HTMLElement[] = [];
  private observer!: ResizeObserver;
  private pageFactory: PageFactory;
  private onPagesChanged: OnPagesChanged;
  private rebalancing = false;

  constructor(pageFactory: PageFactory, onPagesChanged: OnPagesChanged) {
    this.pageFactory    = pageFactory;
    this.onPagesChanged = onPagesChanged;

    this.observer = new ResizeObserver((entries) => {
      if (this.rebalancing) return;
      for (const entry of entries) {
        // entry.target es el inner div — buscamos la página padre
        const inner = entry.target as HTMLElement;
        const page  = inner.parentElement;
        if (!page) continue;
        this.rebalancePage(page);
      }
    });
  }

  setPages(pages: HTMLElement[]): void {
    this.pages.forEach((p) => {
      const inner = this.getInner(p);
      if (inner) this.observer.unobserve(inner);
    });

    this.pages = [...pages];

    this.pages.forEach((p) => {
      const inner = this.getInner(p);
      if (inner) this.observer.observe(inner);
    });
  }

  getPages(): HTMLElement[] {
    return this.pages;
  }

  destroy(): void {
    this.observer.disconnect();
  }

  // Rebalanceo 

  private rebalancePage(page: HTMLElement): void {
    const index = this.pages.indexOf(page);
    if (index === -1) return;

    this.rebalancing = true;
    try {
      this.resolveOverflow(index);
      this.resolveMerge(index);
    } finally {
      this.rebalancing = false;
    }
  }

  private resolveOverflow(index: number): void {
    let safety = 0;
    while (safety++ < 200) {
      const page  = this.pages[index];
      if (!page) break;

      const inner = this.getInner(page);
      if (!inner) break;

      if (inner.offsetHeight <= PAGE_CONTENT_HEIGHT_PX) break;

      const lastChild = this.getLastMeaningfulChild(inner);
      if (!lastChild) break;

      const nextPage  = this.getOrCreateNextPage(index);
      const nextInner = this.getInner(nextPage);
      if (!nextInner) break;

      nextInner.insertBefore(lastChild, nextInner.firstChild);
    }
  }

  private resolveMerge(index: number): void {
    let safety = 0;
    while (safety++ < 200) {
      const page     = this.pages[index];
      const nextPage = this.pages[index + 1];
      if (!page || !nextPage) break;

      const inner     = this.getInner(page);
      const nextInner = this.getInner(nextPage);
      if (!inner || !nextInner) break;

      if (inner.offsetHeight >= PAGE_MERGE_THRESHOLD_PX) break;

      const firstChild = this.getFirstMeaningfulChild(nextInner);
      if (!firstChild) {
        this.removePage(index + 1);
        break;
      }

      inner.appendChild(firstChild);

      if (inner.offsetHeight > PAGE_CONTENT_HEIGHT_PX) {
        nextInner.insertBefore(firstChild, nextInner.firstChild);
        break;
      }

      if (!this.getFirstMeaningfulChild(nextInner)) {
        this.removePage(index + 1);
        break;
      }
    }
  }

  // Gestión de páginas 

  private getOrCreateNextPage(afterIndex: number): HTMLElement {
    if (this.pages[afterIndex + 1]) {
      return this.pages[afterIndex + 1];
    }
    const newPage = this.pageFactory();
    this.pages.splice(afterIndex + 1, 0, newPage);
    const inner = this.getInner(newPage);
    if (inner) this.observer.observe(inner);
    this.onPagesChanged(this.pages);
    return newPage;
  }

  private removePage(index: number): void {
    const page  = this.pages[index];
    if (!page) return;
    const inner = this.getInner(page);
    if (inner) this.observer.unobserve(inner);
    this.pages.splice(index, 1);
    this.onPagesChanged(this.pages);
  }

  // Helpers DOM 

  /** Obtiene el div interno de una página */
  private getInner(page: HTMLElement): HTMLElement | null {
    return page.querySelector(".hwe-page-inner");
  }

  private getLastMeaningfulChild(container: HTMLElement): ChildNode | null {
    const children = Array.from(container.childNodes).filter(
      (n) => !this.isEmptyNode(n)
    );
    if (children.length <= 1) return null;
    return children[children.length - 1];
  }

  private getFirstMeaningfulChild(container: HTMLElement): ChildNode | null {
    const children = Array.from(container.childNodes).filter(
      (n) => !this.isEmptyNode(n)
    );
    return children[0] ?? null;
  }

  private isEmptyNode(node: ChildNode): boolean {
    if (node.nodeType === Node.TEXT_NODE) {
      return (node.textContent ?? "").trim() === "";
    }
    if (node.nodeType === Node.ELEMENT_NODE) {
      const el = node as HTMLElement;
      if (el.tagName === "BR") return true;
      if (
        el.tagName === "P" &&
        el.childNodes.length <= 1 &&
        (el.textContent ?? "").trim() === ""
      ) return true;
    }
    return false;
  }
}
