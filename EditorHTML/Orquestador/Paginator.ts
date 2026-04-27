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

export type PageFactory = (html?: string) => HTMLElement;
export type OnPagesChanged = (pages: HTMLElement[]) => void;

export class Paginator {
  private pages: HTMLElement[] = [];
  private observer!: ResizeObserver;
  private pageFactory: PageFactory;
  private onPagesChanged: OnPagesChanged;
  private rebalancing = false;
  private getContentHeight(inner: HTMLElement): number {
    const range = document.createRange();
    range.selectNodeContents(inner);
    return range.getBoundingClientRect().height;
  }

  private getMaxContentHeight(page: HTMLElement, inner: HTMLElement): number {
    const pageHeight = page.clientHeight;

    const styles = getComputedStyle(inner);
    const paddingTop = parseFloat(styles.paddingTop);
    const paddingBottom = parseFloat(styles.paddingBottom);

    return pageHeight - paddingTop - paddingBottom;
  }

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

  private splitTextNodeByLine(
    textNode: Text,
    maxHeight: number,
    container: HTMLElement
  ): Text | null {
    const text = textNode.textContent ?? "";
    if (!text.trim()) return null;

    let low = 0;
    let high = text.length;
    let best = 0;

    const range = document.createRange();

    while (low <= high) {
      const mid = Math.floor((low + high) / 2);

      range.setStart(textNode, 0);
      range.setEnd(textNode, mid);

      const rects = Array.from(range.getClientRects());
      const lastRect = rects[rects.length - 1];
      const containerRect = container.getBoundingClientRect();

      if (!lastRect) break;

      const usedHeight = lastRect.bottom - containerRect.top;

      if (usedHeight <= maxHeight) {
        best = mid;
        low = mid + 1;
      } else {
        high = mid - 1;
      }
    }

    if (best === 0 || best >= text.length) return null;

    const overflowText = text.slice(best);
    textNode.textContent = text.slice(0, best);

    return document.createTextNode(overflowText);
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

    const nextPage = this.pages[index + 1];
    if (nextPage) {
      this.resolveOverflow(index + 1);
    }
  }

  private resolveOverflow(index: number): void {
    let safety = 0;
    while (safety++ < 200) {
      const page = this.pages[index];
      if (!page) break;

      const inner = this.getInner(page);
      if (!inner) break;

      const maxHeight = this.getMaxContentHeight(page, inner);
      if (this.getContentHeight(inner) <= maxHeight) break;

      const lastChild = this.getLastMeaningfulChild(inner);
      if (!lastChild) break;

      const nextPage = this.getOrCreateNextPage(index);
      const nextInner = this.getInner(nextPage);
      if (!nextInner) break;

      if (lastChild.nodeType === Node.TEXT_NODE) {
        const overflowText = this.splitTextNodeByLine(
          lastChild as Text,
          maxHeight,
          inner
        );

        if (overflowText) {
          nextInner.insertBefore(overflowText, nextInner.firstChild);
          break;
        }
      }

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

      const maxHeight = this.getMaxContentHeight(page, inner);
      if (this.getContentHeight(inner) >= maxHeight) break;

      const firstChild = this.getFirstMeaningfulChild(nextInner);
      if (!firstChild) {
        this.removePage(index + 1);
        break;
      }

      inner.appendChild(firstChild);

      if (inner.offsetHeight > maxHeight) {
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

    public rebalanceFromPage(page: HTMLElement): void {
    const index = this.pages.indexOf(page);
    if (index === -1) return;

    this.rebalancePage(page);

    const next = this.pages[index + 1];
    if (next) {
      this.rebalancePage(next);
    }
  }

}