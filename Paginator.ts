const PAGE_CONTENT_HEIGHT_PX = 929;
const PAGE_MERGE_THRESHOLD_PX = PAGE_CONTENT_HEIGHT_PX * 0.75;

export type PageFactory = (html?: string) => HTMLElement;
export type OnPagesChanged = (pages: HTMLElement[]) => void;

export class Paginator {
  private pages: HTMLElement[] = [];
  private readonly pageFactory: PageFactory;
  private readonly onPagesChanged: OnPagesChanged;
  private rebalancing = false;

  constructor(pageFactory: PageFactory, onPagesChanged: OnPagesChanged) {
    this.pageFactory = pageFactory;
    this.onPagesChanged = onPagesChanged;
  }

  setPages(pages: HTMLElement[]): void {
    this.pages = [...pages];
    this.onPagesChanged(this.pages);
  }

  getPages(): HTMLElement[] {
    return [...this.pages];
  }

  repaginateAll(): void {
    if (this.rebalancing) return;

    this.rebalancing = true;
    try {
      this.rebalanceFromIndex(0);
      this.removeEmptyPages();
      this.onPagesChanged(this.pages);
    } finally {
      this.rebalancing = false;
    }
  }

  rebalanceFromPage(page: HTMLElement): void {
    if (this.rebalancing) return;

    const index = this.pages.indexOf(page);
    if (index === -1) return;

    this.rebalancing = true;
    try {
      this.rebalanceFromIndex(index);
      this.removeEmptyPages();
      this.onPagesChanged(this.pages);
    } finally {
      this.rebalancing = false;
    }
  }

  destroy(): void {
    this.pages = [];
  }

  private rebalanceFromIndex(startIndex: number): void {
    let index = Math.max(0, startIndex);
    let safety = 0;

    while (index < this.pages.length && safety++ < 300) {
      this.resolveOverflow(index);
      this.resolveMerge(index);
      index++;
    }
  }

  private resolveOverflow(index: number): void {
    let safety = 0;

    while (safety++ < 300) {
      const page = this.pages[index];
      const inner = this.getInner(page);
      if (!inner || !this.pageOverflows(page)) break;

      const nextPage = this.getOrCreateNextPage(index);
      const nextInner = this.getInner(nextPage);
      if (!nextInner) break;

      const moved = this.moveOverflowPiece(inner, nextInner, page);
      if (!moved) break;
    }
  }

  private resolveMerge(index: number): void {
    let safety = 0;

    while (safety++ < 100) {
      const page = this.pages[index];
      const nextPage = this.pages[index + 1];
      const inner = this.getInner(page);
      const nextInner = nextPage ? this.getInner(nextPage) : null;

      if (!inner || !nextInner) break;
      if (inner.scrollHeight >= PAGE_MERGE_THRESHOLD_PX) break;

      const firstChild = this.getFirstMeaningfulChild(nextInner);
      if (!firstChild) {
        this.removePage(index + 1);
        break;
      }

      const targetForMerge = this.findMergeTarget(inner, firstChild);
      if (targetForMerge) {
        const movedChildren = this.moveAllChildren(firstChild as HTMLElement, targetForMerge);
        firstChild.parentNode?.removeChild(firstChild);

        if (this.pageOverflows(page)) {
          movedChildren.reverse().forEach((child) => {
            (firstChild as HTMLElement).insertBefore(child, (firstChild as HTMLElement).firstChild);
          });
          nextInner.insertBefore(firstChild, nextInner.firstChild);
          break;
        }
      } else {
        inner.appendChild(firstChild);

        if (this.pageOverflows(page)) {
          nextInner.insertBefore(firstChild, nextInner.firstChild);
          break;
        }
      }

      if (!this.getFirstMeaningfulChild(nextInner)) {
        this.removePage(index + 1);
        break;
      }
    }
  }

  private moveOverflowPiece(
    sourceInner: HTMLElement,
    targetInner: HTMLElement,
    page: HTMLElement
  ): boolean {
    const lastChild = this.getLastMeaningfulChild(sourceInner);
    if (!lastChild) return false;

    if (
      lastChild.nodeType === Node.ELEMENT_NODE &&
      this.unwrapIfSplittableContainer(lastChild as HTMLElement)
    ) {
      return true;
    }

    if (
      lastChild.nodeType === Node.ELEMENT_NODE &&
      (lastChild as HTMLElement).tagName === "TABLE"
    ) {
      return this.splitTable(lastChild as HTMLElement, targetInner, page);
    }

    if (
      lastChild.nodeType === Node.ELEMENT_NODE &&
      this.splitTextBlock(lastChild as HTMLElement, targetInner, page)
    ) {
      return true;
    }

    if (this.getMeaningfulChildren(sourceInner).length <= 1) {
      return false;
    }

    targetInner.insertBefore(lastChild, targetInner.firstChild);
    return true;
  }

  private getOrCreateNextPage(afterIndex: number): HTMLElement {
    const existingPage = this.pages[afterIndex + 1];
    if (existingPage) return existingPage;

    const newPage = this.pageFactory();
    this.pages.splice(afterIndex + 1, 0, newPage);
    return newPage;
  }

  private removePage(index: number): void {
    const page = this.pages[index];
    if (!page) return;

    const previous = page.previousElementSibling;
    const next = page.nextElementSibling;

    if (previous?.classList.contains("hwe-page-divider")) {
      previous.remove();
    } else if (next?.classList.contains("hwe-page-divider")) {
      next.remove();
    }

    page.remove();
    this.pages.splice(index, 1);
  }

  private removeEmptyPages(): void {
    if (this.pages.length <= 1) return;

    for (let index = this.pages.length - 1; index >= 0; index--) {
      const page = this.pages[index];
      const inner = this.getInner(page);
      if (inner && !this.getFirstMeaningfulChild(inner)) {
        this.removePage(index);
      }
    }

    if (this.pages.length === 0) {
      this.pages.push(this.pageFactory("<p><br></p>"));
    }
  }

  private splitTable(
    table: HTMLElement,
    targetInner: HTMLElement,
    page: HTMLElement
  ): boolean {
    const rows = Array.from(table.querySelectorAll("tr")) as HTMLElement[];
    if (rows.length <= 1) return false;

    const pageBottom = page.getBoundingClientRect().bottom;
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

  private splitTextBlock(
    block: HTMLElement,
    targetInner: HTMLElement,
    page: HTMLElement
  ): boolean {
    if (!this.isSplittableTextBlock(block)) return false;

    const overflowBlock = block.cloneNode(false) as HTMLElement;
    targetInner.insertBefore(overflowBlock, targetInner.firstChild);

    let movedAny = false;
    let safety = 0;

    while (this.pageOverflows(page) && safety++ < 300) {
      const moved = this.moveLastInlinePiece(block, overflowBlock);
      if (!moved) break;
      movedAny = true;

      if (this.isEmptyNode(block)) {
        block.remove();
        break;
      }
    }

    if (!movedAny || this.isEmptyNode(overflowBlock)) {
      overflowBlock.remove();
      return false;
    }

    return true;
  }

  private moveLastInlinePiece(source: HTMLElement, target: HTMLElement): boolean {
    const child = source.lastChild;
    if (!child) return false;

    if (this.isEmptyNode(child)) {
      child.remove();
      return true;
    }

    if (child.nodeType === Node.TEXT_NODE) {
      return this.moveLastWordFromTextNode(child as Text, target);
    }

    if (child.nodeType !== Node.ELEMENT_NODE) {
      target.insertBefore(child, target.firstChild);
      return true;
    }

    const element = child as HTMLElement;
    if (this.isAtomicElement(element)) {
      target.insertBefore(element, target.firstChild);
      return true;
    }

    const clone = element.cloneNode(false) as HTMLElement;
    const moved = this.moveLastInlinePiece(element, clone);

    if (!moved) {
      target.insertBefore(element, target.firstChild);
      return true;
    }

    target.insertBefore(clone, target.firstChild);
    if (this.isEmptyNode(element)) element.remove();
    return true;
  }

  private moveLastWordFromTextNode(textNode: Text, target: HTMLElement): boolean {
    const text = textNode.textContent ?? "";
    const trimmedEnd = text.replace(/\s+$/g, "");

    if (!trimmedEnd) {
      textNode.remove();
      return true;
    }

    const match = /(\s*\S+)$/.exec(trimmedEnd);
    if (!match) return false;

    if (match.index <= 0) {
      if (target.childNodes.length > 0) return false;
      target.insertBefore(textNode, target.firstChild);
      return true;
    }

    const prefix = trimmedEnd.slice(0, match.index);
    const suffix = trimmedEnd.slice(match.index);

    textNode.textContent = prefix;
    target.insertBefore(document.createTextNode(suffix), target.firstChild);
    return true;
  }

  private findMergeTarget(inner: HTMLElement, nodeFromNext: ChildNode): HTMLElement | null {
    if (nodeFromNext.nodeType !== Node.ELEMENT_NODE) return null;

    const incoming = nodeFromNext as HTMLElement;
    const last = this.getLastMeaningfulChild(inner);
    if (!last || last.nodeType !== Node.ELEMENT_NODE) return null;

    const current = last as HTMLElement;
    if (!this.isSplittableTextBlock(current)) return null;
    if (current.tagName !== incoming.tagName) return null;

    return current;
  }

  private moveAllChildren(source: HTMLElement, target: HTMLElement): ChildNode[] {
    const moved: ChildNode[] = [];

    while (source.firstChild) {
      const child = source.firstChild;
      moved.push(child);
      target.appendChild(child);
    }

    return moved;
  }

  private getInner(page: HTMLElement | undefined): HTMLElement | null {
    return page?.querySelector(".hwe-page-inner") ?? null;
  }

  private pageOverflows(page: HTMLElement): boolean {
    const inner = this.getInner(page);
    if (!inner) return false;
    return inner.scrollHeight > PAGE_CONTENT_HEIGHT_PX + 1;
  }

  private getLastMeaningfulChild(container: HTMLElement): ChildNode | null {
    const children = this.getMeaningfulChildren(container);
    if (children.length > 1) return children[children.length - 1];

    const onlyChild = children[0];
    if (
      onlyChild?.nodeType === Node.ELEMENT_NODE &&
      (this.isSplittableContainer(onlyChild as HTMLElement) ||
        this.isSplittableTextBlock(onlyChild as HTMLElement))
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

    const parent = element.parentNode;
    while (element.firstChild) {
      parent.insertBefore(element.firstChild, element);
    }
    parent.removeChild(element);
    return true;
  }

  private isSplittableContainer(element: HTMLElement): boolean {
    const splittableTags = new Set(["DIV", "SECTION", "ARTICLE", "MAIN", "BODY"]);
    if (!splittableTags.has(element.tagName)) return false;
    if (element.classList.contains("hwe-page") || element.classList.contains("hwe-page-inner")) {
      return false;
    }

    return this.getMeaningfulChildren(element).length > 0;
  }

  private isSplittableTextBlock(element: HTMLElement): boolean {
    const splittableTags = new Set([
      "P",
      "LI",
      "H1",
      "H2",
      "H3",
      "H4",
      "H5",
      "H6",
      "BLOCKQUOTE",
      "PRE",
    ]);

    if (!splittableTags.has(element.tagName)) return false;
    if (element.querySelector("table, tr, td, th")) return false;
    return (element.textContent ?? "").replace(/\u00a0/g, " ").trim().length > 0;
  }

  private isAtomicElement(element: HTMLElement): boolean {
    return ["IMG", "TABLE", "TR", "TD", "TH", "VIDEO", "CANVAS", "SVG", "BR"].includes(
      element.tagName
    );
  }

  private isEmptyNode(node: ChildNode): boolean {
    if (node.nodeType === Node.COMMENT_NODE) return true;

    if (node.nodeType === Node.TEXT_NODE) {
      return (node.textContent ?? "").replace(/\u00a0/g, " ").trim() === "";
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return false;

    const element = node as HTMLElement;
    if (element.tagName === "BR") return true;
    if (["META", "LINK", "STYLE", "SCRIPT", "XML"].includes(element.tagName)) return true;
    if (element.querySelector("img, table, tr, td, th, video, canvas, svg")) return false;

    return (
      ["P", "DIV", "SECTION", "ARTICLE", "SPAN"].includes(element.tagName) &&
      (element.textContent ?? "").replace(/\u00a0/g, " ").trim() === "" &&
      Array.from(element.childNodes).every((child) => this.isEmptyNode(child))
    );
  }
}
