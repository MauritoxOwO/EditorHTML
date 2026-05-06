export type PageFactory = (html?: string) => HTMLElement;
export type OnPagesChanged = (pages: HTMLElement[]) => void;

const MIN_LINES_ON_SPLIT = 2;
const LINE_TOP_TOLERANCE_PX = 2;

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
      this.repaginateFromIndex(0);
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
      this.repaginateFromIndex(Math.max(0, index - 1));
      this.onPagesChanged(this.pages);
    } finally {
      this.rebalancing = false;
    }
  }

  destroy(): void {
    this.pages = [];
  }

  private repaginateFromIndex(startIndex: number): void {
    const safeStart = Math.max(0, Math.min(startIndex, this.pages.length - 1));
    const nodes = this.collectNodesFromIndex(safeStart);

    this.trimPagesFromIndex(safeStart);

    let pageIndex = safeStart;
    for (const node of nodes) {
      pageIndex = this.appendNodeFlowing(node, pageIndex);
    }

    this.stabilizeOverflow();
    this.removeEmptyPages();
  }

  private stabilizeOverflow(): void {
    let safety = 0;

    while (safety++ < 20) {
      const overflowingIndex = this.pages.findIndex((page) => this.pageOverflows(page));
      if (overflowingIndex === -1) break;

      this.resolveOverflow(overflowingIndex);
    }
  }

  private collectNodesFromIndex(startIndex: number): ChildNode[] {
    const nodes: ChildNode[] = [];

    this.pages.slice(startIndex).forEach((page) => {
      const inner = this.getInner(page);
      if (!inner) return;

      nodes.push(...this.collectFlowNodes(inner));
    });

    return nodes;
  }

  private collectFlowNodes(container: HTMLElement): ChildNode[] {
    const nodes: ChildNode[] = [];
    let inlineParagraph: HTMLElement | null = null;

    const flushInlineParagraph = () => {
      if (!inlineParagraph) return;

      if (this.getMeaningfulChildren(inlineParagraph).length > 0) {
        nodes.push(inlineParagraph);
      }

      inlineParagraph = null;
    };

    Array.from(container.childNodes).forEach((node) => {
      if (this.isEmptyNode(node)) return;

      if (this.isInlineFlowNode(node)) {
        if (!inlineParagraph) inlineParagraph = document.createElement("p");
        inlineParagraph.appendChild(node);
        return;
      }

      flushInlineParagraph();
      nodes.push(...this.flattenFlowNode(node));
    });

    flushInlineParagraph();
    return nodes;
  }

  private flattenFlowNode(node: ChildNode): ChildNode[] {
    if (node.nodeType === Node.TEXT_NODE) {
      const paragraph = document.createElement("p");
      paragraph.appendChild(node);
      return [paragraph];
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return [node];

    const element = node as HTMLElement;
    if (!this.isSplittableContainer(element)) return [node];

    return this.collectFlowNodes(element);
  }

  private trimPagesFromIndex(startIndex: number): void {
    const keepBefore = this.pages.slice(0, startIndex);
    let page = this.pages[startIndex];

    if (!page) {
      page = this.pageFactory();
    }

    const inner = this.getInner(page);
    if (inner) inner.innerHTML = "";

    this.pages.slice(startIndex + 1).forEach((extraPage) => extraPage.remove());
    this.pages = [...keepBefore, page];
  }

  private appendNodeFlowing(node: ChildNode, pageIndex: number): number {
    let currentIndex = pageIndex;
    let currentPage = this.pages[currentIndex] ?? this.getOrCreateNextPage(currentIndex - 1);
    let currentInner = this.getInner(currentPage);
    if (!currentInner) return currentIndex;

    if (node.nodeType === Node.ELEMENT_NODE) {
      this.constrainAtomicElement(node as HTMLElement, currentPage);
    }

    const pageHadContent = !!this.getFirstMeaningfulChild(currentInner);
    currentInner.appendChild(node);

    if (!this.pageOverflows(currentPage)) {
      return currentIndex;
    }

    if (pageHadContent) {
      currentInner.removeChild(node);
      const nextPage = this.getOrCreateNextPage(currentIndex);
      const nextInner = this.getInner(nextPage);
      if (!nextInner) return currentIndex;

      if (node.nodeType === Node.ELEMENT_NODE) {
        this.constrainAtomicElement(node as HTMLElement, nextPage);
      }

      nextInner.appendChild(node);
      if (this.pageOverflows(nextPage)) {
        this.resolveOverflow(currentIndex + 1);
      }
      return currentIndex + 1;
    }

    this.resolveOverflow(currentIndex);

    if (this.pageOverflows(currentPage)) {
      return currentIndex;
    }

    const nextPage = this.pages[currentIndex + 1];
    if (!nextPage) return currentIndex;

    return currentIndex + 1;
  }

  private constrainAtomicElement(element: HTMLElement, page: HTMLElement): void {
    if (element.tagName === "IMG") {
      const contentHeight = this.getContentHeight(page);
      element.style.setProperty("max-width", "100%", "important");
      element.style.setProperty("max-height", `${contentHeight}px`, "important");
      element.style.setProperty("height", "auto", "important");
      element.style.setProperty("object-fit", "contain");
    }
  }

  private getContentHeight(page: HTMLElement): number {
    const inner = this.getInner(page);
    if (!inner) return page.clientHeight;

    const styles = getComputedStyle(inner);
    const paddingTop = parseFloat(styles.paddingTop) || 0;
    const paddingBottom = parseFloat(styles.paddingBottom) || 0;
    return Math.max(0, inner.clientHeight - paddingTop - paddingBottom);
  }

  private resolveOverflow(index: number): void {
    let currentIndex = index;
    let safety = 0;

    while (currentIndex < this.pages.length && safety++ < 500) {
      const page = this.pages[currentIndex];
      const inner = this.getInner(page);

      if (!inner || !this.pageOverflows(page)) {
        currentIndex++;
        continue;
      }

      const nextPage = this.getOrCreateNextPage(currentIndex);
      const nextInner = this.getInner(nextPage);
      if (!nextInner) break;

      const moved = this.moveOverflowPiece(inner, nextInner, page);
      if (!moved) {
        currentIndex++;
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
    this.onPagesChanged(this.pages);
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

    this.ensureMinimumOverflowLines(block, overflowBlock);

    return true;
  }

  private ensureMinimumOverflowLines(source: HTMLElement, target: HTMLElement): void {
    let safety = 0;

    while (
      safety++ < 100 &&
      !this.isEmptyNode(source) &&
      this.countVisualLines(target) < MIN_LINES_ON_SPLIT
    ) {
      const moved = this.moveLastInlinePiece(source, target);
      if (!moved) break;

      if (this.isEmptyNode(source)) {
        source.remove();
        break;
      }
    }
  }

  private countVisualLines(element: HTMLElement): number {
    if (this.isEmptyNode(element)) return 0;

    const range = document.createRange();
    range.selectNodeContents(element);

    const tops: number[] = [];
    Array.from(range.getClientRects()).forEach((rect) => {
      if (rect.width <= 0 || rect.height <= 0) return;

      const hasExistingTop = tops.some(
        (top) => Math.abs(top - rect.top) <= LINE_TOP_TOLERANCE_PX
      );
      if (!hasExistingTop) tops.push(rect.top);
    });

    if (tops.length > 0) return tops.length;
    return (element.textContent ?? "").replace(/\u00a0/g, " ").trim() ? 1 : 0;
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

  private getInner(page: HTMLElement | undefined): HTMLElement | null {
    return page?.querySelector(".hwe-page-inner") ?? null;
  }

  private pageOverflows(page: HTMLElement): boolean {
    const inner = this.getInner(page);
    if (!inner) return false;
    const limit = inner.clientHeight || page.clientHeight;
    if (limit <= 0) return false;
    return inner.scrollHeight > limit + 1;
  }

  private getLastMeaningfulChild(container: HTMLElement): ChildNode | null {
    const children = this.getMeaningfulChildren(container);
    if (children.length > 1) return children[children.length - 1];

    const onlyChild = children[0];
    if (
      onlyChild?.nodeType === Node.ELEMENT_NODE &&
      ((onlyChild as HTMLElement).tagName === "TABLE" ||
        this.isSplittableContainer(onlyChild as HTMLElement) ||
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
    const splittableTags = new Set([
      "DIV",
      "SECTION",
      "ARTICLE",
      "MAIN",
      "BODY",
      "CENTER",
      "HEADER",
      "FOOTER",
      "ASIDE",
      "NAV",
    ]);
    if (!splittableTags.has(element.tagName)) return false;
    if (element.classList.contains("hwe-page") || element.classList.contains("hwe-page-inner")) {
      return false;
    }

    return this.getMeaningfulChildren(element).length > 0;
  }

  private isInlineFlowNode(node: ChildNode): boolean {
    if (node.nodeType === Node.TEXT_NODE) {
      return (node.textContent ?? "").replace(/\u00a0/g, " ").trim() !== "";
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return false;

    const element = node as HTMLElement;
    if (element.hasAttribute("data-hwe-caret")) return true;
    if (element.querySelector("img, table, tr, td, th, video, canvas, svg")) return false;

    return [
      "A",
      "ABBR",
      "B",
      "BDI",
      "BDO",
      "CITE",
      "CODE",
      "EM",
      "FONT",
      "I",
      "KBD",
      "MARK",
      "Q",
      "S",
      "SMALL",
      "SPAN",
      "STRONG",
      "SUB",
      "SUP",
      "TIME",
      "U",
      "VAR",
    ].includes(element.tagName);
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
    if (element.hasAttribute("data-hwe-caret")) return false;
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