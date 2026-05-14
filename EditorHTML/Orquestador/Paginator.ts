import { KeepTogetherController } from "./KeepTogetherController";
import { TablePaginator } from "./TablePaginator";
import { TextBlockSplitter } from "./TextBlockSplitter";
import { hweDebugLog, hweDebugStart } from "../debug/DebugLogger";
import {
  getContentHeight,
  getContentLimitBottom,
  getInner,
  getMeaningfulChildren,
  isEditableBlankBlock,
  isEmptyNode,
  isSplittableContainer as isFlowSplittableContainer,
  isSplittableTextBlock,
  isTableElement,
  isTableFlowWrapper,
  pageOverflows,
  shouldPreserveContainerShell,
} from "./PaginatorDom";

export type PageFactory = (html?: string) => HTMLElement;
export type OnPagesChanged = (pages: HTMLElement[]) => void;
export type OnPageCreated = (page: HTMLElement, afterPage: HTMLElement | null) => void;

export interface RebalanceOptions {
  includePreviousPage?: boolean;
  compactPages?: boolean;
  overflowOnly?: boolean;
}

export class Paginator {
  private static readonly MAX_STABILIZE_MS = 700;
  private static readonly MAX_RESOLVE_MS = 450;
  private static readonly MAX_SPLIT_MS = 180;

  private pages: HTMLElement[] = [];
  private readonly pageFactory: PageFactory;
  private readonly onPagesChanged: OnPagesChanged;
  private readonly onPageCreated?: OnPageCreated;
  private readonly keepTogetherController = new KeepTogetherController();
  private readonly tablePaginator = new TablePaginator();
  private readonly textBlockSplitter = new TextBlockSplitter();
  private rebalancing = false;

  constructor(
    pageFactory: PageFactory,
    onPagesChanged: OnPagesChanged,
    onPageCreated?: OnPageCreated
  ) {
    this.pageFactory = pageFactory;
    this.onPagesChanged = onPagesChanged;
    this.onPageCreated = onPageCreated;
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

    const done = hweDebugStart("paginator.repaginateAll", {
      pages: this.pages.length,
    });
    this.rebalancing = true;
    try {
      this.repaginateFromIndex(0);
      this.onPagesChanged(this.pages);
    } finally {
      this.rebalancing = false;
      done({
        pages: this.pages.length,
      });
    }
  }

  rebalanceFromPage(page: HTMLElement, options: RebalanceOptions = {}): void {
    if (this.rebalancing) return;

    const index = this.pages.indexOf(page);
    if (index === -1) return;

    const includePreviousPage = options.includePreviousPage ?? true;
    const done = hweDebugStart("paginator.rebalanceFromPage", {
      compactPages: options.compactPages ?? true,
      includePreviousPage,
      index,
      pages: this.pages.length,
    });

    this.rebalancing = true;
    try {
      this.repaginateFromIndex(includePreviousPage ? Math.max(0, index - 1) : index, options);
      this.onPagesChanged(this.pages);
    } finally {
      this.rebalancing = false;
      done({
        pages: this.pages.length,
      });
    }
  }

  pushOverflowForwardFromPage(page: HTMLElement): boolean {
    if (this.rebalancing) return false;

    const index = this.pages.indexOf(page);
    if (index === -1) return false;

    const done = hweDebugStart("paginator.pushOverflowForwardFromPage", {
      index,
      pageOverflows: pageOverflows(page),
      pages: this.pages.length,
    });
    this.rebalancing = true;
    try {
      const moved = this.resolveOverflow(index);
      this.trimLeadingBlankBlocksFromContinuationPages();
      this.removeEmptyPages();
      this.onPagesChanged(this.pages);
      const pageStillOverflows = this.pages[index] ? pageOverflows(this.pages[index]) : false;
      done({
        moved,
        pageStillOverflows,
        pages: this.pages.length,
      });
      return moved && !pageStillOverflows;
    } finally {
      this.rebalancing = false;
    }
  }

  destroy(): void {
    this.pages = [];
  }

  private repaginateFromIndex(startIndex: number, options: RebalanceOptions = {}): void {
    const safeStart = Math.max(0, Math.min(startIndex, this.pages.length - 1));
    const shouldCompactPages = options.compactPages ?? true;
    const done = hweDebugStart("paginator.repaginateFromIndex", {
      compactPages: shouldCompactPages,
      pages: this.pages.length,
      startIndex: safeStart,
    });
    const nodes = this.keepTogetherController.groupFlowNodes(
      this.mergeFlowFragments(this.collectNodesFromIndex(safeStart))
    );

    this.trimPagesFromIndex(safeStart);

    let pageIndex = safeStart;
    for (const node of nodes) {
      pageIndex = this.appendNodeFlowing(node, pageIndex);
    }

    this.stabilizeOverflow();
    if (shouldCompactPages) this.compactPages();
    this.stabilizeOverflow();
    this.trimLeadingBlankBlocksFromContinuationPages();
    this.removeEmptyPages();
    done({
      nodes: nodes.length,
      pages: this.pages.length,
    });
  }

  private stabilizeOverflow(): void {
    let safety = 0;
    const startedAt = performance.now();

    while (safety++ < 200) {
      if (performance.now() - startedAt > Paginator.MAX_STABILIZE_MS) {
        hweDebugLog("paginator.stabilizeOverflow.timeBudgetExceeded", {
          elapsedMs: Math.round((performance.now() - startedAt) * 100) / 100,
          iterations: safety - 1,
          pages: this.pages.length,
        });
        break;
      }

      const overflowingIndex = this.pages.findIndex((page) => pageOverflows(page));
      if (overflowingIndex === -1) break;

      const moved = this.resolveOverflow(overflowingIndex);
      if (!moved) {
        hweDebugLog("paginator.stabilizeOverflow.unresolved", {
          overflowingIndex,
          pages: this.pages.length,
        });
        break;
      }
    }
  }

  private collectNodesFromIndex(startIndex: number): ChildNode[] {
    const nodes: ChildNode[] = [];

    this.pages.slice(startIndex).forEach((page) => {
      const inner = getInner(page);
      if (!inner) return;

      getMeaningfulChildren(inner).forEach((node) => {
        nodes.push(...this.flattenFlowNode(node));
      });
    });

    return nodes;
  }

  private mergeFlowFragments(nodes: ChildNode[]): ChildNode[] {
    const merged: ChildNode[] = [];

    nodes.forEach((node) => {
      const previous = merged[merged.length - 1];
      if (this.shouldMergeTableFragments(previous, node)) {
        this.mergeTableRows(previous as HTMLElement, node as HTMLElement);
        return;
      }

      if (this.shouldMergeTextFragments(previous, node)) {
        this.mergeTextBlockContents(previous as HTMLElement, node as HTMLElement);
        return;
      }

      merged.push(node);
    });

    return merged;
  }

  private shouldMergeTableFragments(
    previous: ChildNode | undefined,
    current: ChildNode
  ): boolean {
    if (!previous || !isTableElement(previous) || !isTableElement(current)) {
      return false;
    }

    const previousId = (previous as HTMLElement).getAttribute("data-hwe-table-flow-id");
    const currentId = (current as HTMLElement).getAttribute("data-hwe-table-flow-id");
    return !!previousId && previousId === currentId;
  }

  private mergeTableRows(targetTable: HTMLElement, sourceTable: HTMLElement): void {
    const targetBody = this.getOrCreateTableBody(targetTable);
    this.getTableBodyRows(sourceTable).forEach((row) => targetBody.appendChild(row));
    sourceTable.remove();
  }

  private shouldMergeTextFragments(
    previous: ChildNode | undefined,
    current: ChildNode
  ): boolean {
    if (
      !previous ||
      previous.nodeType !== Node.ELEMENT_NODE ||
      current.nodeType !== Node.ELEMENT_NODE
    ) {
      return false;
    }

    const previousElement = previous as HTMLElement;
    const currentElement = current as HTMLElement;
    if (previousElement.tagName !== currentElement.tagName) return false;

    const previousId = previousElement.getAttribute("data-hwe-text-flow-id");
    const currentId = currentElement.getAttribute("data-hwe-text-flow-id");
    return !!previousId && previousId === currentId;
  }

  private mergeTextBlockContents(targetBlock: HTMLElement, sourceBlock: HTMLElement): void {
    while (sourceBlock.firstChild) {
      targetBlock.appendChild(sourceBlock.firstChild);
    }

    targetBlock.removeAttribute("data-hwe-text-fragment");
    sourceBlock.remove();
  }

  private getOrCreateTableBody(table: HTMLElement): HTMLElement {
    const existingBody = table.querySelector("tbody");
    if (existingBody) return existingBody as HTMLElement;

    const body = document.createElement("tbody");
    table.appendChild(body);
    return body;
  }

  private getTableBodyRows(table: HTMLElement): HTMLElement[] {
    const thead = table.querySelector("thead");
    return (Array.from(table.querySelectorAll("tr")) as HTMLElement[]).filter(
      (row) => !thead?.contains(row)
    );
  }

  private flattenFlowNode(node: ChildNode): ChildNode[] {
    if (node.nodeType !== Node.ELEMENT_NODE) return [node];

    const element = node as HTMLElement;
    if (!this.isSplittableContainer(element) || shouldPreserveContainerShell(element)) {
      return [node];
    }

    const children = getMeaningfulChildren(element);
    if (children.length === 0) return [];

    return children.flatMap((child) => this.flattenFlowNode(child));
  }

  private trimPagesFromIndex(startIndex: number): void {
    const keepBefore = this.pages.slice(0, startIndex);
    let page = this.pages[startIndex];

    if (!page) {
      page = this.pageFactory();
    }

    const inner = getInner(page);
    if (inner) inner.innerHTML = "";

    this.pages.slice(startIndex + 1).forEach((extraPage) => extraPage.remove());
    this.pages = [...keepBefore, page];
  }

  private appendNodeFlowing(node: ChildNode, pageIndex: number): number {
    let currentIndex = pageIndex;
    let currentPage = this.pages[currentIndex] ?? this.getOrCreateNextPage(currentIndex - 1);
    let currentInner = getInner(currentPage);
    if (!currentInner) return currentIndex;

    if (node.nodeType === Node.ELEMENT_NODE) {
      this.constrainAtomicElement(node as HTMLElement, currentPage);
    }

    const pageHadContent = !!this.getFirstMeaningfulChild(currentInner);
    currentInner.appendChild(node);

    if (!pageOverflows(currentPage)) {
      return currentIndex;
    }

    if (
      !pageHadContent &&
      node.nodeType === Node.ELEMENT_NODE &&
      this.fitImagesToAvailableSpace(currentPage)
    ) {
      if (!pageOverflows(currentPage)) return currentIndex;
    }

    if (pageHadContent) {
      if (isTableElement(node)) {
        if (!this.elementFitsOnFreshPage(node as HTMLElement, currentPage)) {
          const nextPage = this.getOrCreateNextPage(currentIndex);
          const nextInner = getInner(nextPage);
          if (
            nextInner &&
            this.tablePaginator.splitTable(node as HTMLElement, nextInner, currentPage, {
              getContentLimitBottom: (inner) => getContentLimitBottom(inner),
              getInner: (page) => getInner(page),
            })
          ) {
            return currentIndex + 1;
          }
        }
      }

      const keepWithPrevious = this.detachPreviousKeepWithNext(currentInner, node);
      currentInner.removeChild(node);
      const nextPage = this.getOrCreateNextPage(currentIndex);
      const nextInner = getInner(nextPage);
      if (!nextInner) return currentIndex;

      if (node.nodeType === Node.ELEMENT_NODE) {
        this.constrainAtomicElement(node as HTMLElement, nextPage);
      }

      if (keepWithPrevious) nextInner.appendChild(keepWithPrevious);
      nextInner.appendChild(node);
      if (pageOverflows(nextPage)) {
        this.resolveOverflow(currentIndex + 1);
      }
      return currentIndex + 1;
    }

    this.resolveOverflow(currentIndex);

    if (pageOverflows(currentPage)) {
      return currentIndex;
    }

    const nextPage = this.pages[currentIndex + 1];
    if (!nextPage) return currentIndex;

    return currentIndex + 1;
  }

  private constrainAtomicElement(element: HTMLElement, page: HTMLElement): void {
    if (element.tagName === "IMG") {
      const contentHeight = getContentHeight(page);
      element.style.maxWidth = "100%";
      element.style.maxHeight = `${contentHeight}px`;
      element.style.height = "auto";
      element.style.objectFit = "contain";
    }
  }

  private elementFitsOnFreshPage(element: HTMLElement, page: HTMLElement): boolean {
    const inner = getInner(page);
    if (!inner) return false;

    const styles = getComputedStyle(inner);
    const paddingTop = parseFloat(styles.paddingTop) || 0;
    const availableHeight =
      getContentLimitBottom(inner) - (inner.getBoundingClientRect().top + paddingTop);

    return element.getBoundingClientRect().height <= availableHeight + 1;
  }

  private resolveOverflow(index: number): boolean {
    let currentIndex = index;
    let safety = 0;
    let movedAny = false;
    const startedAt = performance.now();
    const done = hweDebugStart("paginator.resolveOverflow", {
      index,
      pages: this.pages.length,
    });

    while (currentIndex < this.pages.length && safety++ < 500) {
      if (performance.now() - startedAt > Paginator.MAX_RESOLVE_MS) {
        hweDebugLog("paginator.resolveOverflow.timeBudgetExceeded", {
          currentIndex,
          elapsedMs: Math.round((performance.now() - startedAt) * 100) / 100,
          iterations: safety - 1,
          pages: this.pages.length,
        });
        break;
      }

      const page = this.pages[currentIndex];
      const inner = getInner(page);

      if (!inner || !pageOverflows(page)) {
        currentIndex++;
        continue;
      }

      const nextPage = this.getOrCreateNextPage(currentIndex);
      const nextInner = getInner(nextPage);
      if (!nextInner) break;

      const moved = this.moveOverflowPiece(inner, nextInner, page);
      movedAny = movedAny || moved;
      if (!moved) {
        hweDebugLog("paginator.resolveOverflow.noMove", {
          childCount: getMeaningfulChildren(inner).length,
          currentIndex,
          lastChild: this.describeNode(this.getLastMeaningfulChild(inner)),
          pageOverflows: pageOverflows(page),
        });
        currentIndex++;
      }
    }

    done({
      iterations: safety - 1,
      movedAny,
      pages: this.pages.length,
    });
    return movedAny;
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
      this.splitOrUnwrapSplittableContainer(
        lastChild as HTMLElement,
        targetInner,
        page
      )
    ) {
      return true;
    }

    if (
      lastChild.nodeType === Node.ELEMENT_NODE &&
      (lastChild as HTMLElement).tagName === "TABLE"
    ) {
      if (
        !this.elementFitsOnFreshPage(lastChild as HTMLElement, page) &&
        this.tablePaginator.splitTable(lastChild as HTMLElement, targetInner, page, {
          getContentLimitBottom: (inner) => getContentLimitBottom(inner),
          getInner: (currentPage) => getInner(currentPage),
        })
      ) {
        return true;
      }

      if (getMeaningfulChildren(sourceInner).length > 1) {
        const keepWithPrevious = this.detachPreviousKeepWithNext(sourceInner, lastChild);
        const targetReference = targetInner.firstChild;
        if (keepWithPrevious) targetInner.insertBefore(keepWithPrevious, targetReference);
        targetInner.insertBefore(lastChild, targetReference);
        return true;
      }

      return false;
    }

    if (
      lastChild.nodeType === Node.ELEMENT_NODE &&
      this.keepTogetherController.isKeepTogetherGroup(lastChild as HTMLElement)
    ) {
      if (this.fitImagesToAvailableSpace(page) && !pageOverflows(page)) {
        return true;
      }

      if (getMeaningfulChildren(sourceInner).length > 1) {
        targetInner.insertBefore(lastChild, targetInner.firstChild);
        return true;
      }

      this.keepTogetherController.unwrapGeneratedGroup(lastChild as HTMLElement);
      return true;
    }

    if (
      lastChild.nodeType === Node.ELEMENT_NODE &&
      this.shouldMoveTextBlockWhole(lastChild as HTMLElement, sourceInner, page)
    ) {
      const keepWithPrevious = this.detachPreviousKeepWithNext(sourceInner, lastChild);
      const targetReference = targetInner.firstChild;
      if (keepWithPrevious) targetInner.insertBefore(keepWithPrevious, targetReference);
      targetInner.insertBefore(lastChild, targetReference);
      hweDebugLog("paginator.moveOverflowPiece.textBlockWhole", {
        childCount: getMeaningfulChildren(sourceInner).length,
        tagName: (lastChild as HTMLElement).tagName,
      });
      return true;
    }

    if (
      lastChild.nodeType === Node.ELEMENT_NODE &&
      this.textBlockSplitter.splitTextBlock(lastChild as HTMLElement, targetInner, page)
    ) {
      return true;
    }

    if (getMeaningfulChildren(sourceInner).length <= 1) {
      return false;
    }

    targetInner.insertBefore(lastChild, targetInner.firstChild);
    return true;
  }

  private shouldMoveTextBlockWhole(
    element: HTMLElement,
    sourceInner: HTMLElement,
    page: HTMLElement
  ): boolean {
    return (
      isSplittableTextBlock(element) &&
      getMeaningfulChildren(sourceInner).length > 1 &&
      this.elementFitsOnFreshPage(element, page)
    );
  }

  private compactPages(): void {
    let index = 0;

    while (index < this.pages.length - 1) {
      const currentInner = getInner(this.pages[index]);
      const nextInner = getInner(this.pages[index + 1]);
      if (!currentInner || !nextInner) {
        index++;
        continue;
      }

      let movedAny = false;
      let safety = 0;

      while (safety++ < 100) {
        const candidates = this.takeCompactCandidates(nextInner);
        if (candidates.length === 0) break;
        if (this.shouldRespectUserBlankBarrier(currentInner, candidates)) {
          break;
        }

        const nextReference = candidates[candidates.length - 1].nextSibling;
        candidates.forEach((candidate) => currentInner.appendChild(candidate));
        this.fitImagesToAvailableSpace(this.pages[index]);

        if (pageOverflows(this.pages[index])) {
          const compactedTable = this.splitCompactedTableIntoNextPage(
            candidates,
            currentInner,
            nextInner,
            this.pages[index]
          );
          if (compactedTable) {
            movedAny = true;
            continue;
          }

          candidates.forEach((candidate) => nextInner.insertBefore(candidate, nextReference));
          break;
        }

        movedAny = true;
      }

      if (!movedAny) index++;
    }

    this.removeEmptyPages();
  }

  private endsWithUserBlankBlock(container: HTMLElement): boolean {
    const children = getMeaningfulChildren(container);
    const lastChild = children[children.length - 1];
    if (!lastChild || lastChild.nodeType !== Node.ELEMENT_NODE) return false;

    const element = lastChild as HTMLElement;
    return (
      element.getAttribute("data-hwe-user-blank") === "true" &&
      isEditableBlankBlock(element)
    );
  }

  private shouldRespectUserBlankBarrier(container: HTMLElement, candidates: ChildNode[]): boolean {
    return this.endsWithUserBlankBlock(container);
  }

  private takeCompactCandidates(nextInner: HTMLElement): ChildNode[] {
    const first = this.getFirstMeaningfulChild(nextInner);
    if (!first) return [];

    const candidates = [first];
    if (
      first.nodeType === Node.ELEMENT_NODE &&
      (first as HTMLElement).getAttribute("data-hwe-keep-with-next") === "true"
    ) {
      const next = this.getNextMeaningfulSibling(first);
      if (next && this.isKeepWithNextTarget(next)) {
        candidates.push(next);
      }
    }

    return candidates;
  }

  private splitCompactedTableIntoNextPage(
    candidates: ChildNode[],
    currentInner: HTMLElement,
    nextInner: HTMLElement,
    currentPage: HTMLElement
  ): boolean {
    if (candidates.length !== 1 || !isTableElement(candidates[0])) return false;

    const table = candidates[0] as HTMLElement;
    if (
      table.getAttribute("data-hwe-table-fragment") !== "true" &&
      this.elementFitsOnFreshPage(table, currentPage)
    ) {
      return false;
    }

    if (
      !this.tablePaginator.splitTable(table, nextInner, currentPage, {
        getContentLimitBottom: (inner) => getContentLimitBottom(inner),
        getInner: (page) => getInner(page),
      })
    ) {
      return false;
    }

    if (pageOverflows(currentPage)) {
      const firstFragment = nextInner.firstChild;
      if (firstFragment && isTableElement(firstFragment)) {
        this.mergeTableRows(table, firstFragment as HTMLElement);
      }
      nextInner.insertBefore(table, nextInner.firstChild);
      return false;
    }

    if (!this.getFirstMeaningfulChild(currentInner)) {
      table.remove();
      return false;
    }

    return true;
  }

  private getOrCreateNextPage(afterIndex: number): HTMLElement {
    const existingPage = this.pages[afterIndex + 1];
    if (existingPage) return existingPage;

    const previousPage = this.pages[afterIndex] ?? null;
    const newPage = this.pageFactory();
    this.pages.splice(afterIndex + 1, 0, newPage);
    if (this.onPageCreated) {
      this.onPageCreated(newPage, previousPage);
    } else {
      this.onPagesChanged(this.pages);
    }
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
    if (this.pages.length <= 1) {
      this.ensureSinglePageHasEditableBlank();
      return;
    }

    for (let index = this.pages.length - 1; index >= 0; index--) {
      const page = this.pages[index];
      const inner = getInner(page);
      if (inner && !this.getFirstMeaningfulChild(inner)) {
        this.removePage(index);
      }
    }

    if (this.pages.length === 0) {
      this.pages.push(this.pageFactory("<p><br></p>"));
    }

    this.ensureSinglePageHasEditableBlank();
  }

  private ensureSinglePageHasEditableBlank(): void {
    if (this.pages.length !== 1) return;

    const inner = getInner(this.pages[0]);
    if (!inner || inner.childNodes.length > 0) return;

    inner.innerHTML = "<p><br></p>";
  }

  private trimLeadingBlankBlocksFromContinuationPages(): void {
    this.pages.slice(1).forEach((page) => {
      const inner = getInner(page);
      if (!inner) return;

      let firstChild = inner.firstChild;
      while (firstChild && this.isLeadingPaginationBlank(firstChild)) {
        firstChild.remove();
        firstChild = inner.firstChild;
      }
    });
  }

  private isLeadingPaginationBlank(node: ChildNode): boolean {
    if (node.nodeType === Node.TEXT_NODE) {
      return (node.textContent ?? "").replace(/\u00a0/g, " ").trim() === "";
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return false;

    const element = node as HTMLElement;
    if (element.hasAttribute("data-hwe-user-blank")) return false;

    return isEditableBlankBlock(element);
  }

  private describeNode(node: ChildNode | null): unknown {
    if (!node) return null;
    if (node.nodeType === Node.TEXT_NODE) {
      return {
        text: (node.textContent ?? "").replace(/\s+/g, " ").trim().slice(0, 80),
        type: "text",
      };
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      return { type: `node-${node.nodeType}` };
    }

    const element = node as HTMLElement;
    return {
      blank: element.getAttribute("data-hwe-user-blank") === "true",
      className: element.getAttribute("class") ?? "",
      tagName: element.tagName,
      text: (element.textContent ?? "").replace(/\s+/g, " ").trim().slice(0, 80),
    };
  }

  private getLastMeaningfulChild(container: HTMLElement): ChildNode | null {
    const children = getMeaningfulChildren(container);
    if (children.length > 1) return children[children.length - 1];

    const onlyChild = children[0];
    if (
      onlyChild?.nodeType === Node.ELEMENT_NODE &&
      ((onlyChild as HTMLElement).tagName === "TABLE" ||
        this.keepTogetherController.isKeepTogetherGroup(onlyChild as HTMLElement) ||
        isTableFlowWrapper(onlyChild as HTMLElement) ||
        this.isSplittableContainer(onlyChild as HTMLElement) ||
        isSplittableTextBlock(onlyChild as HTMLElement))
    ) {
      return onlyChild;
    }

    return null;
  }

  private getFirstMeaningfulChild(container: HTMLElement): ChildNode | null {
    return getMeaningfulChildren(container)[0] ?? null;
  }

  private detachPreviousKeepWithNext(container: HTMLElement, node: ChildNode): ChildNode | null {
    const previous = this.getPreviousMeaningfulSibling(node);
    if (!previous || previous.nodeType !== Node.ELEMENT_NODE) return null;

    const element = previous as HTMLElement;
    if (element.getAttribute("data-hwe-keep-with-next") !== "true") return null;
    if (!this.isKeepWithNextTarget(node)) return null;

    container.removeChild(element);
    return element;
  }

  private getPreviousMeaningfulSibling(node: ChildNode): ChildNode | null {
    let previous = node.previousSibling;
    while (previous && isEmptyNode(previous)) {
      previous = previous.previousSibling;
    }

    return previous;
  }

  private getNextMeaningfulSibling(node: ChildNode): ChildNode | null {
    let next = node.nextSibling;
    while (next && isEmptyNode(next)) {
      next = next.nextSibling;
    }

    return next;
  }

  private isKeepWithNextTarget(node: ChildNode): boolean {
    if (node.nodeType !== Node.ELEMENT_NODE) return false;

    return ["IMG", "TABLE", "FIGURE", "VIDEO", "CANVAS", "SVG"].includes(
      (node as HTMLElement).tagName
    );
  }

  private fitImagesToAvailableSpace(page: HTMLElement): boolean {
    const inner = getInner(page);
    if (!inner) return false;

    const contentBottom = getContentLimitBottom(inner);
    let changed = false;

    inner.querySelectorAll<HTMLImageElement>("img").forEach((image) => {
      const imageTop = image.getBoundingClientRect().top;
      const availableHeight = Math.floor(contentBottom - imageTop);
      if (availableHeight < 120) return;

      const currentMaxHeight = parseFloat(image.style.maxHeight || "0");
      if (currentMaxHeight > 0 && currentMaxHeight <= availableHeight) return;

      image.style.maxWidth = "100%";
      image.style.maxHeight = `${availableHeight}px`;
      image.style.height = "auto";
      image.style.objectFit = "contain";
      changed = true;
    });

    return changed;
  }

  private splitOrUnwrapSplittableContainer(
    element: HTMLElement,
    targetInner: HTMLElement,
    page: HTMLElement
  ): boolean {
    if (!this.isSplittableContainer(element) || !element.parentNode) return false;
    if (shouldPreserveContainerShell(element)) {
      return this.splitContainerPreservingShell(element, targetInner, page);
    }

    const parent = element.parentNode;
    while (element.firstChild) {
      parent.insertBefore(element.firstChild, element);
    }
    parent.removeChild(element);
    return true;
  }

  private splitContainerPreservingShell(
    container: HTMLElement,
    targetInner: HTMLElement,
    page: HTMLElement
  ): boolean {
    const startedAt = performance.now();
    const overflowContainer = container.cloneNode(false) as HTMLElement;
    targetInner.insertBefore(overflowContainer, targetInner.firstChild);

    let movedAny = false;
    let safety = 0;

    while (pageOverflows(page) && safety++ < 300) {
      if (performance.now() - startedAt > Paginator.MAX_SPLIT_MS) {
        hweDebugLog("paginator.splitContainerPreservingShell.timeBudgetExceeded", {
          elapsedMs: Math.round((performance.now() - startedAt) * 100) / 100,
          movedAny,
          tagName: container.tagName,
        });
        break;
      }

      const child = this.getLastMeaningfulChild(container) ?? container.lastChild;
      if (!child) break;

      if (
        child.nodeType === Node.ELEMENT_NODE &&
        isTableElement(child) &&
        this.tablePaginator.splitTable(child as HTMLElement, overflowContainer, page, {
          getContentLimitBottom: (inner) => getContentLimitBottom(inner),
          getInner: (currentPage) => getInner(currentPage),
        })
      ) {
        movedAny = true;
        continue;
      }

      if (
        child.nodeType === Node.ELEMENT_NODE &&
        this.textBlockSplitter.splitTextBlock(child as HTMLElement, overflowContainer, page)
      ) {
        movedAny = true;
        continue;
      }

      overflowContainer.insertBefore(child, overflowContainer.firstChild);
      movedAny = true;

      if (isEmptyNode(container, false)) {
        container.remove();
        break;
      }
    }

    if (!movedAny || isEmptyNode(overflowContainer, false)) {
      overflowContainer.remove();
      return false;
    }

    return true;
  }

  private isSplittableContainer(element: HTMLElement): boolean {
    return isFlowSplittableContainer(
      element,
      (candidate) => this.keepTogetherController.isKeepTogetherGroup(candidate)
    );
  }
}
