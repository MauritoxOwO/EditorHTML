import { KeepTogetherController } from "./KeepTogetherController";
import { TablePaginator } from "./TablePaginator";
import { hweDebugLog, hweDebugStart } from "../debug/DebugLogger";

export type PageFactory = (html?: string) => HTMLElement;
export type OnPagesChanged = (pages: HTMLElement[]) => void;

export interface RebalanceOptions {
  includePreviousPage?: boolean;
  compactPages?: boolean;
  overflowOnly?: boolean;
}

export class Paginator {
  private static textFlowCounter = 0;

  private pages: HTMLElement[] = [];
  private readonly pageFactory: PageFactory;
  private readonly onPagesChanged: OnPagesChanged;
  private readonly keepTogetherController = new KeepTogetherController();
  private readonly tablePaginator = new TablePaginator();
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

  pushOverflowForwardFromPage(page: HTMLElement): void {
    if (this.rebalancing) return;

    const index = this.pages.indexOf(page);
    if (index === -1) return;

    const done = hweDebugStart("paginator.pushOverflowForwardFromPage", {
      index,
      pageOverflows: this.pageOverflows(page),
      pages: this.pages.length,
    });
    this.rebalancing = true;
    try {
      const moved = this.resolveOverflow(index);
      this.trimLeadingBlankBlocksFromContinuationPages();
      this.removeEmptyPages();
      this.onPagesChanged(this.pages);
      done({
        moved,
        pageStillOverflows: this.pages[index] ? this.pageOverflows(this.pages[index]) : false,
        pages: this.pages.length,
      });
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

    while (safety++ < 200) {
      const overflowingIndex = this.pages.findIndex((page) => this.pageOverflows(page));
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
      const inner = this.getInner(page);
      if (!inner) return;

      this.getMeaningfulChildren(inner).forEach((node) => {
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
    if (!previous || !this.isTableElement(previous) || !this.isTableElement(current)) {
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
    if (!this.isSplittableContainer(element) || this.shouldPreserveContainerShell(element)) {
      return [node];
    }

    const children = this.getMeaningfulChildren(element);
    if (children.length === 0) return [];

    return children.flatMap((child) => this.flattenFlowNode(child));
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

    if (
      !pageHadContent &&
      node.nodeType === Node.ELEMENT_NODE &&
      this.fitImagesToAvailableSpace(currentPage)
    ) {
      if (!this.pageOverflows(currentPage)) return currentIndex;
    }

    if (pageHadContent) {
      if (this.isTableElement(node)) {
        if (!this.elementFitsOnFreshPage(node as HTMLElement, currentPage)) {
          const nextPage = this.getOrCreateNextPage(currentIndex);
          const nextInner = this.getInner(nextPage);
          if (
            nextInner &&
            this.tablePaginator.splitTable(node as HTMLElement, nextInner, currentPage, {
              getContentLimitBottom: (inner) => this.getContentLimitBottom(inner),
              getInner: (page) => this.getInner(page),
            })
          ) {
            return currentIndex + 1;
          }
        }
      }

      const keepWithPrevious = this.detachPreviousKeepWithNext(currentInner, node);
      currentInner.removeChild(node);
      const nextPage = this.getOrCreateNextPage(currentIndex);
      const nextInner = this.getInner(nextPage);
      if (!nextInner) return currentIndex;

      if (node.nodeType === Node.ELEMENT_NODE) {
        this.constrainAtomicElement(node as HTMLElement, nextPage);
      }

      if (keepWithPrevious) nextInner.appendChild(keepWithPrevious);
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
      element.style.maxWidth = "100%";
      element.style.maxHeight = `${contentHeight}px`;
      element.style.height = "auto";
      element.style.objectFit = "contain";
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

  private elementFitsOnFreshPage(element: HTMLElement, page: HTMLElement): boolean {
    const inner = this.getInner(page);
    if (!inner) return false;

    const styles = getComputedStyle(inner);
    const paddingTop = parseFloat(styles.paddingTop) || 0;
    const availableHeight =
      this.getContentLimitBottom(inner) - (inner.getBoundingClientRect().top + paddingTop);

    return element.getBoundingClientRect().height <= availableHeight + 1;
  }

  private resolveOverflow(index: number): boolean {
    let currentIndex = index;
    let safety = 0;
    let movedAny = false;
    const done = hweDebugStart("paginator.resolveOverflow", {
      index,
      pages: this.pages.length,
    });

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
      movedAny = movedAny || moved;
      if (!moved) {
        hweDebugLog("paginator.resolveOverflow.noMove", {
          childCount: this.getMeaningfulChildren(inner).length,
          currentIndex,
          lastChild: this.describeNode(this.getLastMeaningfulChild(inner)),
          pageOverflows: this.pageOverflows(page),
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
          getContentLimitBottom: (inner) => this.getContentLimitBottom(inner),
          getInner: (currentPage) => this.getInner(currentPage),
        })
      ) {
        return true;
      }

      if (this.getMeaningfulChildren(sourceInner).length > 1) {
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
      if (this.fitImagesToAvailableSpace(page) && !this.pageOverflows(page)) {
        return true;
      }

      if (this.getMeaningfulChildren(sourceInner).length > 1) {
        targetInner.insertBefore(lastChild, targetInner.firstChild);
        return true;
      }

      this.keepTogetherController.unwrapGeneratedGroup(lastChild as HTMLElement);
      return true;
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

  private compactPages(): void {
    let index = 0;

    while (index < this.pages.length - 1) {
      const currentInner = this.getInner(this.pages[index]);
      const nextInner = this.getInner(this.pages[index + 1]);
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

        if (this.pageOverflows(this.pages[index])) {
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
    const children = this.getMeaningfulChildren(container);
    const lastChild = children[children.length - 1];
    if (!lastChild || lastChild.nodeType !== Node.ELEMENT_NODE) return false;

    const element = lastChild as HTMLElement;
    return (
      element.getAttribute("data-hwe-user-blank") === "true" &&
      this.isEditableBlankBlock(element)
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
    if (candidates.length !== 1 || !this.isTableElement(candidates[0])) return false;

    const table = candidates[0] as HTMLElement;
    if (
      table.getAttribute("data-hwe-table-fragment") !== "true" &&
      this.elementFitsOnFreshPage(table, currentPage)
    ) {
      return false;
    }

    if (
      !this.tablePaginator.splitTable(table, nextInner, currentPage, {
        getContentLimitBottom: (inner) => this.getContentLimitBottom(inner),
        getInner: (page) => this.getInner(page),
      })
    ) {
      return false;
    }

    if (this.pageOverflows(currentPage)) {
      const firstFragment = nextInner.firstChild;
      if (firstFragment && this.isTableElement(firstFragment)) {
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

  private trimLeadingBlankBlocksFromContinuationPages(): void {
    this.pages.slice(1).forEach((page) => {
      const inner = this.getInner(page);
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

    return this.isEditableBlankBlock(element);
  }

  private splitTextBlock(
    block: HTMLElement,
    targetInner: HTMLElement,
    page: HTMLElement
  ): boolean {
    if (!this.isSplittableTextBlock(block)) return false;

    const flowId = this.ensureTextFlowId(block);
    const overflowBlock = block.cloneNode(false) as HTMLElement;
    overflowBlock.setAttribute("data-hwe-text-flow-id", flowId);
    overflowBlock.setAttribute("data-hwe-text-fragment", "true");
    targetInner.insertBefore(overflowBlock, targetInner.firstChild);

    let movedAny = false;
    let safety = 0;

    while (this.pageOverflows(page) && safety++ < 300) {
      const moved = this.moveLastInlinePiece(block, overflowBlock);
      if (!moved) break;
      movedAny = true;

      if (this.isEmptyNode(block, false)) {
        block.remove();
        break;
      }
    }

    if (!movedAny || this.isEmptyNode(overflowBlock, false)) {
      overflowBlock.remove();
      return false;
    }

    return true;
  }

  private ensureTextFlowId(block: HTMLElement): string {
    const existingId = block.getAttribute("data-hwe-text-flow-id");
    if (existingId) return existingId;

    const id = `hwe-text-${Date.now().toString(36)}-${Paginator.textFlowCounter++}`;
    block.setAttribute("data-hwe-text-flow-id", id);
    block.setAttribute("data-hwe-text-fragment", "true");
    return id;
  }

  private moveLastInlinePiece(source: HTMLElement, target: HTMLElement): boolean {
    const child = source.lastChild;
    if (!child) return false;

    if (child.nodeType === Node.TEXT_NODE) {
      return this.moveLastWordFromTextNode(child as Text, target);
    }

    if (this.isEmptyNode(child)) {
      child.remove();
      return true;
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
    if (this.isEmptyNode(element, false)) element.remove();
    return true;
  }

  private moveLastWordFromTextNode(textNode: Text, target: HTMLElement): boolean {
    const text = textNode.textContent ?? "";

    if (!text) {
      return true;
    }

    if (/^\s+$/.test(text)) {
      if (target.childNodes.length > 0) return false;
      target.insertBefore(textNode, target.firstChild);
      return true;
    }

    const match = /(\s*\S+\s*)$/.exec(text);
    if (!match) return false;

    if (match.index <= 0) {
      if (target.childNodes.length > 0) return false;
      target.insertBefore(textNode, target.firstChild);
      return true;
    }

    const prefix = text.slice(0, match.index);
    const suffix = text.slice(match.index);

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

    const contentBottom = this.getContentBottom(inner);
    if (contentBottom === null) return false;

    return contentBottom > this.getContentLimitBottom(inner) + 1;
  }

  private getContentLimitBottom(inner: HTMLElement): number {
    const styles = getComputedStyle(inner);
    const paddingBottom = parseFloat(styles.paddingBottom) || 0;
    return inner.getBoundingClientRect().bottom - Math.min(paddingBottom, 12);
  }

  private getContentBottom(inner: HTMLElement): number | null {
    const children = this.getMeaningfulChildren(inner);
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
    const children = this.getMeaningfulChildren(container);
    if (children.length > 1) return children[children.length - 1];

    const onlyChild = children[0];
    if (
      onlyChild?.nodeType === Node.ELEMENT_NODE &&
      ((onlyChild as HTMLElement).tagName === "TABLE" ||
        this.keepTogetherController.isKeepTogetherGroup(onlyChild as HTMLElement) ||
        this.isTableFlowWrapper(onlyChild as HTMLElement) ||
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
    while (previous && this.isEmptyNode(previous)) {
      previous = previous.previousSibling;
    }

    return previous;
  }

  private getNextMeaningfulSibling(node: ChildNode): ChildNode | null {
    let next = node.nextSibling;
    while (next && this.isEmptyNode(next)) {
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

  private getMeaningfulChildren(container: HTMLElement): ChildNode[] {
    return Array.from(container.childNodes).filter((node) => !this.isEmptyNode(node));
  }

  private isTableElement(node: ChildNode): boolean {
    return node.nodeType === Node.ELEMENT_NODE && (node as HTMLElement).tagName === "TABLE";
  }

  private fitImagesToAvailableSpace(page: HTMLElement): boolean {
    const inner = this.getInner(page);
    if (!inner) return false;

    const contentBottom = this.getContentLimitBottom(inner);
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
    if (this.shouldPreserveContainerShell(element)) {
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
    const overflowContainer = container.cloneNode(false) as HTMLElement;
    targetInner.insertBefore(overflowContainer, targetInner.firstChild);

    let movedAny = false;
    let safety = 0;

    while (this.pageOverflows(page) && safety++ < 300) {
      const child = this.getLastMeaningfulChild(container) ?? container.lastChild;
      if (!child) break;

      if (
        child.nodeType === Node.ELEMENT_NODE &&
        this.isTableElement(child) &&
        this.tablePaginator.splitTable(child as HTMLElement, overflowContainer, page, {
          getContentLimitBottom: (inner) => this.getContentLimitBottom(inner),
          getInner: (currentPage) => this.getInner(currentPage),
        })
      ) {
        movedAny = true;
        continue;
      }

      if (
        child.nodeType === Node.ELEMENT_NODE &&
        this.splitTextBlock(child as HTMLElement, overflowContainer, page)
      ) {
        movedAny = true;
        continue;
      }

      overflowContainer.insertBefore(child, overflowContainer.firstChild);
      movedAny = true;

      if (this.isEmptyNode(container, false)) {
        container.remove();
        break;
      }
    }

    if (!movedAny || this.isEmptyNode(overflowContainer, false)) {
      overflowContainer.remove();
      return false;
    }

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
    if (element.classList.contains("hwe-page") || element.classList.contains("hwe-page-inner")) {
      return false;
    }
    if (this.keepTogetherController.isKeepTogetherGroup(element)) return false;
    if (this.isTableFlowWrapper(element)) return true;
    if (!splittableTags.has(element.tagName)) return false;

    return this.getMeaningfulChildren(element).length > 0;
  }

  private shouldPreserveContainerShell(element: HTMLElement): boolean {
    if (element.getAttribute("data-hwe-page-break") === "before") return false;

    return Array.from(element.attributes).some((attr) => {
      const name = attr.name.toLowerCase();
      const value = attr.value.trim();
      if (!value) return false;
      if (name === "contenteditable" || name === "spellcheck") return false;
      if (name.startsWith("data-hwe-")) return false;
      if (name === "style" && /^page-break-before\s*:\s*always\s*;?$/i.test(value)) {
        return false;
      }

      return true;
    });
  }

  private isTableFlowWrapper(element: HTMLElement): boolean {
    if (!element.querySelector("table")) return false;

    return ![
      "TABLE",
      "THEAD",
      "TBODY",
      "TFOOT",
      "TR",
      "TD",
      "TH",
      "COLGROUP",
      "COL",
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

  private isEmptyNode(node: ChildNode, preserveEditableBlankBlocks = false): boolean {
    if (node.nodeType === Node.COMMENT_NODE) return true;

    if (node.nodeType === Node.TEXT_NODE) {
      return (node.textContent ?? "").replace(/\u00a0/g, " ").trim() === "";
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return false;

    const element = node as HTMLElement;
    if (element.hasAttribute("data-hwe-user-blank")) return false;
    if (element.hasAttribute("data-hwe-caret")) return false;
    if (element.tagName === "BR") return true;
    if (["META", "LINK", "STYLE", "SCRIPT", "XML"].includes(element.tagName)) return true;
    if (element.querySelector("img, table, tr, td, th, video, canvas, svg")) return false;
    if (preserveEditableBlankBlocks && this.isEditableBlankBlock(element)) return false;

    return (
      ["P", "DIV", "SECTION", "ARTICLE", "SPAN"].includes(element.tagName) &&
      (element.textContent ?? "").replace(/\u00a0/g, " ").trim() === "" &&
      Array.from(element.childNodes).every((child) => this.isEmptyNode(child, false))
    );
  }

  private isEditableBlankBlock(element: HTMLElement): boolean {
    const blankBlockTags = new Set([
      "P",
      "DIV",
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

    if (!blankBlockTags.has(element.tagName)) return false;
    if ((element.textContent ?? "").replace(/\u00a0/g, " ").trim() !== "") return false;

    return Array.from(element.childNodes).every((child) => {
      if (child.nodeType === Node.TEXT_NODE) {
        return (child.textContent ?? "").replace(/\u00a0/g, " ").trim() === "";
      }

      return child.nodeType === Node.ELEMENT_NODE && (child as HTMLElement).tagName === "BR";
    });
  }
}
