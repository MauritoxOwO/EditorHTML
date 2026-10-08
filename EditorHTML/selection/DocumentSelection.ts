export interface SelectedTextTarget {
  node: Text;
  start: number;
  end: number;
  page: HTMLElement;
  flowId: string;
}

export interface DocumentSelectionSnapshot {
  range: Range;
  anchorNode: Node;
  anchorOffset: number;
  focusNode: Node;
  focusOffset: number;
  collapsed: boolean;
  backward: boolean;
  pages: HTMLElement[];
}

const EXCLUDED = [
  "[contenteditable='false']",
  "[hidden]",
  "[data-hwe-api-header]",
  "[data-hwe-dynamic-header]",
  "[data-hwe-runtime-page-header]",
  "[data-hwe-page-divider]",
  "[data-hwe-caret]",
  "style",
  "script",
].join(",");

const FLOW_BLOCKS = "p, div, li, h1, h2, h3, h4, h5, h6, blockquote, pre";

export class DocumentSelection {
  constructor(private readonly rootProvider: () => HTMLElement) {}

  readSelection(): DocumentSelectionSnapshot | null {
    const root = this.rootProvider();
    const selection = window.getSelection();
    if (!selection?.anchorNode || !selection.focusNode || selection.rangeCount === 0) return null;

    const sourceRange = selection.getRangeAt(0).cloneRange();
    if (!root.contains(selection.anchorNode) || !root.contains(selection.focusNode)) return null;
    const range = this.normalizeRange(sourceRange);
    if (!range) return null;
    const backward = !selection.isCollapsed && this.isAfter(
      selection.anchorNode, selection.anchorOffset, selection.focusNode, selection.focusOffset
    );

    return {
      range,
      anchorNode: selection.anchorNode,
      anchorOffset: selection.anchorOffset,
      focusNode: selection.focusNode,
      focusOffset: selection.focusOffset,
      collapsed: selection.isCollapsed,
      backward,
      pages: this.getAffectedPages(range),
    };
  }

  /** Expand host-level selections (for example Ctrl+A) to the editable page contents. */
  normalizeRange(source: Range): Range | null {
    const root = this.rootProvider();
    if (!root.contains(source.startContainer) || !root.contains(source.endContainer)) return null;
    const pages = this.getAffectedPages(source);
    const firstPage = pages[0] ?? root.querySelector<HTMLElement>(".hwe-page");
    const lastPage = pages[pages.length - 1] ?? firstPage;
    if (!firstPage || !lastPage) return null;
    const startInner = this.getInnerForPoint(source.startContainer);
    const endInner = this.getInnerForPoint(source.endContainer);
    const startExcluded = this.getExcludedAncestor(source.startContainer);
    const endExcluded = this.getExcludedAncestor(source.endContainer);
    if (!source.collapsed && startExcluded && startExcluded === endExcluded) return null;
    if (source.collapsed) {
      if (startInner) return source.cloneRange();
      const page = root.contains(source.startContainer)
        ? this.pageAtPoint(source.startContainer, source.startOffset)
        : firstPage;
      const inner = page?.querySelector<HTMLElement>(".hwe-page-inner");
      if (!inner) return null;
      const rawInner = this.getContainingInner(source.startContainer);
      const excluded = startExcluded;
      const point = excluded && rawInner === inner
        ? this.pointAfter(inner, excluded)
        : this.firstEditablePoint(inner);
      const result = document.createRange();
      result.setStart(point.node, point.offset);
      result.collapse(true);
      return result;
    }

    const first = startInner?.closest<HTMLElement>(".hwe-page")?.querySelector<HTMLElement>(".hwe-page-inner") ??
      (root.contains(source.startContainer) ? this.getContainingInner(source.startContainer) : null) ??
      firstPage.querySelector<HTMLElement>(".hwe-page-inner");
    const last = endInner?.closest<HTMLElement>(".hwe-page")?.querySelector<HTMLElement>(".hwe-page-inner") ??
      (root.contains(source.endContainer) ? this.getContainingInner(source.endContainer) : null) ??
      lastPage.querySelector<HTMLElement>(".hwe-page-inner");
    if (!first || !last || !root.contains(first) || !root.contains(last)) return null;

    const range = source.cloneRange();
    if (!startInner) {
      const excluded = startExcluded;
      const point = excluded ? this.pointAfter(first, excluded) : this.firstEditablePoint(first);
      range.setStart(point.node, point.offset);
    }
    if (!endInner) {
      const excluded = endExcluded;
      const point = excluded ? this.pointBefore(last, excluded) : this.lastEditablePoint(last);
      range.setEnd(point.node, point.offset);
    }
    if (range.collapsed && !source.collapsed) return null;
    return range;
  }

  getAffectedPages(range: Range): HTMLElement[] {
    const root = this.rootProvider();
    return Array.from(root.querySelectorAll<HTMLElement>(".hwe-page")).filter((page) => {
      try {
        return range.intersectsNode(page);
      } catch {
        return false;
      }
    });
  }

  getSelectedTextTargets(range: Range): SelectedTextTarget[] {
    const root = this.rootProvider();
    const targets: SelectedTextTarget[] = [];
    const pages = Array.from(root.querySelectorAll<HTMLElement>(".hwe-page"));

    for (const page of pages) {
      const inner = page.querySelector<HTMLElement>(".hwe-page-inner");
      if (!inner) continue;
      const walker = document.createTreeWalker(inner, NodeFilter.SHOW_TEXT, {
        acceptNode: (node) => {
          const text = node as Text;
          if (!text.data || this.isExcluded(text)) return NodeFilter.FILTER_REJECT;
          return NodeFilter.FILTER_ACCEPT;
        },
      });

      let node: Text | null;
      while ((node = walker.nextNode() as Text | null)) {
        const offsets = this.getSelectedOffsets(range, node);
        if (!offsets || offsets.start >= offsets.end) continue;
        const flowBlock = this.closestFlowBlock(node);
        targets.push({
          node,
          start: offsets.start,
          end: offsets.end,
          page,
          flowId: flowBlock?.getAttribute("data-hwe-text-flow-id") ?? "",
        });
      }
    }

    return targets;
  }

  getSelectedParagraphs(range: Range): HTMLElement[] {
    const root = this.rootProvider();
    const blocks = Array.from(root.querySelectorAll<HTMLElement>(FLOW_BLOCKS));
    return blocks.filter((block) => {
      if (!block.closest(".hwe-page-inner") || this.isExcluded(block)) return false;
      if (block.matches(".hwe-page, .hwe-page-inner, .hwe-workspace, [data-hwe-generated-wrapper], [data-hwe-keep-together]")) return false;
      if (!this.hasContent(block)) return false;

      const blockRange = document.createRange();
      blockRange.selectNodeContents(block);
      const startsBeforeBlockEnds = range.compareBoundaryPoints(Range.START_TO_END, blockRange) > 0;
      const endsAfterBlockStarts = range.compareBoundaryPoints(Range.END_TO_START, blockRange) < 0;
      return startsBeforeBlockEnds && endsAfterBlockStarts;
    });
  }

  isSafeSinglePageRange(range: Range): boolean {
    const startInner = this.getInnerForPoint(range.startContainer);
    const endInner = this.getInnerForPoint(range.endContainer);
    if (!startInner || startInner !== endInner) return false;
    if (this.rangeIntersectsExcludedContent(range, startInner) || this.rangeIntersectsUnsafeStructure(range)) return false;

    const affectedPages = this.getAffectedPages(range);
    return affectedPages.length === 1 && affectedPages[0].contains(startInner);
  }

  isEditableRange(range: Range): boolean {
    const normalized = this.normalizeRange(range);
    if (!normalized) return false;
    const pages = this.getAffectedPages(normalized);
    if (!pages.length) return false;
    return pages.every((page) => !!page.querySelector<HTMLElement>(".hwe-page-inner"));
  }

  isAtPageBoundary(range: Range, direction: "backward" | "forward"): boolean {
    const inner = this.getInnerForPoint(range.startContainer);
    if (!inner || range.startContainer !== range.endContainer) return false;
    const page = inner.closest<HTMLElement>(".hwe-page");
    const pages = Array.from(this.rootProvider().querySelectorAll<HTMLElement>(".hwe-page"));
    if (!page) return false;

    const edgeRange = document.createRange();
    edgeRange.selectNodeContents(inner);
    if (direction === "backward") edgeRange.setEnd(range.startContainer, range.startOffset);
    else edgeRange.setStart(range.startContainer, range.startOffset);

    if (edgeRange.toString().length !== 0 || this.rangeIntersectsEdgeObject(edgeRange, inner)) return false;
    const pageIndex = pages.indexOf(page);
    return direction === "backward" ? pageIndex > 0 : pageIndex >= 0 && pageIndex < pages.length - 1;
  }

  getClipboardContent(range: Range): { html: string; text: string } | null {
    const normalized = this.normalizeRange(range);
    if (!normalized || !this.isEditableRange(normalized)) return null;
    const clipboardRoot = document.createElement("div");
    const copiedFlows = new Map<string, HTMLElement>();
    const seenHeaders = new Set<string>();
    for (const page of this.getAffectedPages(normalized)) {
      const inner = page.querySelector<HTMLElement>(".hwe-page-inner");
      if (!inner) continue;
      const local = this.clipRangeToElement(normalized, inner);
      if (!local || local.collapsed) continue;
      const wrapper = document.createElement("div");
      wrapper.appendChild(local.cloneContents());
      wrapper.querySelectorAll<HTMLElement>(EXCLUDED).forEach((element) => element.remove());
      wrapper.querySelectorAll<HTMLElement>("thead").forEach((header) => {
        const table = header.closest<HTMLTableElement>("table[data-hwe-repeat-header='true']");
        const id = table?.getAttribute("data-hwe-table-flow-id");
        if (id && seenHeaders.has(id)) header.remove();
        else if (id) seenHeaders.add(id);
      });
      wrapper.querySelectorAll<HTMLElement>("[data-hwe-text-flow-id]").forEach((block) => {
        const id = block.getAttribute("data-hwe-text-flow-id");
        if (!id || !block.matches(FLOW_BLOCKS)) return;
        const previous = copiedFlows.get(id);
        if (!previous) {
          copiedFlows.set(id, block);
          return;
        }
        while (block.firstChild) previous.appendChild(block.firstChild);
        block.remove();
      });
      clipboardRoot.append(...Array.from(wrapper.childNodes));
    }

    clipboardRoot.querySelectorAll<HTMLElement>(
      ".hwe-text-flow-column, [data-hwe-generated-wrapper='true'], [data-hwe-keep-together='true']"
    ).forEach((wrapper) => wrapper.replaceWith(...Array.from(wrapper.childNodes)));
    clipboardRoot.querySelectorAll<HTMLElement>("*").forEach((element) => {
      Array.from(element.attributes).filter((attribute) => attribute.name.startsWith("data-hwe-"))
        .forEach((attribute) => element.removeAttribute(attribute.name));
    });
    return { html: clipboardRoot.innerHTML, text: this.getClipboardText(clipboardRoot) };
  }

  private getClipboardText(root: HTMLElement): string {
    const serialize = (node: Node): string => {
      if (node.nodeType === Node.TEXT_NODE) {
        return node.textContent ?? "";
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return "";
      const element = node as HTMLElement;
      if (element.tagName === "BR") return "\n";
      const children = Array.from(element.childNodes).map(serialize);
      if (element.tagName === "TR") {
        return Array.from(element.children)
          .map((cell) => serialize(cell).replace(/\n+$/, ""))
          .join("\t") + "\n";
      }
      if (element.matches("td, th")) return children.join("");
      const isBlock = element.matches("p, div, li, h1, h2, h3, h4, h5, h6, blockquote, pre");
      const content = children.join("");
      return isBlock ? `${content}\n` : content;
    };
    return Array.from(root.childNodes).map(serialize).join("").replace(/\n+$/, "");
  }

  /**
   * Apply a text replacement without ever deleting a range spanning page wrappers.
   * Each page gets its own clipped range, applied from the last page toward the first.
   */
  replaceSelection(range: Range, replacement: string): HTMLElement | null {
    const normalized = this.normalizeRange(range);
    if (!normalized || !this.isEditableRange(normalized)) return null;
    const pages = this.getAffectedPages(normalized);
    const edits: Array<{ page: HTMLElement; range: Range }> = [];

    for (const page of pages) {
      const inner = page.querySelector<HTMLElement>(".hwe-page-inner");
      if (!inner) continue;
      const local = document.createRange();
      const startsInside = inner.contains(normalized.startContainer);
      const endsInside = inner.contains(normalized.endContainer);
      if (startsInside) local.setStart(normalized.startContainer, normalized.startOffset);
      else local.setStart(inner, 0);
      if (endsInside) local.setEnd(normalized.endContainer, normalized.endOffset);
      else local.setEnd(inner, inner.childNodes.length);
      if (!local.collapsed) edits.push({ page, range: local });
    }

    if (edits.length === 0) {
      const pointPage = this.getEditableInnerForNode(normalized.startContainer)?.closest<HTMLElement>(".hwe-page");
      if (!pointPage) return null;
      const local = normalized.cloneRange();
      if (replacement) this.insertReplacement(local, replacement);
      local.collapse(false);
      const selection = window.getSelection();
      selection?.removeAllRanges();
      selection?.addRange(local);
      return pointPage;
    }

    // Save the first clipped range; Range boundaries track the DOM edits and provide
    // the insertion point after the selected content has been removed.
    const first = edits[0];
    const caret = first.range.cloneRange();
    caret.collapse(true);
    const startBlock = this.closestFlowBlock(normalized.startContainer);
    const endBlock = this.closestFlowBlock(normalized.endContainer);
    const crossesTable = this.rangeIntersectsTable(normalized) || this.getAffectedPages(normalized).some((page) => {
      const inner = page.querySelector<HTMLElement>(".hwe-page-inner");
      return !!inner && this.rangeIntersectsExcludedContent(normalized, inner);
    });
    for (const edit of [...edits].reverse()) {
      const inner = edit.page.querySelector<HTMLElement>(".hwe-page-inner")!;
      const plan = this.planLocalDeletion(edit.range, inner);
      for (const table of plan.wholeTables) table.remove();
      for (const segment of plan.segments.reverse()) segment.deleteContents();
    }

    this.joinCompatibleBoundaryBlocks(startBlock, endBlock, crossesTable);

    const startInner = first.page.querySelector<HTMLElement>(".hwe-page-inner");
    if (!startInner || !startInner.contains(caret.startContainer)) return null;
    if (replacement) {
      this.insertReplacement(caret, replacement);
    }
    caret.collapse(true);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(caret);
    this.ensureEditableParagraph(startInner);
    return first.page;
  }

  getEditableInnerForNode(node: Node | null): HTMLElement | null {
    return this.getInnerForPoint(node);
  }

  private getSelectedOffsets(range: Range, node: Text): { start: number; end: number } | null {
    try {
      if (!range.intersectsNode(node)) return null;
      const nodeRange = document.createRange();
      nodeRange.selectNodeContents(node);
      if (range.compareBoundaryPoints(Range.END_TO_START, nodeRange) >= 0 ||
          range.compareBoundaryPoints(Range.START_TO_END, nodeRange) <= 0) return null;
      const start = range.startContainer === node ? range.startOffset : 0;
      const end = range.endContainer === node ? range.endOffset : node.length;
      return { start: Math.max(0, start), end: Math.min(node.length, end) };
    } catch {
      return null;
    }
  }

  private getInnerForPoint(node: Node | null): HTMLElement | null {
    if (!node) return null;
    const root = this.rootProvider();
    if (!root.contains(node)) return null;
    const element = node.nodeType === Node.ELEMENT_NODE ? node as HTMLElement : node.parentElement;
    if (!element || this.isExcluded(element)) return null;
    const inner = element.closest<HTMLElement>(".hwe-page-inner");
    return inner && root.contains(inner) && inner.contains(node) ? inner : null;
  }

  private getContainingInner(node: Node | null): HTMLElement | null {
    const element = node?.nodeType === Node.ELEMENT_NODE ? node as HTMLElement : node?.parentElement;
    return element?.closest<HTMLElement>(".hwe-page-inner") ?? null;
  }

  private getExcludedAncestor(node: Node | null): HTMLElement | null {
    const element = node?.nodeType === Node.ELEMENT_NODE ? node as HTMLElement : node?.parentElement;
    return element?.closest<HTMLElement>(EXCLUDED) ?? null;
  }

  private closestFlowBlock(node: Node): HTMLElement | null {
    const element = node.nodeType === Node.ELEMENT_NODE ? node as HTMLElement : node.parentElement;
    const block = element?.closest<HTMLElement>(FLOW_BLOCKS) ?? null;
    return block?.closest(".hwe-page-inner") ? block : null;
  }

  private hasContent(block: HTMLElement): boolean {
    if ((block.textContent ?? "").length > 0) return true;
    return Boolean(block.querySelector("img, table, br, hr, video, canvas, svg"));
  }

  private isExcluded(node: Node): boolean {
    const element = node.nodeType === Node.ELEMENT_NODE ? node as HTMLElement : node.parentElement;
    return Boolean(element?.closest(EXCLUDED));
  }

  private rangeIntersectsExcludedContent(range: Range, inner: HTMLElement): boolean {
    return Array.from(inner.querySelectorAll<HTMLElement>(EXCLUDED)).some((element) => {
      try {
        return range.intersectsNode(element);
      } catch {
        return false;
      }
    });
  }

  private subtractExcludedRanges(range: Range, inner: HTMLElement): Range[] {
    const exclusions = Array.from(inner.querySelectorAll<HTMLElement>(EXCLUDED)).filter((element) => {
      try { return range.intersectsNode(element); } catch { return false; }
    });
    if (!exclusions.length) return [range.cloneRange()];
    const result: Range[] = [];
    let cursor = range.cloneRange();
    cursor.collapse(true);
    for (const element of exclusions) {
      const before = document.createRange();
      before.setStartBefore(element);
      before.collapse(true);
      const after = document.createRange();
      after.setStartAfter(element);
      after.collapse(true);
      if (cursor.compareBoundaryPoints(Range.START_TO_START, before) < 0) {
        const segment = document.createRange();
        segment.setStart(cursor.startContainer, cursor.startOffset);
        segment.setEnd(before.startContainer, before.startOffset);
        if (!segment.collapsed) result.push(segment);
      }
      if (cursor.compareBoundaryPoints(Range.START_TO_START, after) < 0) {
        cursor = after;
      }
    }
    const end = document.createRange();
    end.setStart(range.endContainer, range.endOffset);
    end.collapse(true);
    if (cursor.compareBoundaryPoints(Range.START_TO_START, end) < 0) {
      const segment = document.createRange();
      segment.setStart(cursor.startContainer, cursor.startOffset);
      segment.setEnd(end.startContainer, end.startOffset);
      if (!segment.collapsed) result.push(segment);
    }
    return result;
  }

  private planLocalDeletion(range: Range, inner: HTMLElement): { segments: Range[]; wholeTables: HTMLElement[] } {
    const segments: Range[] = [];
    const wholeTables: HTMLElement[] = [];
    const ancestorTable = inner.closest("table");
    const tables = Array.from(inner.querySelectorAll<HTMLElement>("table")).filter((table) => {
      const parentTable = table.parentElement?.closest("table");
      if (parentTable !== ancestorTable) return false;
      try { return range.intersectsNode(table); } catch { return false; }
    });
    if (!tables.length) return { segments: this.subtractExcludedRanges(range, inner), wholeTables };
    let cursor = range.cloneRange();
    cursor.collapse(true);
    for (const table of tables) {
      const before = document.createRange();
      before.setStartBefore(table);
      before.collapse(true);
      if (cursor.compareBoundaryPoints(Range.START_TO_START, before) < 0) {
        const outside = document.createRange();
        outside.setStart(cursor.startContainer, cursor.startOffset);
        outside.setEnd(before.startContainer, before.startOffset);
        segments.push(...this.subtractExcludedRanges(outside, inner));
      }
      const protectedInside = Array.from(table.querySelectorAll<HTMLElement>(EXCLUDED)).length > 0;
      if (this.rangeFullyContains(range, table) && !protectedInside) {
        wholeTables.push(table);
        const after = document.createRange();
        after.setStartAfter(table);
        after.collapse(true);
        cursor = after;
        continue;
      }
      for (const cell of Array.from(table.querySelectorAll<HTMLElement>("td, th")).filter((item) => item.closest("table") === table)) {
        let intersects = false;
        try { intersects = range.intersectsNode(cell); } catch { /* detached */ }
        if (!intersects) continue;
        const cellRange = this.clipRangeToElement(range, cell);
        if (cellRange && !cellRange.collapsed) {
          const nested = this.planLocalDeletion(cellRange, cell);
          segments.push(...nested.segments);
          wholeTables.push(...nested.wholeTables);
        }
      }
      const after = document.createRange();
      after.setStartAfter(table);
      after.collapse(true);
      cursor = after;
    }
    const end = document.createRange();
    end.setStart(range.endContainer, range.endOffset);
    end.collapse(true);
    if (cursor.compareBoundaryPoints(Range.START_TO_START, end) < 0) {
      const outside = document.createRange();
      outside.setStart(cursor.startContainer, cursor.startOffset);
      outside.setEnd(end.startContainer, end.startOffset);
      segments.push(...this.subtractExcludedRanges(outside, inner));
    }
    return { segments, wholeTables };
  }

  private clipRangeToElement(range: Range, element: HTMLElement): Range | null {
    const clipped = range.cloneRange();
    const startInside = element.contains(range.startContainer);
    const endInside = element.contains(range.endContainer);
    if (!startInside) clipped.setStart(element, 0);
    if (!endInside) clipped.setEnd(element, element.childNodes.length);
    return clipped;
  }

  private rangeFullyContains(range: Range, element: HTMLElement): boolean {
    const before = document.createRange();
    before.selectNode(element);
    return range.compareBoundaryPoints(Range.START_TO_START, before) <= 0 &&
      range.compareBoundaryPoints(Range.END_TO_END, before) >= 0;
  }

  private joinCompatibleBoundaryBlocks(
    startBlock: HTMLElement | null,
    endBlock: HTMLElement | null,
    crossesTable: boolean
  ): void {
    if (!startBlock || !endBlock || startBlock === endBlock || crossesTable ||
        !startBlock.isConnected || !endBlock.isConnected) return;
    const allowed = "p, h1, h2, h3, h4, h5, h6, blockquote, pre";
    if (!startBlock.matches(allowed) || !endBlock.matches(allowed) ||
        startBlock.tagName !== endBlock.tagName || startBlock.closest("td, th") || endBlock.closest("td, th") ||
        startBlock.getAttribute("class") !== endBlock.getAttribute("class") ||
        startBlock.getAttribute("style") !== endBlock.getAttribute("style")) return;
    if (!(startBlock.compareDocumentPosition(endBlock) & Node.DOCUMENT_POSITION_FOLLOWING)) return;
    while (endBlock.firstChild) startBlock.appendChild(endBlock.firstChild);
    endBlock.remove();
  }

  private pageAtPoint(node: Node, offset: number): HTMLElement | null {
    const root = this.rootProvider();
    const element = node.nodeType === Node.ELEMENT_NODE ? node as HTMLElement : node.parentElement;
    const directPage = element?.closest<HTMLElement>(".hwe-page");
    if (directPage) return directPage;
    const pages = Array.from(root.querySelectorAll<HTMLElement>(".hwe-page"));
    if (node === root) return pages[Math.min(offset, pages.length - 1)] ?? pages[pages.length - 1] ?? null;
    const point = document.createRange();
    point.setStart(node, offset);
    point.collapse(true);
    return pages.find((page) => {
      const pageRange = document.createRange();
      pageRange.selectNode(page);
      return point.compareBoundaryPoints(Range.START_TO_START, pageRange) <= 0;
    }) ?? pages[pages.length - 1] ?? null;
  }

  private firstEditablePoint(inner: HTMLElement): { node: Node; offset: number } {
    const walker = document.createTreeWalker(inner, NodeFilter.SHOW_TEXT, {
      acceptNode: (node) => this.isExcluded(node) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT,
    });
    const text = walker.nextNode();
    return text ? { node: text, offset: 0 } : { node: inner, offset: 0 };
  }

  private pointAfter(inner: HTMLElement, excluded: HTMLElement): { node: Node; offset: number } {
    const walker = document.createTreeWalker(inner, NodeFilter.SHOW_TEXT, {
      acceptNode: (node) => this.isExcluded(node) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT,
    });
    let next: Node | null;
    while ((next = walker.nextNode())) {
      if (excluded.compareDocumentPosition(next) & Node.DOCUMENT_POSITION_FOLLOWING) {
        return { node: next, offset: 0 };
      }
    }
    return { node: inner, offset: inner.childNodes.length };
  }

  private pointBefore(inner: HTMLElement, excluded: HTMLElement): { node: Node; offset: number } {
    const walker = document.createTreeWalker(inner, NodeFilter.SHOW_TEXT, {
      acceptNode: (node) => this.isExcluded(node) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT,
    });
    let previous: Node | null = null;
    let next: Node | null;
    while ((next = walker.nextNode())) {
      if (excluded.compareDocumentPosition(next) & Node.DOCUMENT_POSITION_PRECEDING) break;
      previous = next;
    }
    return previous ? { node: previous, offset: previous.textContent?.length ?? 0 } : { node: inner, offset: 0 };
  }

  private lastEditablePoint(inner: HTMLElement): { node: Node; offset: number } {
    const walker = document.createTreeWalker(inner, NodeFilter.SHOW_TEXT, {
      acceptNode: (node) => this.isExcluded(node) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT,
    });
    let text: Node | null = null;
    let next: Node | null;
    while ((next = walker.nextNode())) text = next;
    return text ? { node: text, offset: text.textContent?.length ?? 0 } :
      { node: inner, offset: inner.childNodes.length };
  }

  private isAfter(anchor: Node, anchorOffset: number, focus: Node, focusOffset: number): boolean {
    const a = document.createRange();
    a.setStart(anchor, anchorOffset);
    a.collapse(true);
    const f = document.createRange();
    f.setStart(focus, focusOffset);
    f.collapse(true);
    return a.compareBoundaryPoints(Range.START_TO_START, f) > 0;
  }

  private ensureEditableParagraph(inner: HTMLElement): void {
    const hasEditableContent = Array.from(inner.childNodes).some((node) => {
      if (node.nodeType === Node.TEXT_NODE) return !!node.textContent?.trim();
      if (node.nodeType !== Node.ELEMENT_NODE) return false;
      const element = node as HTMLElement;
      return !this.isExcluded(element) && !element.matches("[data-hwe-manual-page-break='true']");
    });
    if (!hasEditableContent) {
      const paragraph = document.createElement("p");
      paragraph.appendChild(document.createElement("br"));
      inner.appendChild(paragraph);
    }
  }

  private insertReplacement(caret: Range, replacement: string): void {
    const lines = replacement.replace(/\r\n?/g, "\n").split("\n");
    const fragment = document.createDocumentFragment();
    const inserted: Node[] = [];
    lines.forEach((line, index) => {
      if (line) {
        const text = document.createTextNode(line);
        inserted.push(text);
        fragment.appendChild(text);
      }
      if (index < lines.length - 1) {
        const lineBreak = document.createElement("br");
        inserted.push(lineBreak);
        fragment.appendChild(lineBreak);
      }
    });
    caret.insertNode(fragment);
    const last = inserted.at(-1);
    if (last?.parentNode) {
      if (last.nodeType === Node.TEXT_NODE) caret.setStart(last, last.textContent?.length ?? 0);
      else caret.setStartAfter(last);
    }
  }

  private rangeIntersectsUnsafeStructure(range: Range): boolean {
    const root = this.rootProvider();
    return Array.from(root.querySelectorAll<HTMLElement>(
      ".hwe-page-inner img, .hwe-page-inner table, .hwe-page-inner br, .hwe-page-inner hr, " +
      ".hwe-page-inner [data-hwe-manual-page-break='true']"
    )).some((element) => {
      try {
        return range.intersectsNode(element);
      } catch {
        return false;
      }
    });
  }

  private rangeIntersectsTable(range: Range): boolean {
    return Array.from(this.rootProvider().querySelectorAll<HTMLElement>(".hwe-page-inner table"))
      .some((table) => {
        try { return range.intersectsNode(table); } catch { return false; }
      });
  }

  private rangeIntersectsEdgeObject(range: Range, inner: HTMLElement): boolean {
    return Array.from(inner.querySelectorAll<HTMLElement>("img, table, br, hr, [data-hwe-manual-page-break='true']"))
      .some((element) => {
        try {
          return range.intersectsNode(element);
        } catch {
          return false;
        }
      });
  }
}
