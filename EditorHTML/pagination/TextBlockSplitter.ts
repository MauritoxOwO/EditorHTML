import { hweDebugLog } from "../debug/DebugLogger";
import { ensureFlowIdentity, INLINE_FLOW_ID } from "./FlowIdentity";
import { getPageContentBounds } from "./PageContentGeometry";
import {
  isAtomicElement,
  isEmptyNode,
  isSplittableTextBlock,
  pageOverflows,
} from "./PaginatorDom";

const MAX_SPLIT_MS = 180;

export class TextBlockSplitter {
  private static textFlowCounter = 0;

  splitTextBlock(
    block: HTMLElement,
    targetInner: HTMLElement,
    page: HTMLElement
  ): boolean {
    if (!isSplittableTextBlock(block)) return false;

    const inner = page.querySelector<HTMLElement>(".hwe-page-inner");
    if (inner) {
      const bounds = getPageContentBounds(inner);
      const rect = block.getBoundingClientRect();
      if (rect.height > 0 && rect.top >= bounds.bottom) {
        targetInner.insertBefore(block, targetInner.firstChild);
        return true;
      }
    }

    const startedAt = performance.now();
    const flowId = this.ensureTextFlowId(block);
    const overflowBlock = block.cloneNode(false) as HTMLElement;
    overflowBlock.setAttribute("data-hwe-text-flow-id", flowId);
    overflowBlock.setAttribute("data-hwe-text-fragment", "true");
    targetInner.insertBefore(overflowBlock, targetInner.firstChild);

    let movedAny = this.moveMeasuredTail(block, overflowBlock, page);
    if (movedAny && isEmptyNode(block, false)) {
      block.remove();
      return true;
    }
    let safety = 0;

    while (pageOverflows(page) && safety++ < 300) {
      if (performance.now() - startedAt > MAX_SPLIT_MS) {
        hweDebugLog("paginator.splitTextBlock.timeBudgetExceeded", {
          elapsedMs: Math.round((performance.now() - startedAt) * 100) / 100,
          movedAny,
          tagName: block.tagName,
          textLength: (block.textContent ?? "").length,
        });
        break;
      }

      const moved = this.moveLastInlinePiece(block, overflowBlock);
      if (!moved) break;
      movedAny = true;

      if (isEmptyNode(block, false)) {
        block.remove();
        break;
      }
    }

    if (!movedAny || isEmptyNode(overflowBlock, false)) {
      overflowBlock.remove();
      return false;
    }

    return true;
  }

  private ensureTextFlowId(block: HTMLElement): string {
    const existingId = block.getAttribute("data-hwe-text-flow-id");
    if (existingId) return existingId;

    const id = `hwe-text-${Date.now().toString(36)}-${TextBlockSplitter.textFlowCounter++}`;
    block.setAttribute("data-hwe-text-flow-id", id);
    block.setAttribute("data-hwe-text-fragment", "true");
    return id;
  }

  private moveMeasuredTail(block: HTMLElement, target: HTMLElement, page: HTMLElement): boolean {
    const inner = page.querySelector<HTMLElement>(".hwe-page-inner");
    const probe = document.createRange();
    if (!inner || typeof probe.getBoundingClientRect !== "function") return false;
    const words: { node: Text; start: number; end: number }[] = [];
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    let node: Node | null;
    while ((node = walker.nextNode())) {
      const text = node as Text;
      for (const match of text.data.matchAll(/\S+\s*/g)) {
        words.push({ node: text, start: match.index!, end: match.index! + match[0].length });
      }
    }
    if (words.length < 8) return false;
    const bounds = getPageContentBounds(inner);
    let low = 0;
    let high = words.length;
    // Read the existing line layout without mutating it word by word. A single
    // extracted tail retains nested formatting, images, markers and text order.
    while (low < high) {
      const middle = (low + high) >>> 1;
      const word = words[middle];
      probe.selectNodeContents(block);
      probe.setEnd(word.node, word.end);
      const rect = probe.getBoundingClientRect();
      if (!rect.height) return false;
      if (rect.bottom <= bounds.bottom + bounds.scale) low = middle + 1;
      else high = middle;
    }
    if (low === words.length) return false;
    const firstOverflow = words[low];
    for (let ancestor = firstOverflow.node.parentElement; ancestor && ancestor !== block; ancestor = ancestor.parentElement) {
      ensureFlowIdentity(ancestor, INLINE_FLOW_ID);
    }
    probe.selectNodeContents(block);
    probe.setStart(firstOverflow.node, firstOverflow.start);
    target.appendChild(probe.extractContents());
    return true;
  }

  private moveLastInlinePiece(source: HTMLElement, target: HTMLElement): boolean {
    const child = source.lastChild;
    if (!child) return false;

    if (child.nodeType === Node.TEXT_NODE) {
      return this.moveLastWordFromTextNode(child as Text, target);
    }

    if (isEmptyNode(child)) {
      child.remove();
      return true;
    }

    if (child.nodeType !== Node.ELEMENT_NODE) {
      target.insertBefore(child, target.firstChild);
      return true;
    }

    const element = child as HTMLElement;
    if (isAtomicElement(element)) {
      target.insertBefore(element, target.firstChild);
      return true;
    }

    ensureFlowIdentity(element, INLINE_FLOW_ID);
    const clone = element.cloneNode(false) as HTMLElement;
    const moved = this.moveLastInlinePiece(element, clone);

    if (!moved) {
      target.insertBefore(element, target.firstChild);
      return true;
    }

    if (clone.hasChildNodes()) target.insertBefore(clone, target.firstChild);
    if (!element.hasChildNodes() || isEmptyNode(element, false)) element.remove();
    return true;
  }

  private moveLastWordFromTextNode(textNode: Text, target: HTMLElement): boolean {
    const text = textNode.textContent ?? "";

    if (!text) {
      textNode.remove();
      return true;
    }

    if (/^\s+$/.test(text)) {
      target.insertBefore(textNode, target.firstChild);
      return true;
    }

    const match = /(\s*\S+\s*)$/.exec(text);
    if (!match) return false;

    if (match.index <= 0) {
      // Pagination and caret markers can leave one word per Text node. Keep
      // moving those nodes into this fragment instead of starting a new block
      // for each word (whose temporary margins distort page measurements).
      target.insertBefore(textNode, target.firstChild);
      return true;
    }

    const prefix = text.slice(0, match.index);
    const suffix = text.slice(match.index);

    textNode.textContent = prefix;
    target.insertBefore(document.createTextNode(suffix), target.firstChild);
    return true;
  }
}
