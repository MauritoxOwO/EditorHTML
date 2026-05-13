import { hweDebugLog } from "../debug/DebugLogger";
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

    const startedAt = performance.now();
    const flowId = this.ensureTextFlowId(block);
    const overflowBlock = block.cloneNode(false) as HTMLElement;
    overflowBlock.setAttribute("data-hwe-text-flow-id", flowId);
    overflowBlock.setAttribute("data-hwe-text-fragment", "true");
    targetInner.insertBefore(overflowBlock, targetInner.firstChild);

    let movedAny = false;
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

    const clone = element.cloneNode(false) as HTMLElement;
    const moved = this.moveLastInlinePiece(element, clone);

    if (!moved) {
      target.insertBefore(element, target.firstChild);
      return true;
    }

    target.insertBefore(clone, target.firstChild);
    if (isEmptyNode(element, false)) element.remove();
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
}
