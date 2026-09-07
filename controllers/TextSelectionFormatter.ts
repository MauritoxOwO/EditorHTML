export type TextCase = "uppercase" | "lowercase";

export class TextSelectionFormatter {
  applyFontSize(size: string): HTMLElement[] {
    const selection = window.getSelection();
    if (!size || !selection || selection.rangeCount === 0 || selection.isCollapsed) return [];

    const affectedElements: HTMLElement[] = [];
    for (let index = 0; index < selection.rangeCount; index++) {
      const targets = this.getSelectedTextTargets(selection.getRangeAt(index));
      targets.forEach((target) => {
        const element = this.wrapTextNodeTarget(target, size);
        if (element) affectedElements.push(element);
      });
    }

    return affectedElements;
  }

  transformSelectedTextCase(textCase: TextCase): HTMLElement[] {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0 || selection.isCollapsed) return [];

    const affectedElements = new Set<HTMLElement>();
    for (let index = 0; index < selection.rangeCount; index++) {
      const targets = this.getSelectedTextTargets(selection.getRangeAt(index));
      targets.forEach((target) => {
        const element = this.transformTextTarget(target, textCase);
        if (element) affectedElements.add(element);
      });
    }

    return Array.from(affectedElements);
  }

  private getSelectedTextTargets(range: Range): TextTarget[] {
    const root =
      range.commonAncestorContainer.nodeType === Node.ELEMENT_NODE
        ? (range.commonAncestorContainer as HTMLElement)
        : range.commonAncestorContainer.parentElement;
    if (!root) return [];

    const targets: TextTarget[] = [];
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode: (node) => {
        const textNode = node as Text;
        const parent = textNode.parentElement;
        if (!textNode.data || !parent || parent.closest("[contenteditable='false']")) {
          return NodeFilter.FILTER_REJECT;
        }

        return this.rangeIntersectsNode(range, textNode)
          ? NodeFilter.FILTER_ACCEPT
          : NodeFilter.FILTER_REJECT;
      },
    });

    let current = walker.nextNode() as Text | null;
    while (current) {
      const target = this.getSelectedTextTarget(range, current);
      if (target) targets.push(target);
      current = walker.nextNode() as Text | null;
    }

    return targets;
  }

  private getSelectedTextTarget(range: Range, node: Text): TextTarget | null {
    const textLength = node.data.length;
    const start = range.startContainer === node ? range.startOffset : 0;
    const end = range.endContainer === node ? range.endOffset : textLength;
    if (start >= end) return null;

    return { node, start, end };
  }

  private transformTextTarget(target: TextTarget, textCase: TextCase): HTMLElement | null {
    const { node, start, end } = target;
    const parent = node.parentElement;
    if (!parent) return null;

    const selectedText = node.data.slice(start, end);
    const transformedText =
      textCase === "uppercase"
        ? selectedText.toLocaleUpperCase("es-ES")
        : selectedText.toLocaleLowerCase("es-ES");
    if (transformedText === selectedText) return null;

    node.replaceData(start, end - start, transformedText);
    return parent;
  }

  private wrapTextNodeTarget(target: TextTarget, size: string): HTMLElement | null {
    const { node, start, end } = target;
    const parent = node.parentNode;
    if (!parent) return null;

    let selectedNode = node;
    if (end < node.data.length) {
      node.splitText(end);
    }
    if (start > 0) {
      selectedNode = node.splitText(start);
    }

    const span = document.createElement("span");
    span.style.fontSize = size;
    span.setAttribute("data-hwe-user-font-size", "true");
    parent.insertBefore(span, selectedNode);
    span.appendChild(selectedNode);
    return span;
  }

  private rangeIntersectsNode(range: Range, node: Node): boolean {
    try {
      return range.intersectsNode(node);
    } catch {
      return false;
    }
  }
}

interface TextTarget {
  node: Text;
  start: number;
  end: number;
}
