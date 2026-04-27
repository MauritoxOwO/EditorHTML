export interface CaretSnapshot {
  path: number[];   
  offset: number;   
}

export class CaretManager {

  static save(container: HTMLElement): CaretSnapshot | null {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return null;

    const range = selection.getRangeAt(0);
    if (!container.contains(range.startContainer)) return null;

    return {
      path: this.getNodePath(container, range.startContainer),
      offset: range.startOffset,
    };
  }

  static restore(
    container: HTMLElement,
    snapshot: CaretSnapshot | null
  ): void {
    if (!snapshot) return;

    const targetNode = this.resolveNodePath(container, snapshot.path);
    if (!targetNode) return;

    const offset = Math.min(
      snapshot.offset,
      targetNode.textContent?.length ?? 0
    );

    const range = document.createRange();
    range.setStart(targetNode, offset);
    range.collapse(true);

    const selection = window.getSelection();
    if (!selection) return;

    selection.removeAllRanges();
    selection.addRange(range);
  }


  private static getNodePath(
    root: Node,
    node: Node
  ): number[] {
    const path: number[] = [];
    let current: Node | null = node;

    while (current && current !== root) {
      const parent: Node | null = current.parentNode;
      if (!parent) break;

      const index = Array.prototype.indexOf.call(
        parent.childNodes,
        current
      );
      path.unshift(index);
      current = parent;
    }

    return path;
  }

  private static resolveNodePath(
    root: Node,
    path: number[]
  ): Text | null {
    let current: Node | null = root;

    for (const index of path) {
      if (!current || !current.childNodes[index]) {
        return this.findFirstTextNode(root);
      }
      current = current.childNodes[index];
    }

    if (current.nodeType === Node.TEXT_NODE) {
      return current as Text;
    }

    return this.findFirstTextNode(current);
  }

  /**
   * Busca el primer nodo de texto descendiente.
   * Fallback cuando el nodo original ya no existe.
   */
  private static findFirstTextNode(node: Node): Text | null {
    const walker = document.createTreeWalker(
      node,
      NodeFilter.SHOW_TEXT,
      null
    );

    return walker.nextNode() as Text | null;
  }
}