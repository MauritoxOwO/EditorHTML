export class KeepTogetherController {
  groupFlowNodes(nodes: ChildNode[]): ChildNode[] {
    const grouped: ChildNode[] = [];
    let index = 0;

    while (index < nodes.length) {
      const current = nodes[index];
      const next = nodes[index + 1];

      if (this.shouldKeepWithNext(current, next)) {
        const group = document.createElement("div");
        group.className = "hwe-keep-together";
        group.setAttribute("data-hwe-keep-together", "true");
        group.setAttribute("data-hwe-generated-wrapper", "true");

        while (grouped.length > 0 && this.isUserBlankBlock(grouped[grouped.length - 1])) {
          group.insertBefore(grouped.pop()!, group.firstChild);
        }

        group.appendChild(current);
        group.appendChild(next);
        grouped.push(group);
        index += 2;
        continue;
      }

      if (this.shouldSoftKeepWithTable(current, next)) {
        (current as HTMLElement).setAttribute("data-hwe-keep-with-next", "true");
      }

      grouped.push(current);
      index++;
    }

    return grouped;
  }

  unwrapGeneratedGroup(group: HTMLElement): void {
    if (!group.parentNode || !this.isKeepTogetherGroup(group)) return;

    const parent = group.parentNode;
    while (group.firstChild) {
      parent.insertBefore(group.firstChild, group);
    }
    parent.removeChild(group);
  }

  isKeepTogetherGroup(element: HTMLElement): boolean {
    return element.getAttribute("data-hwe-keep-together") === "true";
  }

  private shouldKeepWithNext(current: ChildNode | undefined, next: ChildNode | undefined): boolean {
    if (!current || !next) return false;
    if (current.nodeType !== Node.ELEMENT_NODE || next.nodeType !== Node.ELEMENT_NODE) return false;

    const currentElement = current as HTMLElement;
    const nextElement = next as HTMLElement;
    return this.isKeepWithNextTrigger(currentElement) && this.isKeepTogetherTarget(nextElement);
  }

  private shouldSoftKeepWithTable(
    current: ChildNode | undefined,
    next: ChildNode | undefined
  ): boolean {
    if (!current || !next) return false;
    if (current.nodeType !== Node.ELEMENT_NODE || next.nodeType !== Node.ELEMENT_NODE) return false;

    const currentElement = current as HTMLElement;
    const nextElement = next as HTMLElement;
    return this.isKeepWithNextTrigger(currentElement) && this.isTableTarget(nextElement);
  }

  private isKeepWithNextTrigger(element: HTMLElement): boolean {
    if (element.getAttribute("data-hwe-keep-with-next") === "true") return true;
    if (element.hasAttribute("data-hwe-user-blank")) return false;
    if (element.querySelector("img, table, figure, video, canvas, svg")) return false;

    const text = (element.textContent ?? "").replace(/\u00a0/g, " ").trim();
    if (!text || text.length > 240) return false;
    if (/^H[1-6]$/.test(element.tagName)) return true;
    if (/caption|titulo|title|heading|epigrafe/i.test(element.className)) return true;
    if (element.querySelector("b, strong")) return true;

    return ["P", "DIV"].includes(element.tagName);
  }

  private isKeepTogetherTarget(element: HTMLElement): boolean {
    if (["IMG", "FIGURE", "VIDEO", "CANVAS", "SVG"].includes(element.tagName)) {
      return true;
    }

    const meaningfulChildren = Array.from(element.children).filter((child) => {
      const text = (child.textContent ?? "").replace(/\u00a0/g, " ").trim();
      return text || child.querySelector("img, table, figure, video, canvas, svg");
    });

    return (
      meaningfulChildren.length === 1 &&
      ["IMG", "FIGURE", "VIDEO", "CANVAS", "SVG"].includes(
        meaningfulChildren[0].tagName
      )
    );
  }

  private isTableTarget(element: HTMLElement): boolean {
    return !!this.getSingleTableTarget(element);
  }

  private getSingleTableTarget(element: HTMLElement): HTMLElement | null {
    if (element.tagName === "TABLE") return element;

    const meaningfulChildren = Array.from(element.children).filter((child) => {
      const text = (child.textContent ?? "").replace(/\u00a0/g, " ").trim();
      return text || child.querySelector("img, table, figure, video, canvas, svg");
    });

    if (meaningfulChildren.length !== 1) return null;

    const onlyChild = meaningfulChildren[0] as HTMLElement;
    return onlyChild.tagName === "TABLE" ? onlyChild : null;
  }

  private isUserBlankBlock(node: ChildNode | undefined): boolean {
    return (
      node?.nodeType === Node.ELEMENT_NODE &&
      (node as HTMLElement).hasAttribute("data-hwe-user-blank")
    );
  }
}

export function unwrapGeneratedKeepTogetherGroups(root: HTMLElement): void {
  root.querySelectorAll<HTMLElement>("[data-hwe-generated-wrapper='true']").forEach((group) => {
    const parent = group.parentNode;
    if (!parent) return;

    while (group.firstChild) {
      parent.insertBefore(group.firstChild, group);
    }
    parent.removeChild(group);
  });
}
