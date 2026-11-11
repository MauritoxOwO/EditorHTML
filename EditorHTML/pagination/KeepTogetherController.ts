import { isEditableBlankBlock, isManualPageBreak, isSplittableTextBlock } from "./PaginatorDom";

export class KeepTogetherController {
  private readonly keepWithNextStyles = new Set<string>();

  setKeepWithNextStyles(classNames: readonly string[]): void {
    this.keepWithNextStyles.clear();
    classNames.forEach((className) => this.keepWithNextStyles.add(className));
  }

  groupFlowNodes(nodes: ChildNode[]): ChildNode[] {
    const grouped: ChildNode[] = [];
    let index = 0;

    while (index < nodes.length) {
      const current = nodes[index];
      const next = nodes[index + 1];

      const titleGroupEnd = this.findTitleGroupEnd(nodes, index);
      if (titleGroupEnd > index) {
        const group = document.createElement("div");
        group.className = "hwe-keep-together";
        group.setAttribute("data-hwe-keep-together", "true");
        group.setAttribute("data-hwe-generated-wrapper", "true");
        group.setAttribute("data-hwe-title-group", "true");
        nodes.slice(index, titleGroupEnd + 1).forEach((node) => group.appendChild(node));
        grouped.push(group);
        index = titleGroupEnd + 1;
        continue;
      }

      if (this.shouldKeepWithNext(current, next)) {
        const group = document.createElement("div");
        group.className = "hwe-keep-together";
        group.setAttribute("data-hwe-keep-together", "true");
        group.setAttribute("data-hwe-generated-wrapper", "true");

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

  isTitleGroup(element: HTMLElement): boolean {
    return this.isKeepTogetherGroup(element) &&
      element.getAttribute("data-hwe-title-group") === "true";
  }

  preserveTitleBlankBarriers(container: HTMLElement): void {
    let afterTitle = false;
    Array.from(container.childNodes).forEach((node) => {
      if (this.isFormattingWhitespace(node)) return;
      if (node.nodeType === Node.ELEMENT_NODE) {
        const element = node as HTMLElement;
        if (element.matches("div, section, article") && !isManualPageBreak(element) &&
            !element.matches("[hidden], [contenteditable='false'], [data-hwe-api-header], [data-hwe-dynamic-header], [data-hwe-runtime-page-header]")) {
          this.preserveTitleBlankBarriers(element);
        }
      }
      const block = this.getFlowTextBlock(node, true);
      if (afterTitle && block && isEditableBlankBlock(block)) {
        block.setAttribute("data-hwe-user-blank", "true");
      }
      afterTitle = this.isTitleBlock(node);
    });
  }

  getFlowTextBlock(node: ChildNode | undefined, includeBlank = false): HTMLElement | null {
    if (!node || node.nodeType !== Node.ELEMENT_NODE) return null;
    const element = node as HTMLElement;
    if (isManualPageBreak(element) || element.hidden ||
        element.getAttribute("contenteditable") === "false" ||
        (!includeBlank && element.hasAttribute("data-hwe-user-blank")) ||
        element.matches("[data-hwe-api-header], [data-hwe-dynamic-header], [data-hwe-runtime-page-header]")) {
      return null;
    }
    if (includeBlank && isEditableBlankBlock(element)) return element;
    if (isSplittableTextBlock(element)) return element;
    if (!element.matches("div, section, article")) return null;
    const children = Array.from(element.childNodes).filter((child) => !this.isFormattingWhitespace(child));
    return children.length === 1 ? this.getFlowTextBlock(children[0], includeBlank) : null;
  }

  private isTitleBlock(node: ChildNode | undefined): boolean {
    const block = this.getFlowTextBlock(node);
    if (!block) return false;
    const appliedStyle = block.getAttribute("data-hwe-paragraph-style");
    return appliedStyle !== null
      ? this.keepWithNextStyles.has(appliedStyle)
      : Array.from(block.classList).some((className) => this.keepWithNextStyles.has(className));
  }

  private findTitleGroupEnd(nodes: ChildNode[], start: number): number {
    if (!this.isTitleBlock(nodes[start])) return start;
    let index = start + 1;
    while (index < nodes.length) {
      if (this.isFormattingWhitespace(nodes[index]) || this.isTitleBlock(nodes[index])) {
        index++;
        continue;
      }
      const paragraph = this.getFlowTextBlock(nodes[index]);
      return paragraph && paragraph.matches("p, div") ? index : start;
    }
    return start;
  }

  private isFormattingWhitespace(node: ChildNode): boolean {
    return node.nodeType === Node.COMMENT_NODE ||
      (node.nodeType === Node.TEXT_NODE && !(node.textContent ?? "").trim());
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
