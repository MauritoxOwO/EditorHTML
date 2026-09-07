import { CaretManager } from "../pagination/CaretManager";

type ListElement = HTMLUListElement | HTMLOListElement;
type ListTagName = "UL" | "OL";

export interface ListCommandControllerOptions {
  rootProvider: () => HTMLElement | null;
  getActiveEditable: () => HTMLElement | null;
  onListChanged: (affectedElement: HTMLElement) => void;
}

export class ListCommandController {
  constructor(private readonly options: ListCommandControllerOptions) {}

  toggleUnorderedList(): boolean {
    return this.toggleList("UL");
  }

  toggleOrderedList(): boolean {
    return this.toggleList("OL");
  }

  handleKeyDown(event: KeyboardEvent): boolean {
    if (
      event.defaultPrevented ||
      event.key !== "Tab" ||
      event.ctrlKey ||
      event.altKey ||
      event.metaKey
    ) {
      return false;
    }

    const context = this.getActiveListItem(true);
    if (!context) return false;

    const canChange = event.shiftKey
      ? this.canOutdent(context.item, context.list)
      : this.canIndent(context.item);
    if (!canChange) return false;

    const marker = CaretManager.createMarker(context.editable);
    if (!marker) return false;

    const changed = event.shiftKey
      ? this.outdentItem(context.item, context.list)
      : this.indentItem(context.item, context.list);

    if (!changed) {
      CaretManager.restoreMarker(marker, context.editable);
      return false;
    }

    event.preventDefault();
    event.stopPropagation();
    this.removeEmptyLists(context.editable);
    CaretManager.restoreMarker(marker, context.editable);
    this.options.onListChanged(context.item);
    return true;
  }

  private toggleList(tagName: ListTagName): boolean {
    const editable = this.options.getActiveEditable();
    const root = this.options.rootProvider();
    const selection = window.getSelection();
    if (!editable || !root || !selection || selection.rangeCount === 0) return false;

    const anchorNode = selection.anchorNode;
    if (!anchorNode || !root.contains(anchorNode) || !editable.contains(anchorNode)) return false;

    const beforeHtml = editable.innerHTML;
    const command = tagName === "UL" ? "insertUnorderedList" : "insertOrderedList";
    document.execCommand(command, false);
    if (editable.innerHTML === beforeHtml) return false;

    this.removeEmptyLists(editable);
    const affectedElement = this.getActiveListItem(false)?.item ?? this.getActiveBlock(editable);
    this.options.onListChanged(affectedElement);
    return true;
  }

  private indentItem(item: HTMLLIElement, sourceList: ListElement): boolean {
    const previousItem = item.previousElementSibling;
    if (!(previousItem instanceof HTMLLIElement)) return false;

    let nestedList = this.getDirectChildList(previousItem, sourceList.tagName as ListTagName);
    if (!nestedList) {
      nestedList = this.cloneListShell(sourceList);
      previousItem.appendChild(nestedList);
    }

    nestedList.appendChild(item);
    return true;
  }

  private outdentItem(item: HTMLLIElement, sourceList: ListElement): boolean {
    const parentItem = sourceList.parentElement;
    if (!(parentItem instanceof HTMLLIElement)) return false;

    const outerList = parentItem.parentElement;
    if (!this.isListElement(outerList)) return false;

    const followingItems: HTMLLIElement[] = [];
    let sibling = item.nextElementSibling;
    while (sibling instanceof HTMLLIElement) {
      followingItems.push(sibling);
      sibling = sibling.nextElementSibling;
    }

    if (followingItems.length > 0) {
      let continuationList = this.getDirectChildList(
        item,
        sourceList.tagName as ListTagName
      );
      if (!continuationList) {
        continuationList = this.cloneListShell(sourceList);
        item.appendChild(continuationList);
      }
      followingItems.forEach((followingItem) => continuationList?.appendChild(followingItem));
    }

    outerList.insertBefore(item, parentItem.nextSibling);
    if (!this.hasDirectListItems(sourceList)) sourceList.remove();
    return true;
  }

  private canIndent(item: HTMLLIElement): boolean {
    return item.previousElementSibling instanceof HTMLLIElement;
  }

  private canOutdent(item: HTMLLIElement, sourceList: ListElement): boolean {
    return (
      sourceList.contains(item) &&
      sourceList.parentElement instanceof HTMLLIElement &&
      this.isListElement(sourceList.parentElement.parentElement)
    );
  }

  private getActiveListItem(
    requireCollapsedSelection: boolean
  ): { editable: HTMLElement; item: HTMLLIElement; list: ListElement } | null {
    const editable = this.options.getActiveEditable();
    const root = this.options.rootProvider();
    const selection = window.getSelection();
    if (!editable || !root || !selection || selection.rangeCount === 0) return null;
    if (requireCollapsedSelection && !selection.isCollapsed) return null;

    const anchorItem = this.getClosestListItem(selection.anchorNode);
    const focusItem = this.getClosestListItem(selection.focusNode);
    if (!anchorItem || anchorItem !== focusItem) return null;
    if (!root.contains(anchorItem) || !editable.contains(anchorItem)) return null;

    const list = anchorItem.parentElement;
    if (!this.isListElement(list)) return null;
    return { editable, item: anchorItem, list };
  }

  private getClosestListItem(node: Node | null): HTMLLIElement | null {
    if (!node) return null;
    const element = node.nodeType === Node.ELEMENT_NODE ? (node as Element) : node.parentElement;
    return element?.closest<HTMLLIElement>("li") ?? null;
  }

  private getActiveBlock(editable: HTMLElement): HTMLElement {
    const selection = window.getSelection();
    const anchorNode = selection?.anchorNode ?? null;
    const element =
      anchorNode?.nodeType === Node.ELEMENT_NODE
        ? (anchorNode as HTMLElement)
        : anchorNode?.parentElement ?? null;
    const block = element?.closest<HTMLElement>(
      "p, div, li, h1, h2, h3, h4, h5, h6, blockquote, pre, ul, ol"
    );
    return block && editable.contains(block) ? block : editable;
  }

  private getDirectChildList(
    item: HTMLLIElement,
    tagName: ListTagName
  ): ListElement | null {
    return (
      Array.from(item.children)
        .reverse()
        .find((child): child is ListElement => child.tagName === tagName) ?? null
    );
  }

  private cloneListShell(source: ListElement): ListElement {
    const clone = source.cloneNode(false) as ListElement;
    clone.removeAttribute("start");
    clone.removeAttribute("data-hwe-list-flow-id");
    clone.removeAttribute("data-hwe-list-fragment");
    return clone;
  }

  private removeEmptyLists(root: HTMLElement): void {
    Array.from(root.querySelectorAll<ListElement>("ul, ol"))
      .reverse()
      .forEach((list) => {
        if (!this.hasDirectListItems(list)) list.remove();
      });
  }

  private hasDirectListItems(list: ListElement): boolean {
    return Array.from(list.children).some((child) => child instanceof HTMLLIElement);
  }

  private isListElement(element: Element | null): element is ListElement {
    return element instanceof HTMLUListElement || element instanceof HTMLOListElement;
  }
}
