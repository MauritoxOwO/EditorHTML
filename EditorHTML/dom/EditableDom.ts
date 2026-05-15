export const EDITABLE_BLANK_BLOCK_SELECTOR =
  "p, div, li, h1, h2, h3, h4, h5, h6, blockquote, pre";

const BLANK_BLOCK_TAGS = new Set([
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

const SPLITTABLE_CONTAINER_TAGS = new Set([
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

export function getMeaningfulChildren(
  container: HTMLElement,
  preserveEditableBlankBlocks = false
): ChildNode[] {
  return Array.from(container.childNodes).filter(
    (node) => !isEmptyNode(node, preserveEditableBlankBlocks)
  );
}

export function isEmptyNode(
  node: ChildNode,
  preserveEditableBlankBlocks = false
): boolean {
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
  if (preserveEditableBlankBlocks && isEditableBlankBlock(element)) return false;

  return (
    ["P", "DIV", "SECTION", "ARTICLE", "SPAN"].includes(element.tagName) &&
    (element.textContent ?? "").replace(/\u00a0/g, " ").trim() === "" &&
    Array.from(element.childNodes).every((child) => isEmptyNode(child, false))
  );
}

export function isEditableBlankBlock(element: HTMLElement): boolean {
  if (!BLANK_BLOCK_TAGS.has(element.tagName)) return false;
  if ((element.textContent ?? "").replace(/\u00a0/g, " ").trim() !== "") return false;

  return Array.from(element.childNodes).every((child) => {
    if (child.nodeType === Node.TEXT_NODE) {
      return (child.textContent ?? "").replace(/\u00a0/g, " ").trim() === "";
    }

    return child.nodeType === Node.ELEMENT_NODE && (child as HTMLElement).tagName === "BR";
  });
}

export function isSplittableContainer(element: HTMLElement): boolean {
  if (!SPLITTABLE_CONTAINER_TAGS.has(element.tagName)) return false;
  if (element.classList.contains("hwe-page") || element.classList.contains("hwe-page-inner")) {
    return false;
  }

  return getMeaningfulChildren(element).length > 0;
}

export function removeComments(root: HTMLElement): void {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_COMMENT);
  const comments: Node[] = [];

  while (walker.nextNode()) {
    comments.push(walker.currentNode);
  }

  comments.forEach((comment) => comment.parentNode?.removeChild(comment));
}

export function unwrapElement(element: HTMLElement): void {
  if (!element.parentNode) return;

  const parent = element.parentNode;
  while (element.firstChild) {
    parent.insertBefore(element.firstChild, element);
  }
  parent.removeChild(element);
}
