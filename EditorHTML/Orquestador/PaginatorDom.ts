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

const TABLE_STRUCTURE_TAGS = new Set([
  "TABLE",
  "THEAD",
  "TBODY",
  "TFOOT",
  "TR",
  "TD",
  "TH",
  "COLGROUP",
  "COL",
]);

const SPLITTABLE_TEXT_TAGS = new Set([
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

export function getInner(page: HTMLElement | undefined): HTMLElement | null {
  return page?.querySelector(".hwe-page-inner") ?? null;
}

export function pageOverflows(page: HTMLElement): boolean {
  const inner = getInner(page);
  if (!inner) return false;

  const scrollOverflows = inner.scrollHeight > inner.clientHeight + 1;
  const contentBottom = getContentBottom(inner);
  if (contentBottom === null) return scrollOverflows;

  return contentBottom > getContentLimitBottom(inner) + 1 || scrollOverflows;
}

export function getContentLimitBottom(inner: HTMLElement): number {
  const styles = getComputedStyle(inner);
  const paddingBottom = parseFloat(styles.paddingBottom) || 0;
  return inner.getBoundingClientRect().bottom - Math.min(paddingBottom, 12);
}

export function getContentHeight(page: HTMLElement): number {
  const inner = getInner(page);
  if (!inner) return page.clientHeight;

  const styles = getComputedStyle(inner);
  const paddingTop = parseFloat(styles.paddingTop) || 0;
  const paddingBottom = parseFloat(styles.paddingBottom) || 0;
  return Math.max(0, inner.clientHeight - paddingTop - paddingBottom);
}

export function getContentBottom(inner: HTMLElement): number | null {
  const children = getMeaningfulChildren(inner);
  if (children.length === 0) return null;

  return children.reduce<number | null>((bottom, child) => {
    const childBottom = getNodeBottom(child);
    if (childBottom === null) return bottom;
    return bottom === null ? childBottom : Math.max(bottom, childBottom);
  }, null);
}

export function getNodeBottom(node: ChildNode): number | null {
  if (node.nodeType === Node.ELEMENT_NODE) {
    return (node as HTMLElement).getBoundingClientRect().bottom;
  }

  const range = document.createRange();
  range.selectNodeContents(node);
  const rect = range.getBoundingClientRect();
  return rect.width > 0 || rect.height > 0 ? rect.bottom : null;
}

export function getMeaningfulChildren(container: HTMLElement): ChildNode[] {
  return Array.from(container.childNodes).filter((node) => !isEmptyNode(node));
}

export function isTableElement(node: ChildNode): boolean {
  return node.nodeType === Node.ELEMENT_NODE && (node as HTMLElement).tagName === "TABLE";
}

export function isTableFlowWrapper(element: HTMLElement): boolean {
  if (TABLE_STRUCTURE_TAGS.has(element.tagName)) return false;

  const children = getMeaningfulChildren(element);
  if (children.length !== 1 || children[0].nodeType !== Node.ELEMENT_NODE) return false;

  const onlyChild = children[0] as HTMLElement;
  if (onlyChild.tagName === "TABLE") return true;
  if (!SPLITTABLE_CONTAINER_TAGS.has(onlyChild.tagName)) return false;

  return isTableFlowWrapper(onlyChild);
}

export function isSplittableContainer(
  element: HTMLElement,
  isKeepTogetherGroup: (candidate: HTMLElement) => boolean
): boolean {
  if (element.classList.contains("hwe-page") || element.classList.contains("hwe-page-inner")) {
    return false;
  }
  if (isKeepTogetherGroup(element)) return false;
  if (isTableFlowWrapper(element)) return true;
  if (!SPLITTABLE_CONTAINER_TAGS.has(element.tagName)) return false;

  return getMeaningfulChildren(element).length > 0;
}

export function shouldPreserveContainerShell(element: HTMLElement): boolean {
  if (element.getAttribute("data-hwe-page-break") === "before") return false;
  if (shouldFlattenContainerShell(element)) return false;

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

export function shouldFlattenContainerShell(element: HTMLElement): boolean {
  if (!SPLITTABLE_CONTAINER_TAGS.has(element.tagName)) return false;
  if (element.classList.contains("hwe-page") || element.classList.contains("hwe-page-inner")) {
    return false;
  }
  if (isGeneratedOrAtomicShell(element)) return false;
  if (isTableFlowWrapper(element)) return false;

  const children = getMeaningfulChildren(element);
  if (children.length === 0) return false;
  if (children.length > 1) return true;

  const onlyChild = children[0];
  if (onlyChild.nodeType !== Node.ELEMENT_NODE) return false;

  const onlyElement = onlyChild as HTMLElement;
  if (isGeneratedOrAtomicShell(onlyElement)) return false;
  if (isTableFlowWrapper(onlyElement)) return false;

  return (
    SPLITTABLE_CONTAINER_TAGS.has(onlyElement.tagName) &&
    getMeaningfulChildren(onlyElement).length > 0
  );
}

function isGeneratedOrAtomicShell(element: HTMLElement): boolean {
  return (
    element.getAttribute("data-hwe-keep-together") === "true" ||
    element.classList.contains("hwe-keep-together") ||
    element.classList.contains("hwe-ocr-wrapper")
  );
}

export function isSplittableTextBlock(element: HTMLElement): boolean {
  if (!SPLITTABLE_TEXT_TAGS.has(element.tagName)) return false;
  if (element.querySelector("table, tr, td, th")) return false;
  return (element.textContent ?? "").replace(/\u00a0/g, " ").trim().length > 0;
}

export function isAtomicElement(element: HTMLElement): boolean {
  return ["IMG", "TABLE", "TR", "TD", "TH", "VIDEO", "CANVAS", "SVG", "BR"].includes(
    element.tagName
  );
}

export function isEmptyNode(node: ChildNode, preserveEditableBlankBlocks = false): boolean {
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
