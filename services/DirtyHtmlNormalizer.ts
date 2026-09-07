import { unwrapElement } from "../dom/EditableDom";
import { unwrapGeneratedKeepTogetherGroups } from "../pagination/KeepTogetherController";
import {
  AUTO_IMAGE_MAX_HEIGHT,
  CONTAINER_FLOW_ID,
  GENERATED_ROW_GROUP,
  INLINE_FLOW_ID,
  ROW_GROUP_ID,
  ROW_GROUP_ORIGIN,
} from "../pagination/FlowIdentity";

const FLOW_ATTRIBUTES = [
  "data-hwe-text-flow-id",
  "data-hwe-table-flow-id",
  CONTAINER_FLOW_ID,
  INLINE_FLOW_ID,
];
const COMPARISON_ONLY_ATTRIBUTES = new Set([
  ...FLOW_ATTRIBUTES,
  "data-hwe-text-fragment",
  "data-hwe-table-fragment",
  "data-hwe-user-blank",
  "data-hwe-keep-with-next",
  "data-hwe-ocr-state",
  "data-hwe-large-image-state",
  AUTO_IMAGE_MAX_HEIGHT,
  GENERATED_ROW_GROUP,
  ROW_GROUP_ID,
  ROW_GROUP_ORIGIN,
  "contenteditable",
  "spellcheck",
]);
const LAYOUT_CLASSES = new Set([
  "hwe-image-selected",
  "hwe-table-selected",
  "hwe-text-flow-block",
  "hwe-image-flow-block",
  "hwe-table-flow-wrapper",
  "hwe-long-word-table",
  "hwe-table-compact",
  "hwe-table-dense",
  "hwe-table-ultra-dense",
]);

/** Canonical comparison copy. Never use this result as the document to save. */
export function normalizeDirtyHtml(html: string): string {
  // Template content belongs to an inert document: parsing comparison copies
  // must not load images, execute handlers or change the live editing surface.
  const template = document.createElement("template");
  template.innerHTML = html;
  const root = template.content.ownerDocument.createElement("div");
  root.appendChild(template.content);

  root.querySelectorAll(
    "[data-hwe-api-header='true'], [data-hwe-dynamic-header='true'], " +
    "[data-hwe-runtime-page-header='true'], [data-hwe-caret], [data-hwe-paste-marker]"
  ).forEach((element) => element.remove());
  unwrapGeneratedKeepTogetherGroups(root);
  root.querySelectorAll<HTMLElement>("*").forEach(cleanPresentationState);
  restoreRowGroups(root);
  root.querySelectorAll<HTMLTableElement>("table").forEach(coalesceRowGroups);

  // Keep explicit boundaries until fragments on either side have been joined.
  // Only direct wrappers emitted by collectHtml() are automatic page wrappers.
  const boundaries: Comment[] = [];
  root.querySelectorAll<HTMLElement>(
    "[data-hwe-document='true'] > div[data-hwe-page-break='before']" +
    ":not([data-hwe-manual-page-break='true'])"
  ).forEach((wrapper) => {
    const previous = wrapper.previousSibling;
    if (previous?.nodeType === Node.TEXT_NODE) {
      const separator = previous as Text;
      // Remove exactly the separator inserted by collectHtml().join("\n").
      if (separator.data.endsWith("\n")) separator.deleteData(separator.length - 1, 1);
      if (separator.length === 0) separator.remove();
    }
    const boundary = document.createComment("automatic page boundary");
    wrapper.before(boundary);
    boundaries.push(boundary);
    unwrapElement(wrapper);
  });
  boundaries.forEach(joinPageBoundary);

  // TextBlockSplitter can also clone nested inline runs several times within
  // one page. Merge only matching identities, never arbitrary equal-looking tags.
  mergeNestedClones(root);
  root.querySelectorAll<HTMLElement>("*").forEach((element) => {
    COMPARISON_ONLY_ATTRIBUTES.forEach((attribute) => element.removeAttribute(attribute));
    sortAttributes(element);
  });
  root.normalize();
  return root.innerHTML;
}

function cleanPresentationState(element: HTMLElement): void {
  const classes = Array.from(element.classList).filter((name) => !LAYOUT_CLASSES.has(name));
  if (classes.length) element.setAttribute("class", classes.sort().join(" "));
  else element.removeAttribute("class");

  // Ignore only a height still equal to the value set by our paginator.
  // User image width, alignment, src and other styles remain in the comparison.
  const automaticHeight = element.getAttribute(AUTO_IMAGE_MAX_HEIGHT);
  if (element.tagName === "IMG" && automaticHeight === element.style.maxHeight) {
    element.style.removeProperty("max-height");
  }
  if (element.hasAttribute("style")) {
    // CSSOM normalizes spacing; do not sort declarations (shorthands can overlap).
    const style = element.style.cssText;
    if (style) element.setAttribute("style", style);
    else element.removeAttribute("style");
  }
}

function attributesKey(element: Element): string {
  return JSON.stringify(Array.from(element.attributes)
    .filter((attribute) => !COMPARISON_ONLY_ATTRIBUTES.has(attribute.name))
    .map((attribute) => [attribute.name, attribute.value])
    .sort(([left], [right]) => left.localeCompare(right)));
}

function hasSameFlow(left: HTMLElement, right: HTMLElement, attributes = FLOW_ATTRIBUTES): boolean {
  if (left.tagName !== right.tagName || attributesKey(left) !== attributesKey(right)) return false;
  if (left.hasAttribute("data-hwe-manual-page-break") ||
      right.hasAttribute("data-hwe-manual-page-break")) return false;
  return attributes.some((attribute) => {
    const id = left.getAttribute(attribute);
    return !!id && id === right.getAttribute(attribute);
  });
}

function joinPageBoundary(boundary: Comment): void {
  const left = boundary.previousSibling;
  const right = boundary.nextSibling;
  if (left?.nodeType !== Node.ELEMENT_NODE || right?.nodeType !== Node.ELEMENT_NODE) {
    boundary.remove();
    return;
  }
  const target = left as HTMLElement;
  const source = right as HTMLElement;
  if (!hasSameFlow(target, source)) {
    boundary.remove();
    return;
  }
  if (target.tagName === "TABLE") {
    joinTables(target as HTMLTableElement, source as HTMLTableElement);
    boundary.remove();
    return;
  }
  // Move the boundary into a joined container to reconstruct nested fragments too.
  target.appendChild(boundary);
  while (source.firstChild) target.appendChild(source.firstChild);
  source.remove();
  joinPageBoundary(boundary);
}

function comparisonCopyKey(element: Element): string {
  const clone = element.cloneNode(true) as HTMLElement;
  [clone, ...Array.from(clone.querySelectorAll<HTMLElement>("*"))].forEach((child) => {
    COMPARISON_ONLY_ATTRIBUTES.forEach((attribute) => child.removeAttribute(attribute));
    sortAttributes(child);
  });
  return clone.outerHTML;
}

function joinTables(target: HTMLTableElement, source: HTMLTableElement): void {
  const sourceShell = Array.from(source.children).filter((child) =>
    ["CAPTION", "COLGROUP", "THEAD"].includes(child.tagName));
  const targetShell = Array.from(target.children).filter((child) =>
    ["CAPTION", "COLGROUP", "THEAD"].includes(child.tagName));

  // A repeated header/column definition may have been edited independently.
  // If it differs, retain both tables so that the real change is not hidden.
  for (const tag of ["CAPTION", "COLGROUP", "THEAD"]) {
    const sourceParts = sourceShell.filter((child) => child.tagName === tag);
    const targetParts = targetShell.filter((child) => child.tagName === tag);
    if (sourceParts.length && JSON.stringify(sourceParts.map(comparisonCopyKey)) !==
        JSON.stringify(targetParts.map(comparisonCopyKey))) return;
  }
  const targetColumns = targetShell.filter((child) => child.tagName === "COLGROUP");
  const sourceColumns = sourceShell.filter((child) => child.tagName === "COLGROUP");
  if (targetColumns.length !== sourceColumns.length) return;
  if (target.getAttribute("data-hwe-repeat-header") === "true" &&
      !!target.tHead !== !!source.tHead) return;

  // Preserve row-group attributes and nested tables. Never use querySelectorAll
  // ('tr') here: that would extract rows belonging to a table inside a cell.
  for (const group of Array.from(source.children)) {
    if (group.tagName !== "TBODY" && group.tagName !== "TFOOT") continue;
    const groupId = group.getAttribute(ROW_GROUP_ID);
    const destination = groupId
      ? Array.from(target.children).find((child) => child.getAttribute(ROW_GROUP_ID) === groupId)
      : target.lastElementChild;
    if (destination?.tagName === group.tagName && attributesKey(destination) === attributesKey(group) &&
        (!groupId || destination.getAttribute(ROW_GROUP_ID) === groupId)) {
      while (group.firstChild) destination.appendChild(group.firstChild);
      group.remove();
    } else {
      target.appendChild(group);
    }
  }
  source.remove();
}

function mergeNestedClones(parentElement: HTMLElement): void {
  let child = parentElement.firstChild;
  while (child) {
    const next = child.nextSibling;
    if (child.nodeType === Node.ELEMENT_NODE && next?.nodeType === Node.ELEMENT_NODE &&
        hasSameFlow(child as HTMLElement, next as HTMLElement, [INLINE_FLOW_ID, CONTAINER_FLOW_ID])) {
      const element = child as HTMLElement;
      const boundary = parentElement.ownerDocument.createComment("cloned container boundary");
      element.appendChild(boundary);
      while (next.firstChild) element.appendChild(next.firstChild);
      next.remove();
      joinPageBoundary(boundary);
      continue;
    }
    if (child.nodeType === Node.ELEMENT_NODE) mergeNestedClones(child as HTMLElement);
    child = next;
  }
}

type RowGroupOrigin = [string, "TBODY" | "TFOOT", [string, string][]];

function readRowGroupOrigin(value: string | null): RowGroupOrigin | null {
  if (!value) return null;
  try {
    const parsed: unknown = JSON.parse(value);
    if (!Array.isArray(parsed) || parsed.length !== 3 || typeof parsed[0] !== "string" ||
        !["TBODY", "TFOOT"].includes(parsed[1]) || !Array.isArray(parsed[2]) ||
        !parsed[2].every((attribute: unknown) => Array.isArray(attribute) && attribute.length === 2 &&
          attribute.every((part: unknown) => typeof part === "string"))) return null;
    return parsed as RowGroupOrigin;
  } catch {
    return null;
  }
}

function restoreRowGroups(root: HTMLElement): void {
  root.querySelectorAll<HTMLElement>(`tbody[${GENERATED_ROW_GROUP}='true']`).forEach((generated) => {
    let current: HTMLElement | null = null;
    let previousOrigin: string | null | undefined;
    for (const row of Array.from(generated.children)) {
      const value = row.getAttribute(ROW_GROUP_ORIGIN);
      const origin = readRowGroupOrigin(value);
      if (!current || previousOrigin !== value) {
        current = root.ownerDocument.createElement(origin ? origin[1].toLowerCase() : "tbody");
        if (origin) {
          current.setAttribute(ROW_GROUP_ID, origin[0]);
          origin[2].forEach(([name, content]) => {
            // Metadata may come from an imported file; invalid names must not
            // interrupt dirty-state reporting.
            if (/^[a-zA-Z_][a-zA-Z0-9_.:-]*$/.test(name)) current?.setAttribute(name, content);
          });
        }
        // Preserve any real formatting subsequently applied to the overflow group.
        const formerGroupId = generated.getAttribute(ROW_GROUP_ID);
        if (!origin || !formerGroupId || formerGroupId === origin[0]) {
          Array.from(generated.attributes).forEach((attribute) => {
            if (!COMPARISON_ONLY_ATTRIBUTES.has(attribute.name)) {
              current?.setAttribute(attribute.name, attribute.value);
            }
          });
        }
        cleanPresentationState(current);
        generated.before(current);
        previousOrigin = value;
      }
      current.appendChild(row);
    }
    generated.remove();
  });
}

function coalesceRowGroups(table: HTMLTableElement): void {
  const groups = new Map<string, Element>();
  for (const group of Array.from(table.children)) {
    const id = group.getAttribute(ROW_GROUP_ID);
    if (!id || !["TBODY", "TFOOT"].includes(group.tagName)) continue;
    const key = JSON.stringify([id, group.tagName, attributesKey(group)]);
    const first = groups.get(key);
    if (!first) {
      groups.set(key, group);
      continue;
    }
    while (group.firstChild) first.appendChild(group.firstChild);
    group.remove();
  }
}

function sortAttributes(element: Element): void {
  const attributes = Array.from(element.attributes)
    .map((attribute) => [attribute.name, attribute.value])
    .sort(([left], [right]) => left.localeCompare(right));
  Array.from(element.attributes).forEach((attribute) => element.removeAttribute(attribute.name));
  attributes.forEach(([name, value]) => element.setAttribute(name, value));
}
