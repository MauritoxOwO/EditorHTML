// These identities describe clones made by pagination, not user formatting.
export const CONTAINER_FLOW_ID = "data-hwe-container-flow-id";
export const INLINE_FLOW_ID = "data-hwe-inline-flow-id";
export const AUTO_IMAGE_MAX_HEIGHT = "data-hwe-auto-max-height";
export const ROW_GROUP_ID = "data-hwe-row-group-id";
export const ROW_GROUP_ORIGIN = "data-hwe-row-group-origin";
export const GENERATED_ROW_GROUP = "data-hwe-generated-row-group";

let flowCounter = 0;

export function ensureFlowIdentity(element: HTMLElement, attribute: string): void {
  if (element.hasAttribute(attribute)) return;
  element.setAttribute(attribute, `hwe-flow-${Date.now().toString(36)}-${flowCounter++}`);
}

// Enter clona atributos del parrafo. La parte posterior es un parrafo nuevo,
// aunque sus continuaciones ya estuvieran repartidas en otras paginas.
export function separateParagraphFlow(root: HTMLElement): void {
  const node = window.getSelection()?.anchorNode;
  const element = node?.nodeType === Node.ELEMENT_NODE ? node as HTMLElement : node?.parentElement;
  const block = element?.closest<HTMLElement>("p, div, li, h1, h2, h3, h4, h5, h6, blockquote, pre");
  if (!block || !root.contains(block) || block.hasAttribute("contenteditable")) return;

  for (const attribute of ["data-hwe-text-flow-id", CONTAINER_FLOW_ID, INLINE_FLOW_ID]) {
    const identities = new Map<string, string>();
    [block, ...Array.from(block.querySelectorAll<HTMLElement>(`[${attribute}]`))].forEach((part) => {
      const id = part.getAttribute(attribute);
      if (id) identities.set(id, `hwe-flow-${Date.now().toString(36)}-${flowCounter++}`);
    });
    if (!identities.size) continue;
    root.querySelectorAll<HTMLElement>(`[${attribute}]`).forEach((part) => {
      const id = identities.get(part.getAttribute(attribute) ?? "");
      if (id && (block === part || block.contains(part) ||
          (block.compareDocumentPosition(part) & Node.DOCUMENT_POSITION_FOLLOWING))) {
        part.setAttribute(attribute, id);
      }
    });
  }
}
