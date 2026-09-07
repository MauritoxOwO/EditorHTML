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
