export const EDITOR_ROOT_CLASS = "pcf-html-editor-root";
export const EDITOR_ROOT_SELECTOR = `.${EDITOR_ROOT_CLASS}`;

export function removeEditorRootFromSelector(selector: string): string {
  const trimmedSelector = selector.trim();
  const scopedPrefix = `${EDITOR_ROOT_SELECTOR} `;
  return trimmedSelector.startsWith(scopedPrefix)
    ? trimmedSelector.slice(scopedPrefix.length)
    : trimmedSelector;
}
