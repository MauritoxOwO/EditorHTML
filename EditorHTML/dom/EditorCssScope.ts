export const EDITOR_ROOT_CLASS = "pcf-html-editor-root";
export const EDITOR_ROOT_SELECTOR = `.${EDITOR_ROOT_CLASS}`;

export function removeEditorRootFromSelector(selector: string): string {
  const trimmedSelector = selector.trim();
  const scopedPrefix = `${EDITOR_ROOT_SELECTOR} `;
  return trimmedSelector.startsWith(scopedPrefix)
    ? trimmedSelector.slice(scopedPrefix.length)
    : trimmedSelector;
}

export function rebaseEditorCssForPdf(css: string): string {
  const editorRootPattern = escapeRegExp(EDITOR_ROOT_SELECTOR);

  return css
    .replace(
      new RegExp(`${editorRootPattern}\\s+\\.hwe-page-inner\\b`, "g"),
      ".hwe-page-inner"
    )
    .replace(new RegExp(`${editorRootPattern}\\s+\\.hwe-page\\b`, "g"), ".hwe-page")
    .replace(new RegExp(`${editorRootPattern}\\s+`, "g"), "");
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
