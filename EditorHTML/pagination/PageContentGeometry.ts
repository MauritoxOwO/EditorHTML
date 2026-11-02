export interface PageContentBounds {
  top: number;
  bottom: number;
  /** Content height in document CSS pixels, before the view zoom is applied. */
  contentHeight: number;
  scale: number;
}

/** Viewport bounds exclude the complete page padding, including at non-default zoom. */
export function getPageContentBounds(inner: HTMLElement): PageContentBounds {
  const styles = getComputedStyle(inner);
  const rect = inner.getBoundingClientRect();
  const paddingTop = parseFloat(styles.paddingTop) || 0;
  const paddingBottom = parseFloat(styles.paddingBottom) || 0;
  const borderTop = parseFloat(styles.borderTopWidth) || 0;
  const borderBottom = parseFloat(styles.borderBottomWidth) || 0;
  const verticalInsets = paddingTop + paddingBottom + borderTop + borderBottom;
  const cssHeight = parseFloat(styles.height);
  // Computed height preserves fractional CSS pixels; offsetHeight rounds them.
  const layoutHeight = Number.isFinite(cssHeight) && cssHeight > 0
    ? cssHeight + (styles.boxSizing === "border-box" ? 0 : verticalInsets)
    : inner.offsetHeight;
  const scale = layoutHeight > 0 && rect.height > 0 ? rect.height / layoutHeight : 1;

  return {
    top: rect.top + (borderTop + paddingTop) * scale,
    bottom: rect.bottom - (borderBottom + paddingBottom) * scale,
    contentHeight: Math.max(0, layoutHeight > 0
      ? layoutHeight - verticalInsets
      : inner.clientHeight - paddingTop - paddingBottom),
    scale,
  };
}
