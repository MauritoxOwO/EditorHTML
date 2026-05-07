export interface PageSetup {
  width: string;
  height: string;
  marginTop: string;
  marginRight: string;
  marginBottom: string;
  marginLeft: string;
}

export const DEFAULT_PAGE_SETUP: PageSetup = {
  width: "210mm",
  height: "297mm",
  marginTop: "25.4mm",
  marginRight: "25.4mm",
  marginBottom: "25.4mm",
  marginLeft: "25.4mm",
};

const PAGE_SETUP_ATTRS: Record<keyof PageSetup, string> = {
  width: "data-hwe-page-width",
  height: "data-hwe-page-height",
  marginTop: "data-hwe-margin-top",
  marginRight: "data-hwe-margin-right",
  marginBottom: "data-hwe-margin-bottom",
  marginLeft: "data-hwe-margin-left",
};

const PAGE_SETUP_VARS: Record<keyof PageSetup, string> = {
  width: "--hwe-page-width",
  height: "--hwe-page-height",
  marginTop: "--hwe-margin-top",
  marginRight: "--hwe-margin-right",
  marginBottom: "--hwe-margin-bottom",
  marginLeft: "--hwe-margin-left",
};

const CSS_LENGTH_RE = /^-?(?:\d+|\d*\.\d+)(?:px|pt|pc|mm|cm|in)$/i;

export function normalizePageSetup(setup?: Partial<PageSetup> | null): PageSetup {
  return {
    width: sanitizeCssLength(setup?.width) ?? DEFAULT_PAGE_SETUP.width,
    height: sanitizeCssLength(setup?.height) ?? DEFAULT_PAGE_SETUP.height,
    marginTop: sanitizeCssLength(setup?.marginTop) ?? DEFAULT_PAGE_SETUP.marginTop,
    marginRight: sanitizeCssLength(setup?.marginRight) ?? DEFAULT_PAGE_SETUP.marginRight,
    marginBottom: sanitizeCssLength(setup?.marginBottom) ?? DEFAULT_PAGE_SETUP.marginBottom,
    marginLeft: sanitizeCssLength(setup?.marginLeft) ?? DEFAULT_PAGE_SETUP.marginLeft,
  };
}

export function applyPageSetup(target: HTMLElement, setup: PageSetup): void {
  const normalized = normalizePageSetup(setup);

  (Object.keys(PAGE_SETUP_VARS) as Array<keyof PageSetup>).forEach((key) => {
    target.style.setProperty(PAGE_SETUP_VARS[key], normalized[key]);
  });
}

export function readPageSetupFromElement(element: Element | null): PageSetup | null {
  if (!element) return null;

  const setup = (Object.keys(PAGE_SETUP_ATTRS) as Array<keyof PageSetup>).reduce<
    Partial<PageSetup>
  >((acc, key) => {
    const value = element.getAttribute(PAGE_SETUP_ATTRS[key]);
    if (value) acc[key] = value;
    return acc;
  }, {});

  return Object.keys(setup).length > 0 ? normalizePageSetup(setup) : null;
}

export function writePageSetupAttributes(element: HTMLElement, setup: PageSetup): void {
  const normalized = normalizePageSetup(setup);

  element.setAttribute("data-hwe-document", "true");
  (Object.keys(PAGE_SETUP_ATTRS) as Array<keyof PageSetup>).forEach((key) => {
    element.setAttribute(PAGE_SETUP_ATTRS[key], normalized[key]);
  });
}

export function makePageSetupWrapper(html: string, setup: PageSetup): string {
  const wrapper = document.createElement("div");
  writePageSetupAttributes(wrapper, setup);
  wrapper.innerHTML = html;
  return wrapper.outerHTML;
}

export function sanitizeCssLength(value: string | null | undefined): string | null {
  if (!value) return null;

  const normalized = value.trim().replace(/^(-?\d*\.\d+)0+(?=[a-z])/i, "$1");
  if (normalized === "0") return normalized;
  if (!CSS_LENGTH_RE.test(normalized)) return null;

  return normalized;
}