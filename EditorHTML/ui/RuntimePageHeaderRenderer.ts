export const RUNTIME_PAGE_HEADER_ATTR = "data-hwe-runtime-page-header";

export class RuntimePageHeaderRenderer {
  ensureHeader(page: HTMLElement): void {
    const existing = this.getDirectHeader(page);
    if (existing) {
      if (page.firstElementChild !== existing) {
        page.insertBefore(existing, page.firstChild);
      }
      return;
    }

    page.insertBefore(this.createHeader(), page.firstChild);
  }

  private createHeader(): HTMLElement {
    const host = document.createElement("div");
    host.className = "hwe-runtime-page-header";
    host.setAttribute(RUNTIME_PAGE_HEADER_ATTR, "true");
    host.setAttribute("contenteditable", "false");
    host.setAttribute("aria-hidden", "true");

    host.innerHTML = `
      <header class="bocm-header">
        <div class="bocm-logo">
          <div class="madrid-flag">
            <div class="stars">&#10022;&#10022;&#10022;&#10022;&#10022;</div>
            <div class="stars">&#10022;&#10022;&#10022;&#10022;&#10022;</div>
          </div>
          <div class="bocm-title">
            <h1>BOLET&Iacute;N OFICIAL</h1>
            <h2>DE LA COMUNIDAD DE MADRID</h2>
          </div>
        </div>
        <div class="bocm-meta-bar"></div>
        <div class="bocm-info-row">
          <span>B.O.C.M. N&uacute;m. X</span>
          <span>XXXXXX X DE XXXXX DE 2026</span>
          <span>P&aacute;g. 1</span>
        </div>
      </header>
    `;

    return host;
  }

  private getDirectHeader(page: HTMLElement): HTMLElement | null {
    const first = page.firstElementChild as HTMLElement | null;
    if (first?.getAttribute(RUNTIME_PAGE_HEADER_ATTR) === "true") return first;

    return (
      Array.from(page.children).find(
        (child): child is HTMLElement =>
          child instanceof HTMLElement &&
          child.getAttribute(RUNTIME_PAGE_HEADER_ATTR) === "true"
      ) ?? null
    );
  }
}
