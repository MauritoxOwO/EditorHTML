export const RUNTIME_PAGE_HEADER_ATTR = "data-hwe-runtime-page-header";
const BOCM_LOGO_SELECTOR = ".bocm-logo";

export class RuntimePageHeaderRenderer {
  private logoSrc = "";

  setLogoSrc(src: string): void {
    this.logoSrc = src;
  }

  ensureHeader(page: HTMLElement, index?: number): void {
    const existing = this.getDirectHeader(page);
    if (existing) {
      if (page.firstElementChild !== existing) {
        page.insertBefore(existing, page.firstChild);
      }
      this.updateLogo(existing);

      if (index !== undefined) {
        this.updatePageNumber(existing, index);
      }
      return;
    }

    const header = this.createHeader();
    this.updateLogo(header);
    page.insertBefore(header, page.firstChild);
  }

  private updatePageNumber(header: HTMLElement, index: number): void {
    const pageNumberElement = header.querySelector(".bocm-meta-right");
    if (!pageNumberElement) return;

    pageNumberElement.textContent = `P&aacute;g. ${index}`;
  }

  private createHeader(): HTMLElement {
    const host = document.createElement("div");
    host.className = "hwe-runtime-page-header";
    host.setAttribute(RUNTIME_PAGE_HEADER_ATTR, "true");
    host.setAttribute("contenteditable", "false");
    host.setAttribute("aria-hidden", "true");

    host.innerHTML = `
      <header class="bocm-page-header">
        <div class="bocm-header-top">
          <img class="bocm-logo" alt="BOCM" height="68" />
          <div class="bocm-header-title">BOLET&Iacute;N OFICIAL DE LA COMUNIDAD DE MADRID</div>
        </div>
        <div class="bocm-separator"></div>
        <div class="bocm-header-meta">
          <span class="bocm-meta-left">B.O.C.M. N&uacute;m. X</span>
          <span class="bocm-meta-center">XXXXXX X DE XXXXX DE 2026</span>
          <span class="bocm-meta-right">P&aacute;g. X</span>
        </div>
      </header>
    `;

    return host;
  }

  private updateLogo(header: HTMLElement): void {
    const logo = header.querySelector<HTMLImageElement>(BOCM_LOGO_SELECTOR);
    if (!logo) return;

    if (this.logoSrc) {
      if (logo.src !== this.logoSrc) logo.src = this.logoSrc;
      return;
    }

    logo.removeAttribute("src");
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
