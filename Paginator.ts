/**
 * Paginator.ts
 *
 * Gestiona la paginación automática del editor:
 *  - Vigila el scrollHeight de cada página con ResizeObserver.
 *  - Si una página desborda (scrollHeight > PAGE_MAX_HEIGHT_PX),
 *    mueve el último nodo hijo al inicio de la página siguiente.
 *  - Si una página queda por debajo del umbral de fusión y existe
 *    la siguiente, mueve el primer nodo de la siguiente a esta.
 *  - Crea y elimina páginas según sea necesario.
 *  - Notifica al caller mediante callbacks para que pueda
 *    reconstruir los separadores visuales.
 */

// A4 a 96 dpi: 297mm → 297 * (96/25.4) ≈ 1122 px
// Márgenes 2 × 25.4mm → 2 × 96.38 ≈ 193 px
// Área de contenido ≈ 1122 − 193 = 929 px
const PAGE_MAX_HEIGHT_PX = 929;

// Si una página tiene menos de este contenido Y existe la siguiente,
// intentamos fusionar (traer nodos de la siguiente).
const PAGE_MERGE_THRESHOLD_PX = PAGE_MAX_HEIGHT_PX * 0.85; // 85 %

export type PageFactory = (html?: string) => HTMLElement;
export type OnPagesChanged = (pages: HTMLElement[]) => void;

export class Paginator {
  private pages: HTMLElement[] = [];
  private observer!: ResizeObserver;
  private pageFactory: PageFactory;
  private onPagesChanged: OnPagesChanged;

  /** Evita reentradas mientras rebalanceamos */
  private rebalancing = false;

  constructor(pageFactory: PageFactory, onPagesChanged: OnPagesChanged) {
    this.pageFactory = pageFactory;
    this.onPagesChanged = onPagesChanged;

    this.observer = new ResizeObserver((entries) => {
      if (this.rebalancing) return;
      for (const entry of entries) {
        const page = entry.target as HTMLElement;
        this.rebalancePage(page);
      }
    });
  }

  // ── API pública ───────────────────────────────────────────────

  /** Registra las páginas iniciales (ya renderizadas en el DOM). */
  setPages(pages: HTMLElement[]): void {
    // Desconectar observadores previos
    this.pages.forEach((p) => this.observer.unobserve(p));
    this.pages = [...pages];
    this.pages.forEach((p) => this.observer.observe(p));
  }

  /** Retorna la lista actual de páginas (puede haber cambiado). */
  getPages(): HTMLElement[] {
    return this.pages;
  }

  /** Libera el ResizeObserver. Llamar en destroy(). */
  destroy(): void {
    this.observer.disconnect();
  }

  // ── Lógica de rebalanceo ──────────────────────────────────────

  private rebalancePage(page: HTMLElement): void {
    const index = this.pages.indexOf(page);
    if (index === -1) return;

    this.rebalancing = true;
    try {
      // Primero resolvemos overflow (puede crear nuevas páginas)
      this.resolveOverflow(index);
      // Luego intentamos fusión desde atrás
      this.resolveMerge(index);
    } finally {
      this.rebalancing = false;
    }
  }

  /**
   * Si la página desborda, mueve el último hijo al inicio
   * de la página siguiente. Repite hasta que no desborde.
   */
  private resolveOverflow(index: number): void {
    let safetyCounter = 0;
    const MAX_ITERATIONS = 200;

    while (safetyCounter++ < MAX_ITERATIONS) {
      const page = this.pages[index];
      if (!page) break;

      const contentHeight = this.getContentHeight(page);
      if (contentHeight <= PAGE_MAX_HEIGHT_PX) break;

      const lastChild = this.getLastMeaningfulChild(page);
      if (!lastChild) break; // no hay nodos movibles

      // Obtener o crear la página siguiente
      const nextPage = this.getOrCreateNextPage(index);

      // Mover el nodo al inicio de la siguiente página
      nextPage.insertBefore(lastChild, nextPage.firstChild);

      // Si la siguiente página también desborda, continuamos
      // con ella en la próxima iteración del ResizeObserver.
      // Aquí solo resolvemos la página actual.
    }
  }

  /**
   * Si la página tiene poco contenido y existe la siguiente,
   * trae nodos de la siguiente hasta llenarla o vaciarla.
   */
  private resolveMerge(index: number): void {
    let safetyCounter = 0;
    const MAX_ITERATIONS = 200;

    while (safetyCounter++ < MAX_ITERATIONS) {
      const page     = this.pages[index];
      const nextPage = this.pages[index + 1];
      if (!page || !nextPage) break;

      const contentHeight = this.getContentHeight(page);
      if (contentHeight >= PAGE_MERGE_THRESHOLD_PX) break;

      const firstChild = this.getFirstMeaningfulChild(nextPage);
      if (!firstChild) {
        // Página siguiente vacía: eliminarla
        this.removePage(index + 1);
        break;
      }

      // Mover el primer hijo de la siguiente a esta página
      page.appendChild(firstChild);

      // Comprobar si hemos desbordado al hacer la fusión
      if (this.getContentHeight(page) > PAGE_MAX_HEIGHT_PX) {
        // Devolver el nodo que acabamos de mover
        nextPage.insertBefore(firstChild, nextPage.firstChild);
        break;
      }

      // Si la siguiente quedó vacía, eliminarla
      if (!this.getFirstMeaningfulChild(nextPage)) {
        this.removePage(index + 1);
        break;
      }
    }
  }

  // ── Gestión de páginas ────────────────────────────────────────

  private getOrCreateNextPage(afterIndex: number): HTMLElement {
    if (this.pages[afterIndex + 1]) {
      return this.pages[afterIndex + 1];
    }

    const newPage = this.pageFactory();
    this.pages.splice(afterIndex + 1, 0, newPage);
    this.observer.observe(newPage);
    this.onPagesChanged(this.pages);
    return newPage;
  }

  private removePage(index: number): void {
    const page = this.pages[index];
    if (!page) return;

    this.observer.unobserve(page);
    this.pages.splice(index, 1);
    this.onPagesChanged(this.pages);
  }

  // ── Utilidades DOM ────────────────────────────────────────────

  /**
   * Altura real del contenido, independiente del min-height CSS.
   * Usamos scrollHeight para capturar contenido que desborda.
   */
  private getContentHeight(page: HTMLElement): number {
    // scrollHeight incluye padding pero no el min-height de CSS,
    // por lo que refleja el contenido real.
    return page.scrollHeight;
  }

  /**
   * Último hijo que no sea solo un <br> o nodo de texto vacío.
   * Si solo hay un hijo, no lo movemos para evitar páginas vacías.
   */
  private getLastMeaningfulChild(page: HTMLElement): ChildNode | null {
    const children = Array.from(page.childNodes).filter(
      (n) => !this.isEmptyNode(n)
    );
    if (children.length <= 1) return null; // mantener al menos uno
    return children[children.length - 1];
  }

  private getFirstMeaningfulChild(page: HTMLElement): ChildNode | null {
    const children = Array.from(page.childNodes).filter(
      (n) => !this.isEmptyNode(n)
    );
    return children[0] ?? null;
  }

  private isEmptyNode(node: ChildNode): boolean {
    if (node.nodeType === Node.TEXT_NODE) {
      return (node.textContent ?? "").trim() === "";
    }
    if (node.nodeType === Node.ELEMENT_NODE) {
      const el = node as HTMLElement;
      if (el.tagName === "BR") return true;
      // Párrafo vacío: <p><br></p> o <p></p>
      if (
        el.tagName === "P" &&
        el.childNodes.length <= 1 &&
        (el.textContent ?? "").trim() === ""
      ) {
        return true;
      }
    }
    return false;
  }
}
