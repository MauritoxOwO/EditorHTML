export class DocumentSearch {
  private readonly input = document.createElement("input");
  private readonly caseButton = document.createElement("button");
  private readonly replacement = document.createElement("input");
  private readonly counter = document.createElement("span");
  private readonly previous = document.createElement("button");
  private readonly next = document.createElement("button");
  private readonly replace = document.createElement("button");
  private readonly replaceAllButton = document.createElement("button");
  private readonly observer: MutationObserver;
  private matches: Range[] = [];
  private index = -1;
  private caseSensitive = false;

  constructor(
    private readonly workspace: HTMLElement,
    private readonly beforeReplace: () => void,
    private readonly afterReplace: (page: HTMLElement, forceRebalance?: boolean) => void
  ) {
    this.observer = new MutationObserver(() => this.refresh());
  }

  build(): HTMLElement {
    const bar = document.createElement("div");
    bar.className = "hwe-toolbar hwe-search-bar";
    bar.setAttribute("role", "search");
    this.input.type = "search";
    this.input.placeholder = "Buscar en el documento";
    this.input.setAttribute("aria-label", "Buscar en el documento");
    this.input.addEventListener("input", () => this.refresh());
    this.input.addEventListener("keydown", (event) => {
      if (event.key !== "Enter" || event.isComposing) return;
      event.preventDefault();
      this.navigate(event.shiftKey ? -1 : 1);
    });
    const searchInput = document.createElement("span");
    searchInput.className = "hwe-search-input";
    this.caseButton.type = "button";
    this.caseButton.textContent = "Aa";
    this.caseButton.title = "Distinguir mayúsculas y minúsculas";
    this.caseButton.setAttribute("aria-label", this.caseButton.title);
    this.caseButton.setAttribute("aria-pressed", "false");
    this.caseButton.addEventListener("click", () => {
      this.caseSensitive = !this.caseSensitive;
      this.caseButton.setAttribute("aria-pressed", String(this.caseSensitive));
      this.caseButton.classList.toggle("hwe-active", this.caseSensitive);
      this.refresh();
    });
    searchInput.append(this.input, this.caseButton);
    this.counter.setAttribute("role", "status");
    this.counter.setAttribute("aria-live", "polite");
    this.previous.textContent = "Anterior";
    this.next.textContent = "Siguiente";
    for (const [button, direction] of [[this.previous, -1], [this.next, 1]] as const) {
      button.type = "button";
      button.addEventListener("click", () => this.navigate(direction));
    }
    this.replacement.type = "text";
    this.replacement.placeholder = "Reemplazar por";
    this.replacement.setAttribute("aria-label", "Reemplazar por");
    this.replacement.addEventListener("keydown", (event) => {
      if (event.key !== "Enter" || event.isComposing) return;
      event.preventDefault();
      this.replaceCurrent();
    });
    this.replace.type = "button";
    this.replace.textContent = "Reemplazar";
    this.replace.addEventListener("click", () => this.replaceCurrent());
    this.replaceAllButton.type = "button";
    this.replaceAllButton.textContent = "Reemplazar todo";
    this.replaceAllButton.addEventListener("click", () => this.replaceAll());
    bar.append(searchInput, this.previous, this.next, this.counter, this.replacement, this.replace, this.replaceAllButton);
    this.refresh();
    this.observer.observe(this.workspace, { childList: true, characterData: true, subtree: true });
    return bar;
  }

  destroy(): void {
    this.observer.disconnect();
    this.matches = [];
  }

  private refresh(): void {
    this.matches = this.findMatches(this.input.value);
    this.index = -1;
    this.previous.disabled = this.next.disabled = this.matches.length === 0;
    this.replace.disabled = true;
    this.replaceAllButton.disabled = this.matches.length === 0;
    this.counter.textContent = this.input.value ? `${this.matches.length} coincidencias` : "";
  }

  private navigate(direction: number): void {
    if (!this.matches.length) return;
    this.index = this.index < 0
      ? (direction > 0 ? 0 : this.matches.length - 1)
      : (this.index + direction + this.matches.length) % this.matches.length;
    const range = this.matches[this.index];
    const element = range.startContainer.parentElement!;
    element.closest<HTMLElement>(".hwe-page-inner")?.focus({ preventScroll: true });
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);

    // Desplazar solo el documento, usando coordenadas que ya incluyen el zoom.
    const matchRect = range.getBoundingClientRect();
    const viewport = this.workspace.getBoundingClientRect();
    this.workspace.scrollTop += matchRect.top - viewport.top - this.workspace.clientHeight / 2;
    this.workspace.scrollLeft += matchRect.left - viewport.left - this.workspace.clientWidth / 2;
    this.counter.textContent = `${this.index + 1} de ${this.matches.length}`;
    this.replace.disabled = false;
  }

  private replaceCurrent(): void {
    const range = this.matches[this.index];
    const page = range?.startContainer.parentElement?.closest<HTMLElement>(".hwe-page");
    if (!page || !this.workspace.contains(page)) return;
    this.beforeReplace();
    range.deleteContents();
    const inserted = document.createTextNode(this.replacement.value);
    range.insertNode(inserted);
    const selection = window.getSelection();
    const caret = document.createRange();
    caret.setStartAfter(inserted);
    caret.collapse(true);
    selection?.removeAllRanges();
    selection?.addRange(caret);
    this.afterReplace(page);
    this.observer.takeRecords();
    this.refresh();
  }

  private replaceAll(): void {
    // Obtener una instantánea actual: el reemplazo puede contener el texto buscado.
    const matches = this.findMatches(this.input.value);
    const page = matches[0]?.startContainer.parentElement?.closest<HTMLElement>(".hwe-page");
    if (!page || !this.workspace.contains(page)) return;

    this.beforeReplace();
    let firstInserted: Text | undefined;
    // Empezar por el final conserva las posiciones de las coincidencias anteriores.
    for (let index = matches.length - 1; index >= 0; index--) {
      const range = matches[index];
      range.deleteContents();
      firstInserted = document.createTextNode(this.replacement.value);
      range.insertNode(firstInserted);
    }
    if (firstInserted) {
      const caret = document.createRange();
      caret.setStartAfter(firstInserted);
      caret.collapse(true);
      const selection = window.getSelection();
      selection?.removeAllRanges();
      selection?.addRange(caret);
    }
    // Registrar una sola edición y repaginar desde la primera página afectada.
    this.afterReplace(page, true);
    this.observer.takeRecords();
    this.refresh();
  }

  private findMatches(query: string): Range[] {
    if (!query.trim()) return [];
    const parts: { node: Text; start: number; end: number }[] = [];
    let text = "";
    let previousBlock: Element | null = null;
    const walker = document.createTreeWalker(
      this.workspace,
      NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT,
      { acceptNode: (node) => {
        const element = node.nodeType === Node.ELEMENT_NODE ? node as Element : node.parentElement;
        return element?.closest("[contenteditable='false'], [data-hwe-api-header], [data-hwe-dynamic-header], [data-hwe-caret], script, style, [hidden]")
          ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT;
      } }
    );
    while (walker.nextNode()) {
      const node = walker.currentNode;
      if (node.nodeType === Node.ELEMENT_NODE) {
        if ((node as Element).tagName === "BR") text += "\n";
        continue;
      }
      if (!node.parentElement?.closest(".hwe-page-inner")) continue;
      const block = node.parentElement.closest("p, div, li, h1, h2, h3, h4, h5, h6, blockquote, pre, td, th");
      if (block !== previousBlock) text += "\n";
      previousBlock = block;
      const start = text.length;
      text += (node as Text).data;
      parts.push({ node: node as Text, start, end: text.length });
    }

    // Unir los nodos de texto permite encontrar palabras partidas por spans de formato.
    const pattern = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), this.caseSensitive ? "g" : "gi");
    const matches: Range[] = [];
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(text))) {
      const start = parts.find((part) => part.start <= match!.index && part.end > match!.index);
      const endOffset = match.index + match[0].length;
      const end = parts.find((part) => part.start < endOffset && part.end >= endOffset);
      if (!start || !end) continue;
      const range = document.createRange();
      range.setStart(start.node, match.index - start.start);
      range.setEnd(end.node, endOffset - end.start);
      matches.push(range);
    }
    return matches;
  }
}
