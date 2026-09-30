const PARAGRAPH_SELECTOR = "p, div, li, h1, h2, h3, h4, h5, h6, blockquote, pre";
const BLOCK_SELECTOR = `${PARAGRAPH_SELECTOR}, ul, ol, section, article, figure, figcaption, hr, table`;
const EXCLUDED_SELECTOR = "table, [contenteditable='false'], [data-hwe-api-header], [data-hwe-dynamic-header], [hidden], script, style";

// Normalizar solo el contenido de la celda: nunca convertir la celda ni una tabla
// anidada en un párrafo, ni aplicar clases a contenedores de otros párrafos.
export function getCellParagraphs(cell: HTMLTableCellElement): HTMLElement[] {
  const paragraphs: HTMLElement[] = [];

  const visit = (container: HTMLElement): void => {
    const nodes = Array.from(container.childNodes);
    const hasBlockChildren = nodes.some((node) =>
      node instanceof HTMLElement && node.matches(BLOCK_SELECTOR)
    );
    if (container !== cell && container.matches(PARAGRAPH_SELECTOR) && !hasBlockChildren) {
      paragraphs.push(container);
      return;
    }

    let inlineNodes: ChildNode[] = [];
    const flushInlineNodes = (): void => {
      if (inlineNodes.some((node) => node.nodeType === Node.ELEMENT_NODE || node.textContent?.trim())) {
        const paragraph = document.createElement("p");
        container.insertBefore(paragraph, inlineNodes[0]);
        paragraph.append(...inlineNodes);
        paragraphs.push(paragraph);
      }
      inlineNodes = [];
    };

    nodes.forEach((node) => {
      if (node instanceof HTMLElement && node.matches(EXCLUDED_SELECTOR)) {
        flushInlineNodes();
      } else if (node instanceof HTMLElement && node.matches(BLOCK_SELECTOR)) {
        flushInlineNodes();
        if (node.tagName !== "HR") visit(node);
      } else {
        inlineNodes.push(node);
      }
    });
    flushInlineNodes();

    if (container === cell && !cell.children.length && !cell.textContent?.trim()) {
      const paragraph = document.createElement("p");
      paragraph.appendChild(document.createElement("br"));
      cell.replaceChildren(paragraph);
      paragraphs.push(paragraph);
    }
  };

  visit(cell);
  return paragraphs;
}
