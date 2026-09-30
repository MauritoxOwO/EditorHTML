interface ContentPoint {
  offset: number;
  preferNext: boolean;
  header?: { flowId: string; index: number };
}

export interface HistoryViewState {
  anchor?: ContentPoint;
  focus?: ContentPoint;
  caretTop?: number;
  pageIndex: number;
  scrollTop: number;
  scrollLeft: number;
}

interface ContentUnit { node: Node; length: number }
interface DomPoint { node: Node; offset: number }

const EXCLUDED = "[contenteditable='false'], [hidden], [data-hwe-api-header], [data-hwe-dynamic-header], [data-hwe-runtime-page-header], style, script";

// Las posiciones cuentan contenido, no nodos: paginar puede dividir textos y clonar contenedores.
function contentUnits(root: HTMLElement): ContentUnit[] {
  const seenHeaders = new Set<string>();
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT, {
    acceptNode: (node) => {
      if (node instanceof HTMLElement) {
        if (node.matches(EXCLUDED)) return NodeFilter.FILTER_REJECT;
        if (node.tagName === "THEAD") {
          const table = node.closest("table[data-hwe-repeat-header='true']");
          const id = table?.getAttribute("data-hwe-table-flow-id");
          if (id) {
            if (seenHeaders.has(id)) return NodeFilter.FILTER_REJECT;
            seenHeaders.add(id);
          }
        }
        return node.matches("br, img, hr, p:empty, div:empty, td:empty, th:empty, li:empty")
          ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP;
      }
      const parent = node.parentElement;
      if (!node.textContent || !parent?.closest(".hwe-page-inner")) return NodeFilter.FILTER_REJECT;
      if (!node.textContent.trim() && parent.matches(
        ".hwe-page-inner, [data-hwe-page-break], table, thead, tbody, tfoot, tr, ul, ol"
      )) return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT;
    },
  });
  const units: ContentUnit[] = [];
  let node: Node | null;
  while ((node = walker.nextNode())) {
    units.push({ node, length: node.nodeType === Node.TEXT_NODE ? (node as Text).length : 1 });
  }
  return units;
}

function headers(workspace: HTMLElement, flowId: string): HTMLElement[] {
  return Array.from(workspace.querySelectorAll<HTMLElement>("thead")).filter(header =>
    header.closest("table")?.getAttribute("data-hwe-table-flow-id") === flowId
  );
}

function capturePoint(workspace: HTMLElement, node: Node, offset: number): ContentPoint {
  const element = node instanceof HTMLElement ? node : node.parentElement;
  const header = element?.closest<HTMLElement>("thead");
  const flowId = header?.closest("table[data-hwe-repeat-header='true']")?.getAttribute("data-hwe-table-flow-id");
  const scope = flowId && header ? header : workspace;
  const range = document.createRange();
  range.setStart(node, offset);
  range.collapse(true);
  let position = 0;
  for (const unit of contentUnits(scope)) {
    if (unit.node === node) {
      position += node.nodeType === Node.TEXT_NODE ? offset : 0;
      break;
    }
    if (range.comparePoint(unit.node, 0) >= 0) break;
    position += unit.length;
  }
  return {
    offset: position,
    preferNext: offset === 0,
    header: flowId && header ? { flowId, index: headers(workspace, flowId).indexOf(header) } : undefined,
  };
}

function resolvePoint(workspace: HTMLElement, point: ContentPoint): DomPoint | null {
  const candidates = point.header ? headers(workspace, point.header.flowId) : [];
  const scope = point.header ? candidates[Math.min(point.header.index, candidates.length - 1)] : workspace;
  if (!scope) return null;
  const units = contentUnits(scope);
  let remaining = point.offset;
  for (let index = 0; index < units.length; index++) {
    const { node, length } = units[index];
    if (remaining < length || (remaining === length && !point.preferNext) || index === units.length - 1) {
      if (node.nodeType === Node.TEXT_NODE) return { node, offset: Math.min(remaining, length) };
      if (!(node as Element).matches("br, img, hr")) return { node, offset: 0 };
      const parent = node.parentNode;
      if (!parent) return null;
      return { node: parent, offset: Array.from(parent.childNodes).indexOf(node as ChildNode) + (remaining > 0 ? 1 : 0) };
    }
    remaining -= length;
  }
  return null;
}

function caretTop(node: Node, offset: number): number | undefined {
  const range = document.createRange();
  range.setStart(node, offset);
  range.collapse(true);
  const rect = range.getClientRects()[0];
  if (rect?.height) return rect.top;
  const element = node instanceof HTMLElement ? node : node.parentElement;
  return element?.getBoundingClientRect().top;
}

export function captureHistoryView(workspace: HTMLElement): HistoryViewState {
  const selection = window.getSelection();
  const focusElement = selection?.focusNode instanceof HTMLElement ? selection.focusNode : selection?.focusNode?.parentElement;
  const page = focusElement?.closest<HTMLElement>(".hwe-page");
  const state: HistoryViewState = {
    pageIndex: Math.max(0, Array.from(workspace.querySelectorAll(".hwe-page")).indexOf(page as HTMLElement)),
    scrollTop: workspace.scrollTop,
    scrollLeft: workspace.scrollLeft,
  };
  if (selection?.anchorNode && selection.focusNode && workspace.contains(selection.anchorNode) &&
      workspace.contains(selection.focusNode) && focusElement?.closest(".hwe-page-inner") &&
      !focusElement.closest(EXCLUDED)) {
    state.anchor = capturePoint(workspace, selection.anchorNode, selection.anchorOffset);
    state.focus = capturePoint(workspace, selection.focusNode, selection.focusOffset);
    const top = caretTop(selection.focusNode, selection.focusOffset);
    if (top !== undefined) state.caretTop = top - workspace.getBoundingClientRect().top;
  }
  return state;
}

export function restoreHistoryView(workspace: HTMLElement, state: HistoryViewState): void {
  const anchor = state.anchor ? resolvePoint(workspace, state.anchor) : null;
  const focus = state.focus ? resolvePoint(workspace, state.focus) : null;
  if (anchor && focus) {
    const element = focus.node instanceof HTMLElement ? focus.node : focus.node.parentElement;
    element?.closest<HTMLElement>(".hwe-page-inner")?.focus({ preventScroll: true });
    window.getSelection()?.setBaseAndExtent(anchor.node, anchor.offset, focus.node, focus.offset);
  } else {
    const editables = workspace.querySelectorAll<HTMLElement>(".hwe-page-inner");
    const editable = editables[Math.min(state.pageIndex, editables.length - 1)];
    if (editable) {
      editable.focus({ preventScroll: true });
      window.getSelection()?.collapse(editable, 0);
    }
  }
  workspace.scrollTop = state.scrollTop;
  workspace.scrollLeft = state.scrollLeft;
  if (focus && state.caretTop !== undefined) {
    const top = caretTop(focus.node, focus.offset);
    if (top !== undefined) workspace.scrollTop += top - workspace.getBoundingClientRect().top - state.caretTop;
  }
}
