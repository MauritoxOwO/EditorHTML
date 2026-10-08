import { captureHistoryView, restoreHistoryView } from "./HistoryViewState";

export type TextCase = "uppercase" | "lowercase";
type InlineFormat = "bold" | "italic" | "underline" | "fontSize" | "fontFamily" | "foreColor" | "hiliteColor";
interface SelectionServices {
  readSelection(): { range: Range; collapsed: boolean } | null;
  getSelectedTextTargets(range: Range): Array<{ node: Text; start: number; end: number }>;
  getSelectedParagraphs(range: Range): HTMLElement[];
}

const BLOCKS = "p, div, li, h1, h2, h3, h4, h5, h6, blockquote, pre";

/** Applies formatting to the current content selection. The editor owns history/bookmarks. */
export class TextSelectionFormatter {
  constructor(
    private readonly rootProvider: () => HTMLElement,
    private readonly documentSelection: SelectionServices
  ) {}

  applyCommand(command: string, value?: string): HTMLElement[] {
    const selection = this.readSelection();
    if (!selection) return [];
    const format = this.commandFormat(command);
    if (!format) return this.applyParagraphCommand(command, selection.range, selection.root);
    if (selection.range.collapsed) return [];
    const targets = this.documentSelection.getSelectedTextTargets(selection.range);
    if (!targets.length) return [];
    const bookmarkRoot = selection.root.querySelector<HTMLElement>(".hwe-workspace") ?? selection.root;
    const bookmark = captureHistoryView(bookmarkRoot);
    const toggle = format === "bold" || format === "italic" || format === "underline";
    const remove = toggle && targets.every((target) => this.isFormatActive(target, format));
    const affected = new Set<HTMLElement>();
    // Offsets are collected before mutation. Reverse order keeps later offsets valid.
    for (const target of [...targets].reverse()) {
      const span = this.wrapTarget(target, (element) => {
        this.setInlineFormat(element, format, remove ? "" : value ?? "", remove);
      }, remove && format === "underline");
      if (span) affected.add(span);
    }
    if (affected.size) restoreHistoryView(bookmarkRoot, bookmark);
    return [...affected];
  }

  applyFontSize(size: string): HTMLElement[] {
    return this.applyCommand("fontSize", size);
  }

  transformSelectedTextCase(textCase: TextCase): HTMLElement[] {
    const selection = this.readSelection();
    if (!selection || selection.range.collapsed) return [];
    const targets = this.documentSelection.getSelectedTextTargets(selection.range);
    const changed: HTMLElement[] = [];
    const nativeSelection = window.getSelection()!;
    const anchor = { node: nativeSelection.anchorNode!, offset: nativeSelection.anchorOffset };
    const focus = { node: nativeSelection.focusNode!, offset: nativeSelection.focusOffset };
    const mappings = new Map<Text, { start: number; end: number; replacement: string }>();
    for (const target of [...targets].reverse()) {
      const parent = target.node.parentElement;
      if (!parent) continue;
      const text = target.node.data.slice(target.start, target.end);
      const mapped = textCase === "uppercase" ? text.toLocaleUpperCase("es-ES") : text.toLocaleLowerCase("es-ES");
      if (mapped === text) continue;
      target.node.replaceData(target.start, target.end - target.start, mapped);
      mappings.set(target.node, { start: target.start, end: target.end, replacement: mapped });
      changed.push(parent);
    }
    const remap = (point: { node: Node; offset: number }): number => {
      if (point.node.nodeType !== Node.TEXT_NODE) return point.offset;
      const mapping = mappings.get(point.node as Text);
      if (!mapping) return point.offset;
      if (point.offset <= mapping.start) return point.offset;
      if (point.offset >= mapping.end) return point.offset + mapping.replacement.length - (mapping.end - mapping.start);
      const prefix = (point.node as Text).data.slice(mapping.start, Math.min(point.offset, mapping.start + mapping.replacement.length));
      return mapping.start + prefix.length;
    };
    if (mappings.size) nativeSelection.setBaseAndExtent(anchor.node, remap(anchor), focus.node, remap(focus));
    return changed;
  }

  private readSelection(): { range: Range; root: HTMLElement } | null {
    const shared = this.documentSelection?.readSelection();
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0 || !selection.anchorNode || !selection.focusNode) return null;
    const root = this.rootProvider();
    const range = shared?.range;
    if (!shared || !range) return null;
    if (!root || !root.contains(range.startContainer) || !root.contains(range.endContainer)) return null;
    return { range, root };
  }

  private wrapTarget(target: TextTarget, configure: (span: HTMLElement) => void, removeUnderline = false): HTMLElement | null {
    const { node, start, end } = target;
    let parent = node.parentNode;
    if (!parent) return null;
    if (end < node.length) node.splitText(end);
    const selected = start > 0 ? node.splitText(start) : node;
    if (removeUnderline) {
      this.liftOutOfUnderline(selected);
      parent = selected.parentNode;
    }
    if (!parent) return null;
    const span = document.createElement("span");
    configure(span);
    parent.insertBefore(span, selected);
    span.appendChild(selected);
    return span;
  }

  private liftOutOfUnderline(selected: Node): void {
    const startElement = selected.parentElement;
    if (!startElement) return;
    const inner = startElement.closest(".hwe-page-inner");
    for (let attempt = 0; attempt < 8; attempt++) {
      const currentElement = selected.parentElement;
      const underline = currentElement ? this.findUnderlineSource(currentElement, inner) : null;
      if (!underline) break;
      this.liftOutOfAncestor(selected, underline);
    }
  }

  private findUnderlineSource(start: HTMLElement, inner: Element | null): HTMLElement | null {
    for (let current: HTMLElement | null = start; current && current !== inner; current = current.parentElement) {
      if (current.tagName === "U" || current.tagName === "INS") return current;
      const ownDecoration = `${current.style.textDecoration} ${current.style.textDecorationLine}`.toLowerCase();
      if (ownDecoration.includes("underline")) return current;
      const computedStyle = window.getComputedStyle(current);
      const computed = `${computedStyle.textDecorationLine} ${computedStyle.textDecoration}`.toLowerCase();
      const parentComputed = current.parentElement
        ? `${window.getComputedStyle(current.parentElement).textDecorationLine} ${window.getComputedStyle(current.parentElement).textDecoration}`.toLowerCase() : "none";
      if (computed.includes("underline") && !parentComputed.includes("underline")) return current;
    }
    return null;
  }

  private liftOutOfAncestor(selected: Node, underline: HTMLElement): void {
    if (underline.matches("p, div, li, h1, h2, h3, h4, h5, h6, blockquote, pre")) {
      this.splitBlockDecoration(selected, underline);
      return;
    }
    let branch = selected;
    let parent = branch.parentNode;
    while (parent && parent !== underline) {
      const grandparent = parent.parentNode;
      if (!grandparent) return;
      const before = parent.cloneNode(false);
      const after = parent.cloneNode(false);
      const selectedWrapper = parent.cloneNode(false);
      while (branch.previousSibling) before.insertBefore(branch.previousSibling, before.firstChild);
      while (branch.nextSibling) after.appendChild(branch.nextSibling);
      selectedWrapper.appendChild(branch);
      if (before.hasChildNodes()) grandparent.insertBefore(before, parent);
      grandparent.insertBefore(selectedWrapper, parent);
      if (after.hasChildNodes()) grandparent.insertBefore(after, parent);
      grandparent.removeChild(parent);
      branch = selectedWrapper;
      parent = selectedWrapper.parentNode;
    }
    if (!parent || parent !== underline) return;
    const grandparent = underline.parentNode;
    if (!grandparent) return;
    const before = underline.cloneNode(false);
    const after = underline.cloneNode(false);
    const selectedWrapper = underline.tagName === "U" || underline.tagName === "INS"
      ? document.createElement("span") : underline.cloneNode(false) as HTMLElement;
    Array.from(underline.attributes).forEach((attribute) => selectedWrapper.setAttribute(attribute.name, attribute.value));
    selectedWrapper.style.textDecoration = "none";
    while (branch.previousSibling) before.insertBefore(branch.previousSibling, before.firstChild);
    while (branch.nextSibling) after.appendChild(branch.nextSibling);
    selectedWrapper.appendChild(branch);
    if (before.hasChildNodes()) grandparent.insertBefore(before, underline);
    grandparent.insertBefore(selectedWrapper, underline);
    if (after.hasChildNodes()) grandparent.insertBefore(after, underline);
    grandparent.removeChild(underline);
  }

  private splitBlockDecoration(selected: Node, block: HTMLElement): void {
    let branch = selected;
    let parent = branch.parentNode;
    while (parent && parent !== block) {
      const grandparent = parent.parentNode;
      if (!grandparent) return;
      const before = parent.cloneNode(false);
      const after = parent.cloneNode(false);
      const selectedWrapper = parent.cloneNode(false);
      while (branch.previousSibling) before.insertBefore(branch.previousSibling, before.firstChild);
      while (branch.nextSibling) after.appendChild(branch.nextSibling);
      selectedWrapper.appendChild(branch);
      if (before.hasChildNodes()) grandparent.insertBefore(before, parent);
      grandparent.insertBefore(selectedWrapper, parent);
      if (after.hasChildNodes()) grandparent.insertBefore(after, parent);
      grandparent.removeChild(parent);
      branch = selectedWrapper;
      parent = selectedWrapper.parentNode;
    }
    if (parent !== block || !block.parentNode) return;
    const computed = window.getComputedStyle(block);
    const decoration = computed.textDecoration || computed.textDecorationLine || "underline";
    const decorationLines = (computed.textDecorationLine || decoration)
      .split(/\s+/).filter((line) => line && line.toLowerCase() !== "underline");
    block.style.textDecorationLine = decorationLines.length ? decorationLines.join(" ") : "none";
    const wrapSiblings = (side: "before" | "after"): void => {
      const wrapper = document.createElement("span");
      wrapper.style.textDecoration = decoration;
      if (side === "before") {
        while (branch.previousSibling) wrapper.insertBefore(branch.previousSibling, wrapper.firstChild);
        if (wrapper.hasChildNodes()) block.insertBefore(wrapper, branch);
      } else {
        while (branch.nextSibling) wrapper.appendChild(branch.nextSibling);
        if (wrapper.hasChildNodes()) block.insertBefore(wrapper, branch.nextSibling);
      }
    };
    wrapSiblings("before");
    wrapSiblings("after");
  }

  private commandFormat(command: string): InlineFormat | null {
    const normalized = command.toLowerCase();
    if (["bold", "italic", "underline", "fontsize", "fontname", "forecolor", "hilitecolor", "backcolor"].includes(normalized)) {
      return ({ fontsize: "fontSize", fontname: "fontFamily", forecolor: "foreColor", hilitecolor: "hiliteColor", backcolor: "hiliteColor" } as Record<string, InlineFormat>)[normalized] ?? normalized as InlineFormat;
    }
    return null;
  }

  private isFormatActive(target: TextTarget, format: InlineFormat): boolean {
    const style = window.getComputedStyle(target.node.parentElement!);
    if (format === "bold") return Number.parseInt(style.fontWeight, 10) >= 600 || style.fontWeight === "bold";
    if (format === "italic") return style.fontStyle === "italic" || style.fontStyle === "oblique";
    if (format === "underline") {
      const parent = target.node.parentElement;
      return Boolean(parent && this.findUnderlineSource(parent, parent.closest(".hwe-page-inner")));
    }
    return false;
  }

  private setInlineFormat(span: HTMLElement, format: InlineFormat, value: string, remove: boolean): void {
    const styles: Record<InlineFormat, [keyof CSSStyleDeclaration, string]> = {
      bold: ["fontWeight", "bold"], italic: ["fontStyle", "italic"], underline: ["textDecoration", "underline"],
      fontSize: ["fontSize", ""], fontFamily: ["fontFamily", ""], foreColor: ["color", ""], hiliteColor: ["backgroundColor", ""],
    };
    const [property, defaultValue] = styles[format];
    const removeValue: Partial<Record<InlineFormat, string>> = {
      bold: "normal", italic: "normal", underline: "none",
    };
    (span.style[property] as string) = remove ? removeValue[format] ?? "" : value || defaultValue;
    if (format === "fontSize") span.setAttribute("data-hwe-user-font-size", "true");
  }

  private applyParagraphCommand(command: string, range: Range, root: HTMLElement): HTMLElement[] {
    const alignment: Record<string, string> = {
      justifyleft: "left", justifycenter: "center", justifyright: "right", justifyfull: "justify",
    };
    const value = alignment[command.toLowerCase()];
    if (!value) return [];
    const blocks = range.collapsed
      ? this.getCollapsedParagraph(range, root)
      : this.documentSelection.getSelectedParagraphs(range);
    const flowIds = new Set(blocks.map((block) => block.getAttribute("data-hwe-text-flow-id")).filter((id): id is string => Boolean(id)));
    const affected = new Set(blocks);
    if (flowIds.size) for (const block of Array.from(root.querySelectorAll<HTMLElement>(BLOCKS))) {
      if (block.closest(".hwe-page-inner") && !block.closest("[contenteditable='false'], [hidden], [data-hwe-api-header], [data-hwe-dynamic-header], [data-hwe-runtime-page-header]") &&
          !block.matches(".hwe-page, .hwe-page-inner, .hwe-workspace, .hwe-text-flow-column, [data-hwe-generated-wrapper], [data-hwe-keep-together]") &&
          !(block.tagName === "DIV" && block.querySelector(BLOCKS)) &&
          block.getAttribute("data-hwe-text-flow-id") && flowIds.has(block.getAttribute("data-hwe-text-flow-id")!)) affected.add(block);
    }
    affected.forEach((block) => { block.style.textAlign = value; });
    return [...affected];
  }

  private getCollapsedParagraph(range: Range, root: HTMLElement): HTMLElement[] {
    const node = range.startContainer;
    const element = node.nodeType === Node.ELEMENT_NODE ? node as HTMLElement : node.parentElement;
    const block = element?.closest<HTMLElement>(BLOCKS) ?? null;
    if (!block || !root.contains(block) || !block.closest(".hwe-page-inner") || block.closest("[contenteditable='false'], [hidden], [data-hwe-api-header], [data-hwe-dynamic-header], [data-hwe-runtime-page-header]")) return [];
    if (block.matches(".hwe-page, .hwe-page-inner, .hwe-workspace, .hwe-text-flow-column, [data-hwe-generated-wrapper], [data-hwe-keep-together]")) return [];
    if (block.tagName === "DIV" && block.querySelector(BLOCKS)) return [];
    return [block];
  }
}

interface TextTarget { node: Text; start: number; end: number; }
