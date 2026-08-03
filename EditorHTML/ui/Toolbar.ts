
import type { TextCase } from "../controllers/TextSelectionFormatter";

export interface ParagraphStyleOption {
  label: string;
  className: string;
}

export interface SelectionFormattingInterface {
  fontSize?: string;
  fontFamily?: string;
  paragraphStyle?: string;
}

export const CLEAR_PARAGRAPH_STYLE_VALUE = "__hwe-clear-paragraph-style";
const DEFAULT_FONT_FAMILIES = [
  "Calibri",
  "Arial",
  "Times New Roman",
  "Georgia",
  "Courier New",
  "Verdana",
];

export interface ToolbarOptions {
  onUndo?: () => void;
  onRedo?: () => void;
  onToggleUnorderedList?: () => void;
  onToggleOrderedList?: () => void;
  onInsertTable?: () => void;
  onInsertRowAfter?: () => void;
  onDeleteRow?: () => void;
  onInsertPageBreak?: () => void;
  onApplyParagraphStyle?: (className: string) => void;
  onApplyTextCase?: (textCase: TextCase) => void;
  onApplyFontSize?: (fontSize: string) => void;
  onCommand?: (command: string) => boolean;
}

export class Toolbar {
  private toolbar!: HTMLElement;
  private saveBtn!: HTMLButtonElement;
  private undoBtn!: HTMLButtonElement;
  private redoBtn!: HTMLButtonElement;
  private styleSelect!: HTMLSelectElement;
  private sizeSelect!: HTMLSelectElement;
  private fontSelect!: HTMLSelectElement;

  private commandButtons = new Map<string, HTMLButtonElement>();
  private readonly handleSelectionChange = (): void => this.updateActiveStates();

  constructor(private readonly options: ToolbarOptions = {}) {}

  build(): HTMLElement {
    this.toolbar = document.createElement("div");
    this.toolbar.className = "hwe-toolbar";

    // Formato básico
    // Historial
    this.undoBtn = this.addActionButton("↶", "Deshacer (Ctrl+Z)", () => this.options.onUndo?.());
    this.redoBtn = this.addActionButton("↷", "Rehacer (Ctrl+Y)", () => this.options.onRedo?.());
    this.setHistoryAvailability(false, false);
    this.addSep();

    this.addCmdButton("B",  "bold",      "<b>N</b>",  "Negrita (Ctrl+B)");
    this.addCmdButton("I",  "italic",    "<i>K</i>",  "Cursiva (Ctrl+I)");
    this.addCmdButton("U",  "underline", "<u>S</u>",  "Subrayado (Ctrl+U)");
    this.addSep();

    // Alineación
    this.addCmdButton("justifyLeft",   "justifyLeft",   "≡L", "Alinear izquierda");
    this.addCmdButton("justifyCenter", "justifyCenter", "≡C", "Centrar");
    this.addCmdButton("justifyRight",  "justifyRight",  "≡R", "Alinear derecha");
    this.addCmdButton("justifyFull",   "justifyFull",   "≡J", "Justificar");
    this.addSep();

    // Listas
    this.addStatefulActionButton(
      "insertUnorderedList",
      "• Lista",
      "Lista con viñetas",
      () => this.options.onToggleUnorderedList?.()
    );
    this.addStatefulActionButton(
      "insertOrderedList",
      "1. Lista",
      "Lista numerada",
      () => this.options.onToggleOrderedList?.()
    );
    this.addSep();

    // Transformacion del texto seleccionado
    this.addActionButton("ABC", "Convertir texto seleccionado a mayusculas", () =>
      this.options.onApplyTextCase?.("uppercase")
    );
    this.addActionButton("abc", "Convertir texto seleccionado a minusculas", () =>
      this.options.onApplyTextCase?.("lowercase")
    );
    this.addSep();

    // Fuente 
    this.styleSelect = this.makeSelect(
      "Estilo de parrafo",
      [{ value: "", label: "Estilo", selected: true }],
      (value) => {
        if (value) this.options.onApplyParagraphStyle?.(value);
      }
    );
    this.styleSelect.className = "hwe-style-select";
    this.styleSelect.disabled = true;
    this.toolbar.appendChild(this.styleSelect);
    this.addSep();

    this.fontSelect = this.makeSelect(
      "Fuente",
      this.makeFontFamilyOptions(DEFAULT_FONT_FAMILIES),
      (value) => document.execCommand("fontName", false, value)
    );
    this.toolbar.appendChild(this.fontSelect);

    // Tamaño de fuente
    this.sizeSelect = this.makeSelect(
      "Tamaño",
      ["", 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48].map((s) => ({
        value: String(s),
        label: String(s),
        selected: s === "",
      })),
      (value) => this.options.onApplyFontSize?.(value + "pt")
    );
    this.toolbar.appendChild(this.sizeSelect);

    this.addSep();

    // Color de texto
    const colorBtn = document.createElement("button");
    colorBtn.title = "Color de texto";
    colorBtn.innerHTML = "A";
    colorBtn.style.cssText = "position:relative;overflow:hidden;";

    const colorInput = document.createElement("input");
    colorInput.type = "color";
    colorInput.value = "#000000";
    colorInput.style.cssText =
      "position:absolute;top:0;left:0;width:100%;height:100%;opacity:0;cursor:pointer;";
    colorInput.addEventListener("input", () => {
      document.execCommand("foreColor", false, colorInput.value);
    });
    colorBtn.appendChild(colorInput);
    this.toolbar.appendChild(colorBtn);

    // Resaltar
    const highlightBtn = document.createElement("button");
    highlightBtn.title = "Color de fondo de texto";
    highlightBtn.innerHTML = "🖊";
    highlightBtn.style.cssText = "position:relative;overflow:hidden;";

    const highlightInput = document.createElement("input");
    highlightInput.type = "color";
    highlightInput.value = "#ffff00";
    highlightInput.style.cssText =
      "position:absolute;top:0;left:0;width:100%;height:100%;opacity:0;cursor:pointer;";
    highlightInput.addEventListener("input", () => {
      document.execCommand("hiliteColor", false, highlightInput.value);
    });
    highlightBtn.appendChild(highlightInput);
    this.toolbar.appendChild(highlightBtn);

    this.addSep();

    // Salto de página manual 
    this.addActionButton("Tabla", "Insertar tabla", () => this.options.onInsertTable?.());
    this.addActionButton("+ Fila", "Insertar fila debajo", () => this.options.onInsertRowAfter?.());
    this.addActionButton("- Fila", "Eliminar fila", () => this.options.onDeleteRow?.());
    this.addSep();

    const breakBtn = document.createElement("button");
    breakBtn.title = "Insertar salto de página manual";
    breakBtn.textContent = "⊞ Salto";
    breakBtn.addEventListener("mousedown", (e) => {
      e.preventDefault();

      this.options.onInsertPageBreak?.();
    });
    this.toolbar.appendChild(breakBtn);

    this.addSep();

    // btnGuardar
    this.saveBtn = document.createElement("button");
    this.saveBtn.className = "hwe-save-btn";
    this.saveBtn.textContent = "💾 Guardar";
    this.toolbar.appendChild(this.saveBtn);

    document.addEventListener("selectionchange", this.handleSelectionChange);

    return this.toolbar;
  }

  setSelectionFormatting(state: SelectionFormattingInterface): void {
    const fontSize = this.getCleanFontSizeValue(state.fontSize);
    this.setSelectValue(this.sizeSelect, fontSize);
    this.setSelectValue(this.fontSelect, state.fontFamily);

    const paragraphStyle =
      state.paragraphStyle === "" ? CLEAR_PARAGRAPH_STYLE_VALUE : state.paragraphStyle;
    this.setSelectValue(this.styleSelect, paragraphStyle);
  }

  setHistoryAvailability(canUndo: boolean, canRedo: boolean): void {
    if (this.undoBtn) this.undoBtn.disabled = !canUndo;
    if (this.redoBtn) this.redoBtn.disabled = !canRedo;
  }

  getSaveButton(): HTMLButtonElement {
    return this.saveBtn;
  }

  setParagraphStyles(styles: ParagraphStyleOption[]): void {
    if (!this.styleSelect) return;

    this.styleSelect.innerHTML = "";
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = styles.length > 0 ? "Estilo" : "Sin estilos";
    placeholder.selected = true;
    this.styleSelect.appendChild(placeholder);

    if (styles.length > 0) {
      const clearOption = document.createElement("option");
      clearOption.value = CLEAR_PARAGRAPH_STYLE_VALUE;
      clearOption.textContent = "Sin estilo";
      this.styleSelect.appendChild(clearOption);
    }

    styles.forEach((style) => {
      const option = document.createElement("option");
      option.value = style.className;
      option.textContent = style.label;
      this.styleSelect.appendChild(option);
    });

    this.styleSelect.disabled = styles.length === 0;
  }

  setFontFamilies(fontFamilies: string[]): void {
    if (!this.fontSelect) return;

    const previousValue = this.fontSelect.value;
    const options = this.makeFontFamilyOptions([...DEFAULT_FONT_FAMILIES, ...fontFamilies]);
    this.fontSelect.innerHTML = "";
    options.forEach(({ value, label, selected }) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = label;
      if (selected) option.selected = true;
      this.fontSelect.appendChild(option);
    });

    if (previousValue && this.hasSelectValue(this.fontSelect, previousValue)) {
      this.fontSelect.value = previousValue;
    }
  }

  updateActiveStates(): void {
    this.commandButtons.forEach((btn, command) => {
      try {
        const active = document.queryCommandState(command);
        btn.classList.toggle("hwe-active", active);
      } catch (error) {
        console.log(error);
      }
    });
  }

  destroy(): void {
    document.removeEventListener("selectionchange", this.handleSelectionChange);
  }

// Funciones auxiliares
  private addCmdButton(
    id: string,
    command: string,
    html: string,
    title: string
  ): void {
    const btn = document.createElement("button");
    btn.innerHTML = html;
    btn.title = title;

    btn.addEventListener("mousedown", (e) => {
      e.preventDefault();
      if (this.options.onCommand?.(command)) return;
      document.execCommand(command, false);
    });

    this.commandButtons.set(command, btn);
    this.toolbar.appendChild(btn);
  }

  private addActionButton(
    label: string,
    title: string,
    onAction: () => void
  ): HTMLButtonElement {
    const btn = document.createElement("button");
    btn.textContent = label;
    btn.title = title;
    btn.addEventListener("mousedown", (e) => {
      e.preventDefault();
      onAction();
    });
    this.toolbar.appendChild(btn);
    return btn;
  }

  private addStatefulActionButton(
    command: string,
    label: string,
    title: string,
    onAction: () => void
  ): void {
    const btn = document.createElement("button");
    btn.textContent = label;
    btn.title = title;
    btn.addEventListener("mousedown", (event) => {
      event.preventDefault();
      onAction();
    });
    this.commandButtons.set(command, btn);
    this.toolbar.appendChild(btn);
  }

  private addSep(): void {
    const sep = document.createElement("div");
    sep.className = "hwe-sep";
    this.toolbar.appendChild(sep);
  }

  private makeSelect(
    title: string,
    options: { value: string; label: string; selected?: boolean }[],
    onChange: (value: string) => void
  ): HTMLSelectElement {
    const sel = document.createElement("select");
    sel.title = title;

    options.forEach(({ value, label, selected }) => {
      const opt = document.createElement("option");
      opt.value = value;
      opt.textContent = label;
      if (selected) opt.selected = true;
      sel.appendChild(opt);
    });

    sel.addEventListener("mousedown", (e) => e.stopPropagation());
    sel.addEventListener("change", () => onChange(sel.value));
    return sel;
  }

  private makeFontFamilyOptions(
    fontFamilies: string[]
  ): { value: string; label: string; selected?: boolean }[] {
    return this.dedupeFontFamilies(fontFamilies).map((fontFamily, index) => ({
      value: fontFamily,
      label: fontFamily,
      selected: index === 0,
    }));
  }

  private dedupeFontFamilies(fontFamilies: string[]): string[] {
    const seen = new Set<string>();
    return fontFamilies
      .map((fontFamily) => fontFamily.trim())
      .filter(Boolean)
      .filter((fontFamily) => {
        const key = fontFamily.toLowerCase();
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
  }

  private hasSelectValue(select: HTMLSelectElement, value: string): boolean {
    return Array.from(select.options).some((option) => option.value === value);
  }

  private setSelectValue(select: HTMLSelectElement | undefined, value?: string): void {
    if (!select) return;
    if (!value) {
      select.value = "";
      return;
    }

    const matchingOption = Array.from(select.options).find(
      (option) => option.value.toLowerCase() === value.toLowerCase()
    );
    select.value = matchingOption?.value ?? "";
  }

  private getCleanFontSizeValue(fontSize?: string): string {
    if (!fontSize) return "";

    const value = Number.parseFloat(fontSize);
    if (!Number.isFinite(value)) return "";

    const points = fontSize.endsWith("px") ? value * 0.75 : value;
    return String(Math.round(points * 100) / 100);
  }
}
