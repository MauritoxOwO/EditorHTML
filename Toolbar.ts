/**
 * Toolbar.ts
 *
 * Construye la barra de herramientas del editor y vincula
 * los comandos de formato mediante document.execCommand.
 *
 * Expone:
 *  - build(): HTMLElement   → el nodo toolbar listo para insertar
 *  - getSaveButton()        → referencia al botón Guardar
 *  - updateActiveStates()   → resalta botones según la selección actual
 */

export class Toolbar {
  private toolbar!: HTMLElement;
  private saveBtn!: HTMLButtonElement;

  // Mapa comando → botón, para resaltar el estado activo
  private commandButtons: Map<string, HTMLButtonElement> = new Map();

  build(): HTMLElement {
    this.toolbar = document.createElement("div");
    this.toolbar.className = "hwe-toolbar";

    // ── Formato de texto ──────────────────────────────────────
    this.addCmdButton("B",  "bold",      "<b>N</b>",  "Negrita (Ctrl+B)");
    this.addCmdButton("I",  "italic",    "<i>K</i>",  "Cursiva (Ctrl+I)");
    this.addCmdButton("U",  "underline", "<u>S</u>",  "Subrayado (Ctrl+U)");
    this.addSep();

    // ── Alineación ────────────────────────────────────────────
    this.addCmdButton("justifyLeft",   "justifyLeft",   "≡L", "Alinear izquierda");
    this.addCmdButton("justifyCenter", "justifyCenter", "≡C", "Centrar");
    this.addCmdButton("justifyRight",  "justifyRight",  "≡R", "Alinear derecha");
    this.addCmdButton("justifyFull",   "justifyFull",   "≡J", "Justificar");
    this.addSep();

    // ── Listas ────────────────────────────────────────────────
    this.addCmdButton("insertUnorderedList", "insertUnorderedList", "• Lista", "Lista con viñetas");
    this.addCmdButton("insertOrderedList",   "insertOrderedList",   "1. Lista", "Lista numerada");
    this.addSep();

    // ── Fuente ────────────────────────────────────────────────
    const fontSelect = this.makeSelect(
      "Fuente",
      [
        { value: "Calibri",          label: "Calibri"          },
        { value: "Arial",            label: "Arial"            },
        { value: "Times New Roman",  label: "Times New Roman"  },
        { value: "Georgia",          label: "Georgia"          },
        { value: "Courier New",      label: "Courier New"      },
        { value: "Verdana",          label: "Verdana"          },
      ],
      (value) => document.execCommand("fontName", false, value)
    );
    this.toolbar.appendChild(fontSelect);

    // ── Tamaño ────────────────────────────────────────────────
    const sizeSelect = this.makeSelect(
      "Tamaño",
      [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48].map((s) => ({
        value: String(s),
        label: String(s),
        selected: s === 11,
      })),
      (value) => this.applyFontSize(value + "pt")
    );
    this.toolbar.appendChild(sizeSelect);

    this.addSep();

    // ── Color de texto ────────────────────────────────────────
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

    // ── Resaltado ─────────────────────────────────────────────
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

    // ── Insertar salto de página manual ──────────────────────
    const breakBtn = document.createElement("button");
    breakBtn.title = "Insertar salto de página manual";
    breakBtn.textContent = "⊞ Salto";
    breakBtn.addEventListener("mousedown", (e) => {
      e.preventDefault();
      // Insertar un div con page-break-before en la posición del cursor.
      // El Paginator lo detectará como separador al guardar/recargar.
      document.execCommand(
        "insertHTML",
        false,
        '<div style="page-break-before:always"><br></div>'
      );
    });
    this.toolbar.appendChild(breakBtn);

    this.addSep();

    // ── Guardar ───────────────────────────────────────────────
    this.saveBtn = document.createElement("button");
    this.saveBtn.className = "hwe-save-btn";
    this.saveBtn.textContent = "💾 Guardar";
    this.toolbar.appendChild(this.saveBtn);

    // Actualizar estado activo al cambiar selección
    document.addEventListener("selectionchange", () =>
      this.updateActiveStates()
    );

    return this.toolbar;
  }

  getSaveButton(): HTMLButtonElement {
    return this.saveBtn;
  }

  /** Resalta los botones que corresponden al formato activo en la selección. */
  updateActiveStates(): void {
    this.commandButtons.forEach((btn, command) => {
      try {
        const active = document.queryCommandState(command);
        btn.classList.toggle("hwe-active", active);
      } catch {
        // queryCommandState puede lanzar en algunos contextos
      }
    });
  }

  // ── Helpers privados ──────────────────────────────────────────

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
      // Prevenir pérdida de foco en el editor
      e.preventDefault();
      document.execCommand(command, false);
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

  /**
   * Aplica tamaño de fuente envolviendo la selección en un <span>.
   * document.execCommand("fontSize") solo acepta 1-7 (legacy HTML),
   * así que usamos span con style directo.
   */
  private applyFontSize(size: string): void {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return;

    const range = sel.getRangeAt(0);
    const span  = document.createElement("span");
    span.style.fontSize = size;

    try {
      range.surroundContents(span);
    } catch {
      // surroundContents falla si la selección cruza elementos.
      // Fallback: extraer y envolver.
      const fragment = range.extractContents();
      span.appendChild(fragment);
      range.insertNode(span);
    }
  }
}
