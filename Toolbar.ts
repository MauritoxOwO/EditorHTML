
export class Toolbar {
  private toolbar!: HTMLElement;
  private saveBtn!: HTMLButtonElement;

  private commandButtons = new Map<string, HTMLButtonElement>;

  build(): HTMLElement {
    this.toolbar = document.createElement("div");
    this.toolbar.className = "hwe-toolbar";

    // Formato básico
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
    this.addCmdButton("insertUnorderedList", "insertUnorderedList", "• Lista", "Lista con viñetas");
    this.addCmdButton("insertOrderedList",   "insertOrderedList",   "1. Lista", "Lista numerada");
    this.addSep();

    // Fuente 
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

    // Tamaño de fuente
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
    const breakBtn = document.createElement("button");
    breakBtn.title = "Insertar salto de página manual";
    breakBtn.textContent = "⊞ Salto";
    breakBtn.addEventListener("mousedown", (e) => {
      e.preventDefault();

      document.execCommand(
        "insertHTML",
        false,
        '<div style="page-break-before:always"><br></div>'
      );
    });
    this.toolbar.appendChild(breakBtn);

    this.addSep();

    // btnGuardar
    this.saveBtn = document.createElement("button");
    this.saveBtn.className = "hwe-save-btn";
    this.saveBtn.textContent = "💾 Guardar";
    this.toolbar.appendChild(this.saveBtn);

    document.addEventListener("selectionchange", () =>
      this.updateActiveStates()
    );

    return this.toolbar;
  }

  getSaveButton(): HTMLButtonElement {
    return this.saveBtn;
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

  private applyFontSize(size: string): void {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return;

    const range = sel.getRangeAt(0);
    const span  = document.createElement("span");
    span.style.fontSize = size;

    try {
      range.surroundContents(span);
    } catch {
      const fragment = range.extractContents();
      span.appendChild(fragment);
      range.insertNode(span);
    }
  }
}
