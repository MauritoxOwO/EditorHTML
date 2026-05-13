export type EditorView = "visual" | "source";

export class EditorViewController {
  private visualTabBtn!: HTMLButtonElement;
  private sourceTabBtn!: HTMLButtonElement;

  constructor(
    private readonly header: HTMLElement,
    private readonly onSwitchView: (view: EditorView) => void
  ) {}

  build(): void {
    const viewTabs = document.createElement("div");
    viewTabs.className = "hwe-view-tabs";

    this.visualTabBtn = this.makeViewTabButton("Editor", "visual");
    this.sourceTabBtn = this.makeViewTabButton("HTML", "source");

    viewTabs.appendChild(this.visualTabBtn);
    viewTabs.appendChild(this.sourceTabBtn);
    this.header.appendChild(viewTabs);
  }

  update(activeView: EditorView, workspace: HTMLElement, sourceEditor: HTMLElement): void {
    this.visualTabBtn.classList.toggle("hwe-active", activeView === "visual");
    this.sourceTabBtn.classList.toggle("hwe-active", activeView === "source");
    workspace.hidden = activeView !== "visual";
    sourceEditor.hidden = activeView !== "source";
  }

  private makeViewTabButton(label: string, view: EditorView): HTMLButtonElement {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = label;
    button.addEventListener("click", () => this.onSwitchView(view));
    return button;
  }
}