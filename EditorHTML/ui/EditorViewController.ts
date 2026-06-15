export class EditorViewController {
  private visualTabBtn!: HTMLButtonElement;

  constructor(private readonly header: HTMLElement) {}

  build(): void {
    const viewTabs = document.createElement("div");
    viewTabs.className = "hwe-view-tabs";

    this.visualTabBtn = this.makeViewTabButton("Editor");

    viewTabs.appendChild(this.visualTabBtn);
    this.header.appendChild(viewTabs);
  }

  update(workspace: HTMLElement): void {
    this.visualTabBtn.classList.add("hwe-active");
    workspace.hidden = false;
  }

  private makeViewTabButton(label: string): HTMLButtonElement {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = label;
    return button;
  }
}
