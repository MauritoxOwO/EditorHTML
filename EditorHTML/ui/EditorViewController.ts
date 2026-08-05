export class EditorViewController {
  private visualTabBtn!: HTMLButtonElement;

  constructor(private readonly header: HTMLElement) {}

  build(): void {
    const viewTabs = document.createElement("div");
    viewTabs.className = "hwe-view-tabs";

    this.visualTabBtn = this.makeViewTabButton("Editor");

    const versionLabel = document.createElement("span");
    versionLabel.className = "hwe-version-label";
    versionLabel.textContent = `v1.10`;
    

    viewTabs.appendChild(this.visualTabBtn);
    viewTabs.appendChild(versionLabel);
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
