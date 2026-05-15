import {
  getHweDebugReport,
  saveHweDebugReport,
  startHweDebugHeartbeat,
} from "../debug/DebugLogger";

type StatusType = "success" | "error" | "saving" | "";
type SetStatus = (message: string, type: StatusType) => void;

export class EditorDiagnosticsController {
  private stopDebugHeartbeat: (() => void) | null = null;
  private readonly handleBeforeUnload = (): void => {
    saveHweDebugReport(this.rootProvider());
  };

  constructor(
    private readonly rootProvider: () => HTMLElement | null,
    private readonly setStatus: SetStatus
  ) {}

  makeButton(): HTMLButtonElement {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "hwe-diagnostics-btn";
    button.textContent = "Copiar diagnostico";
    button.addEventListener("click", () => {
      void this.copyDiagnostics();
    });
    return button;
  }

  start(): void {
    this.stopDebugHeartbeat = startHweDebugHeartbeat(this.rootProvider);
    window.addEventListener("beforeunload", this.handleBeforeUnload);
  }

  destroy(): void {
    saveHweDebugReport(this.rootProvider());
    this.stopDebugHeartbeat?.();
    this.stopDebugHeartbeat = null;
    window.removeEventListener("beforeunload", this.handleBeforeUnload);
  }

  private async copyDiagnostics(): Promise<void> {
    const report = getHweDebugReport(this.rootProvider());
    saveHweDebugReport(this.rootProvider());
    const text = JSON.stringify(report, null, 2);

    try {
      await navigator.clipboard.writeText(text);
      this.setStatus("Diagnostico copiado.", "success");
    } catch {
      this.copyTextFallback(text);
      this.setStatus("Diagnostico copiado.", "success");
    }

    window.setTimeout(() => this.setStatus("", ""), 2500);
  }

  private copyTextFallback(text: string): void {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.setAttribute("readonly", "true");
    textarea.style.cssText =
      "position:fixed;left:-9999px;top:0;width:1px;height:1px;opacity:0;";
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand("copy");
    textarea.remove();
  }
}
