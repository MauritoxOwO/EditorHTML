import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { EditorComponent } from "./lifecycle/EditorComponent";


export class EditorHTML2 implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private editor!: EditorComponent;

  public init(
    context: ComponentFramework.Context<IInputs>,
    _notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.editor = new EditorComponent(container, context);
    this.editor.init().catch((err) => {
      console.error("[HtmlWordEditor] init error:", err);
    });
  }

  public updateView(_context: ComponentFramework.Context<IInputs>): void {
    return;
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
    this.editor?.destroy();
  }
}


