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
    context.mode.trackContainerResize(true);
    this.editor = new EditorComponent(container, context);
    this.editor.resize(context.mode.allocatedWidth, context.mode.allocatedHeight);
    this.editor.init().catch((err) => {
      console.error("[HtmlWordEditor] init error:", err);
    });
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this.editor?.resize(context.mode.allocatedWidth, context.mode.allocatedHeight);
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
    this.editor?.destroy();
  }
}

