import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { EditorComponent } from "./app/EditorComponent";


export class EditorHTML2 implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private editor!: EditorComponent;
    private notifyOutputChanges!: () => void;
    // Sin estado hasta que el editor haya terminado de cargar el documento.
    private documentDirty: boolean | undefined;
    private destroyed = false;

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    context.mode.trackContainerResize(true);
    this.notifyOutputChanges = notifyOutputChanged;
    this.editor = new EditorComponent(container, context, {
      onDirtyChanged: (dirty: boolean) => {
        this.setDocumentDirty(dirty);
      }
    });
    this.editor.resize(context.mode.allocatedWidth, context.mode.allocatedHeight);
    this.editor.init().catch((err) => {
      console.error("[HtmlWordEditor] init error:", err);
    });
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this.editor?.resize(context.mode.allocatedWidth, context.mode.allocatedHeight);
  }

  public getOutputs(): IOutputs {
    if (this.documentDirty === undefined) return {};
    return {
      documentDirty: this.documentDirty,
    };
  }

  private setDocumentDirty(value : boolean) : void{
    if (this.destroyed || this.documentDirty === value) return;

    this.documentDirty = value;
    this.notifyOutputChanges();
  }

  public destroy(): void {
    this.destroyed = true;
    this.editor?.destroy();
  }
}
