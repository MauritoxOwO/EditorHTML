/**
 * index.ts
 * Entry point del PCF. Delega toda la lógica en EditorComponent.
 */

import { EditorComponent } from "./EditorComponent";

export class HtmlWordEditor
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
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
    // El editor gestiona su propio estado.
    // Si se necesita reaccionar a cambios externos de propiedades,
    // se puede añadir lógica aquí.
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
    this.editor?.destroy();
  }
}

interface IInputs {
  htmlContent: ComponentFramework.PropertyTypes.StringProperty;
  entityName:  ComponentFramework.PropertyTypes.StringProperty;
  fieldName:   ComponentFramework.PropertyTypes.StringProperty;
}

interface IOutputs {}
