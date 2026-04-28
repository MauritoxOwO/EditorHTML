import { Paginator } from "../Orquestador/Paginator";
import { Toolbar } from "../Resize/Toolbar";
import { fetchHtmlFromFileField, saveHtmlToFileField } from "../execCommand/fileApi";

type PcfContext = ComponentFramework.Context<IInputs>;
type StatusType = "success" | "error" | "saving" | "";

export class EditorComponent {
  private readonly container: HTMLElement;

  private root!: HTMLElement;
  private workspace!: HTMLElement;
  private statusMsg!: HTMLElement;
  private pageCountEl!: HTMLElement;

  private pages: HTMLElement[] = [];
  private paginator!: Paginator;
  private toolbar!: Toolbar;

  private readonly baseUrl: string;
  private readonly entityName: string;
  private readonly entityId: string;
  private readonly fieldName: string;

  private isDirty = false;

  constructor(container: HTMLElement, context: PcfContext) {
    this.container = container;

    const page = (context as unknown as {
      page: { getClientUrl: () => string; entityId?: string };
    }).page;

    this.baseUrl = page.getClientUrl();
    this.entityId = page.entityId ?? "";
    this.entityName = "mcdev_htmldevtests";
    this.fieldName = "mcdev_htmlarchivooriginal";
  }

  async init(): Promise<void> {
    this.buildShell();
    this.paginator = new Paginator(
      (html?: string) => this.createPageElement(html),
      (pages: HTMLElement[]) => this.onPagesChanged(pages)
    );
    await this.loadContent();
  }

  private buildShell(): void {
    this.container.innerHTML = "";
    this.container.style.cssText =
      "width:100%;height:100%;overflow:hidden;display:flex;flex-direction:column;";

    this.root = document.createElement("div");
    this.root.className = "hwe-root";

    this.toolbar = new Toolbar();
    const toolbarEl = this.toolbar.build();
    this.toolbar.getSaveButton().addEventListener("click", () => void this.save());
    this.root.appendChild(toolbarEl);

    this.workspace = document.createElement("div");
    this.workspace.className = "hwe-workspace";
    this.root.appendChild(this.workspace);

    const statusBar = document.createElement("div");
    statusBar.className = "hwe-statusbar";

    this.pageCountEl = document.createElement("span");
    this.pageCountEl.textContent = "Paginas: 0";
    statusBar.appendChild(this.pageCountEl);

    this.statusMsg = document.createElement("span");
    this.statusMsg.className = "hwe-status-msg";
    statusBar.appendChild(this.statusMsg);

    this.root.appendChild(statusBar);
    this.container.appendChild(this.root);
  }

  private async loadContent(): Promise<void> {
    this.setStatus("Cargando contenido...", "saving");

    try {
      const html = await fetchHtmlFromFileField(
        this.baseUrl,
        this.entityName,
        this.entityId,
        this.fieldName
      );

      this.renderInitial(html || "<p><br></p>");
      await this.waitFrames(2);
      await this.paginateContent();

      this.setStatus("", "");
    } catch (err) {
      this.setStatus(`Error al cargar: ${(err as Error).message}`, "error");
      this.renderInitial("<p><br></p>");
      this.paginator.setPages(this.pages);
