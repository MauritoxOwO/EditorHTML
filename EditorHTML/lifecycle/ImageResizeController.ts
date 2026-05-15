export interface ImageResizeControllerOptions {
  rootProvider: () => HTMLElement | null;
  onImageChanged: (image: HTMLImageElement) => void;
}

const MIN_IMAGE_WIDTH_PERCENT = 15;
const MAX_IMAGE_WIDTH_PERCENT = 100;
const IMAGE_ALIGNMENT_COMMANDS = new Set([
  "justifyLeft",
  "justifyCenter",
  "justifyRight",
  "justifyFull",
]);

export class ImageResizeController {
  private overlay: HTMLElement | null = null;
  private panel: HTMLElement | null = null;
  private slider: HTMLInputElement | null = null;
  private valueLabel: HTMLElement | null = null;
  private selectedImage: HTMLImageElement | null = null;
  private dragState:
    | {
        containerWidth: number;
        startClientX: number;
        startWidth: number;
      }
    | null = null;

  private readonly handleRootClick = (event: MouseEvent): void => this.onRootClick(event);
  private readonly handleDocumentPointerDown = (event: PointerEvent): void =>
    this.onDocumentPointerDown(event);
  private readonly handleDocumentPointerMove = (event: PointerEvent): void =>
    this.onDocumentPointerMove(event);
  private readonly handleDocumentPointerUp = (): void => this.onDocumentPointerUp();
  private readonly handleScrollOrResize = (): void => this.refresh();
  private readonly handleKeyDown = (event: KeyboardEvent): void => {
    if (event.key === "Escape") this.clearSelection();
  };

  constructor(private readonly options: ImageResizeControllerOptions) {}

  start(): void {
    const root = this.options.rootProvider();
    if (!root) return;

    root.addEventListener("click", this.handleRootClick, true);
    root.addEventListener("scroll", this.handleScrollOrResize, true);
    document.addEventListener("pointerdown", this.handleDocumentPointerDown, true);
    document.addEventListener("pointermove", this.handleDocumentPointerMove);
    document.addEventListener("pointerup", this.handleDocumentPointerUp);
    document.addEventListener("keydown", this.handleKeyDown);
    window.addEventListener("resize", this.handleScrollOrResize);
    window.addEventListener("scroll", this.handleScrollOrResize, true);
  }

  destroy(): void {
    const root = this.options.rootProvider();
    root?.removeEventListener("click", this.handleRootClick, true);
    root?.removeEventListener("scroll", this.handleScrollOrResize, true);
    document.removeEventListener("pointerdown", this.handleDocumentPointerDown, true);
    document.removeEventListener("pointermove", this.handleDocumentPointerMove);
    document.removeEventListener("pointerup", this.handleDocumentPointerUp);
    document.removeEventListener("keydown", this.handleKeyDown);
    window.removeEventListener("resize", this.handleScrollOrResize);
    window.removeEventListener("scroll", this.handleScrollOrResize, true);
    this.clearSelection();
  }

  clearSelection(): void {
    this.selectedImage?.classList.remove("hwe-image-selected");
    this.selectedImage = null;
    this.dragState = null;
    this.overlay?.remove();
    this.panel?.remove();
    this.overlay = null;
    this.panel = null;
    this.slider = null;
    this.valueLabel = null;
  }

  refresh(): void {
    if (!this.selectedImage || !this.isSelectableImage(this.selectedImage)) {
      this.clearSelection();
      return;
    }

    this.positionControls();
  }

  handleToolbarCommand(command: string): boolean {
    if (!this.selectedImage || !IMAGE_ALIGNMENT_COMMANDS.has(command)) return false;

    this.applyAlignment(this.selectedImage, command);
    this.positionControls();
    this.options.onImageChanged(this.selectedImage);
    return true;
  }

  private onRootClick(event: MouseEvent): void {
    const target = event.target as Element | null;
    const image = target?.closest?.("img") as HTMLImageElement | null;
    if (!image || !this.isSelectableImage(image)) return;

    event.preventDefault();
    this.selectImage(image);
  }

  private onDocumentPointerDown(event: PointerEvent): void {
    const target = event.target as HTMLElement | null;
    if (!target) return;

    if (target.closest(".hwe-image-resize-panel, .hwe-editor-header")) return;

    if (target.classList.contains("hwe-image-resize-handle") && this.selectedImage) {
      event.preventDefault();
      const imageRect = this.selectedImage.getBoundingClientRect();
      const containerRect = this.getImageContainer(this.selectedImage).getBoundingClientRect();
      this.dragState = {
        containerWidth: Math.max(1, containerRect.width),
        startClientX: event.clientX,
        startWidth: imageRect.width,
      };
      return;
    }

    if (!target.closest("img")) {
      this.clearSelection();
    }
  }

  private onDocumentPointerMove(event: PointerEvent): void {
    if (!this.dragState || !this.selectedImage) return;

    event.preventDefault();
    const nextWidth = this.dragState.startWidth + (event.clientX - this.dragState.startClientX);
    this.applyImageWidth(this.selectedImage, (nextWidth / this.dragState.containerWidth) * 100);
    this.positionControls();
  }

  private onDocumentPointerUp(): void {
    if (!this.dragState || !this.selectedImage) return;

    this.dragState = null;
    this.options.onImageChanged(this.selectedImage);
  }

  private selectImage(image: HTMLImageElement): void {
    if (this.selectedImage === image) {
      this.refresh();
      return;
    }

    this.selectedImage?.classList.remove("hwe-image-selected");
    this.selectedImage = image;
    image.classList.add("hwe-image-selected");
    this.ensureControls();
    this.syncSliderValue();
    this.positionControls();
  }

  private ensureControls(): void {
    if (!this.overlay) {
      this.overlay = document.createElement("div");
      this.overlay.className = "hwe-image-resize-overlay";

      const handle = document.createElement("div");
      handle.className = "hwe-image-resize-handle";
      handle.title = "Redimensionar imagen";
      this.overlay.appendChild(handle);
      this.getControlsHost().appendChild(this.overlay);
    }

    if (!this.panel) {
      this.panel = document.createElement("div");
      this.panel.className = "hwe-image-resize-panel";

      const shrinkButton = this.makeButton("25%", "Imagen al 25%", () => this.applySelectedWidth(25));
      const halfButton = this.makeButton("50%", "Imagen al 50%", () => this.applySelectedWidth(50));
      const fullButton = this.makeButton("100%", "Imagen al 100%", () => this.applySelectedWidth(100));
      const resetButton = this.makeButton("Auto", "Restablecer ancho automatico", () =>
        this.resetSelectedWidth()
      );

      this.slider = document.createElement("input");
      this.slider.type = "range";
      this.slider.min = String(MIN_IMAGE_WIDTH_PERCENT);
      this.slider.max = String(MAX_IMAGE_WIDTH_PERCENT);
      this.slider.step = "1";
      this.slider.title = "Ancho de imagen";
      this.slider.addEventListener("input", () => {
        if (!this.selectedImage || !this.slider) return;
        this.applyImageWidth(this.selectedImage, Number(this.slider.value));
        this.positionControls();
      });
      this.slider.addEventListener("change", () => {
        if (this.selectedImage) this.options.onImageChanged(this.selectedImage);
      });

      this.valueLabel = document.createElement("span");
      this.valueLabel.className = "hwe-image-resize-value";

      this.panel.append(shrinkButton, halfButton, fullButton, resetButton, this.slider, this.valueLabel);
      this.getControlsHost().appendChild(this.panel);
    }
  }

  private makeButton(label: string, title: string, onClick: () => void): HTMLButtonElement {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = label;
    button.title = title;
    button.addEventListener("click", (event) => {
      event.preventDefault();
      onClick();
    });
    return button;
  }

  private applySelectedWidth(percent: number): void {
    if (!this.selectedImage) return;

    this.applyImageWidth(this.selectedImage, percent);
    this.positionControls();
    this.options.onImageChanged(this.selectedImage);
  }

  private resetSelectedWidth(): void {
    if (!this.selectedImage) return;

    this.selectedImage.style.removeProperty("width");
    this.selectedImage.removeAttribute("data-hwe-image-user-sized");
    this.syncSliderValue();
    this.positionControls();
    this.options.onImageChanged(this.selectedImage);
  }

  private applyImageWidth(image: HTMLImageElement, percent: number): void {
    const clamped = Math.max(
      MIN_IMAGE_WIDTH_PERCENT,
      Math.min(MAX_IMAGE_WIDTH_PERCENT, Math.round(percent))
    );
    image.style.width = `${clamped}%`;
    image.style.height = "auto";
    image.style.maxWidth = "100%";
    image.setAttribute("data-hwe-image-user-sized", "true");
    this.syncSliderValue(clamped);
  }

  private applyAlignment(image: HTMLImageElement, command: string): void {
    image.style.display = "block";

    if (command === "justifyCenter") {
      image.style.marginLeft = "auto";
      image.style.marginRight = "auto";
      image.setAttribute("data-hwe-image-align", "center");
      return;
    }

    if (command === "justifyRight") {
      image.style.marginLeft = "auto";
      image.style.marginRight = "0";
      image.setAttribute("data-hwe-image-align", "right");
      return;
    }

    if (command === "justifyFull") {
      image.style.width = "100%";
      image.style.marginLeft = "0";
      image.style.marginRight = "0";
      image.setAttribute("data-hwe-image-user-sized", "true");
      image.setAttribute("data-hwe-image-align", "full");
      this.syncSliderValue(100);
      return;
    }

    image.style.marginLeft = "0";
    image.style.marginRight = "auto";
    image.setAttribute("data-hwe-image-align", "left");
  }

  private syncSliderValue(explicitPercent?: number): void {
    if (!this.selectedImage || !this.slider || !this.valueLabel) return;

    const percent = explicitPercent ?? this.getCurrentWidthPercent(this.selectedImage);
    this.slider.value = String(percent);
    this.valueLabel.textContent = `${percent}%`;
  }

  private getCurrentWidthPercent(image: HTMLImageElement): number {
    const container = this.getImageContainer(image);
    const containerWidth = Math.max(1, container.getBoundingClientRect().width);
    const imageWidth = image.getBoundingClientRect().width;
    return Math.max(
      MIN_IMAGE_WIDTH_PERCENT,
      Math.min(MAX_IMAGE_WIDTH_PERCENT, Math.round((imageWidth / containerWidth) * 100))
    );
  }

  private positionControls(): void {
    if (!this.selectedImage || !this.overlay || !this.panel) return;

    const rect = this.selectedImage.getBoundingClientRect();
    const host = this.getControlsHost();
    const hostRect = host.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0 || !this.intersects(rect, hostRect)) {
      this.clearSelection();
      return;
    }

    const toHostX = (value: number) => value - hostRect.left + host.scrollLeft;
    const toHostY = (value: number) => value - hostRect.top + host.scrollTop;
    const localLeft = toHostX(rect.left);
    const localTop = toHostY(rect.top);

    this.overlay.style.left = `${localLeft}px`;
    this.overlay.style.top = `${localTop}px`;
    this.overlay.style.width = `${rect.width}px`;
    this.overlay.style.height = `${rect.height}px`;

    const visibleTop = Math.max(rect.top, hostRect.top);
    const visibleBottom = Math.min(rect.bottom, hostRect.bottom);
    const controlsTop = this.getControlsTopBound(host, hostRect);
    const panelWidth = Math.max(260, this.panel.offsetWidth || 0);
    const panelHeight = Math.max(30, this.panel.offsetHeight || 0);
    const panelViewportTop = this.clamp(
      visibleTop - panelHeight - 8,
      controlsTop + 6,
      visibleBottom - panelHeight
    );
    const panelViewportLeft = this.clamp(
      rect.left,
      hostRect.left + 6,
      hostRect.right - panelWidth - 6
    );
    const panelTop = toHostY(panelViewportTop);
    const panelLeft = toHostX(panelViewportLeft);
    this.panel.style.left = `${panelLeft}px`;
    this.panel.style.top = `${panelTop}px`;
    this.syncSliderValue();
  }

  private intersects(rect: DOMRect, bounds: DOMRect): boolean {
    return (
      rect.bottom > bounds.top &&
      rect.top < bounds.bottom &&
      rect.right > bounds.left &&
      rect.left < bounds.right
    );
  }

  private clamp(value: number, min: number, max: number): number {
    if (max < min) return min;
    return Math.max(min, Math.min(max, value));
  }

  private getControlsTopBound(host: HTMLElement, hostRect: DOMRect): number {
    const header = host.querySelector<HTMLElement>(".hwe-editor-header");
    if (!header) return hostRect.top;

    const headerRect = header.getBoundingClientRect();
    if (!this.intersects(headerRect, hostRect)) return hostRect.top;

    return Math.min(hostRect.bottom, Math.max(hostRect.top, headerRect.bottom));
  }

  private getImageContainer(image: HTMLImageElement): HTMLElement {
    return (
      image.closest<HTMLElement>("td, th") ??
      image.closest<HTMLElement>(".hwe-page-inner") ??
      image.parentElement ??
      image
    );
  }

  private isSelectableImage(image: HTMLImageElement): boolean {
    const root = this.options.rootProvider();
    return (
      !!root &&
      root.contains(image) &&
      image.closest(".hwe-page-inner") !== null &&
      image.getAttribute("data-hwe-large-image-placeholder") !== "true"
    );
  }

  private getControlsHost(): HTMLElement {
    return this.options.rootProvider() ?? document.body;
  }
}
