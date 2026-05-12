type OcrStatusType = "success" | "error" | "saving" | "";

type OcrWorker = {
  recognize: (image: string) => Promise<{ data?: { text?: string } }>;
  terminate: () => Promise<void>;
};

export interface ImageOcrServiceOptions {
  setStatus?: (message: string, type: OcrStatusType) => void;
}

export class ImageOcrService {
  private workerPromise: Promise<OcrWorker> | null = null;
  private runPromise: Promise<void> | null = null;

  constructor(private readonly options: ImageOcrServiceOptions = {}) {}

  queue(root: HTMLElement): void {
    if (this.getCandidateImages(root).length === 0) return;

    window.setTimeout(() => {
      void this.run(root);
    }, 0);
  }

  async run(root: HTMLElement, reportStatus = false): Promise<void> {
    if (this.runPromise) return this.runPromise;

    this.runPromise = this.recognizeImages(root, reportStatus).finally(() => {
      this.runPromise = null;
    });
    return this.runPromise;
  }

  destroy(): void {
    void this.workerPromise?.then((worker) => worker.terminate());
    this.workerPromise = null;
    this.runPromise = null;
  }

  private async recognizeImages(root: HTMLElement, reportStatus: boolean): Promise<void> {
    const images = this.getCandidateImages(root);
    if (images.length === 0) return;

    if (reportStatus) this.options.setStatus?.("Ejecutando OCR de imagenes...", "saving");

    let recognized = 0;
    try {
      const worker = await this.getWorker();

      for (const image of images) {
        image.setAttribute("data-hwe-ocr-state", "running");
        try {
          const result = await worker.recognize(image.src);
          const text = this.normalizeText(result.data?.text ?? "");
          if (text) {
            image.setAttribute("data-hwe-ocr-text", text);
            if (!image.getAttribute("alt")) image.setAttribute("alt", this.truncateText(text));
            if (!image.getAttribute("title")) image.setAttribute("title", this.truncateText(text));
            image.setAttribute("data-hwe-ocr-state", "done");
            recognized++;
          } else {
            image.setAttribute("data-hwe-ocr-state", "empty");
          }
        } catch (error) {
          console.warn("[HtmlWordEditor] image OCR failed:", error);
          image.setAttribute("data-hwe-ocr-state", "error");
        }
      }
    } catch (error) {
      console.warn("[HtmlWordEditor] OCR worker unavailable:", error);
      images.forEach((image) => image.setAttribute("data-hwe-ocr-state", "error"));
      if (reportStatus) {
        this.options.setStatus?.("OCR no disponible; se exportara el PDF sin texto OCR.", "error");
      }
      return;
    }

    if (reportStatus && recognized > 0) {
      this.options.setStatus?.(`OCR completado en ${recognized} imagen(es).`, "success");
    }
  }

  private getCandidateImages(root: HTMLElement): HTMLImageElement[] {
    return Array.from(root.querySelectorAll<HTMLImageElement>("img")).filter((image) => {
      const src = image.getAttribute("src") ?? "";
      const state = image.getAttribute("data-hwe-ocr-state");
      return (
        src.startsWith("data:image/") &&
        !image.getAttribute("data-hwe-ocr-text") &&
        state !== "running" &&
        state !== "done" &&
        state !== "empty"
      );
    });
  }

  private async getWorker(): Promise<OcrWorker> {
    if (!this.workerPromise) {
      this.workerPromise = import("tesseract.js").then(async (module) => {
        const tesseract = module as unknown as {
          createWorker: (
            langs?: string | string[],
            oem?: number,
            options?: { logger?: (message: unknown) => void }
          ) => Promise<OcrWorker>;
        };
        return tesseract.createWorker(["spa", "eng"], 1, {
          logger: () => undefined,
        });
      });
    }

    return this.workerPromise;
  }

  private normalizeText(text: string): string {
    return text
      .split(/\r?\n/)
      .map((line) => line.replace(/\s+/g, " ").trim())
      .filter(Boolean)
      .join("\n")
      .trim();
  }

  private truncateText(text: string): string {
    return text.length > 512 ? `${text.slice(0, 509)}...` : text;
  }
}
