type OcrStatusType = "success" | "error" | "saving" | "";

type OcrWorkerRequest = {
  id: string;
  src: string;
};

type OcrWorkerResponse = {
  id: string;
  ok: boolean;
  text?: string;
  error?: string;
};

type PendingOcrRequest = {
  reject: (error: Error) => void;
  resolve: (text: string) => void;
};

const AUTO_OCR_DELAY_MS = 750;
const MAX_OCR_IMAGE_BYTES = 2_500_000;
const MAX_OCR_IMAGE_PIXELS = 2_500_000;
const SUPPORTED_OCR_MIME_TYPES = new Set([
  "image/png",
  "image/jpeg",
  "image/jpg",
  "image/webp",
  "image/bmp",
]);
const TERMINAL_OCR_STATES = new Set([
  "done",
  "empty",
  "error",
  "skipped-large",
  "skipped-unsupported",
]);

export interface ImageOcrServiceOptions {
  setStatus?: (message: string, type: OcrStatusType) => void;
}

export class ImageOcrService {
  private ocrWorker: Worker | null = null;
  private runPromise: Promise<void> | null = null;
  private autoQueueTimer: number | null = null;
  private requestId = 0;
  private readonly pendingRequests = new Map<string, PendingOcrRequest>();

  constructor(private readonly options: ImageOcrServiceOptions = {}) {}

  queue(root: HTMLElement): void {
    if (this.collectCandidateImages(root).images.length === 0) return;

    if (this.autoQueueTimer !== null) {
      window.clearTimeout(this.autoQueueTimer);
    }

    this.autoQueueTimer = window.setTimeout(() => {
      this.autoQueueTimer = null;
      void this.run(root);
    }, AUTO_OCR_DELAY_MS);
  }

  async run(root: HTMLElement, reportStatus = false): Promise<void> {
    if (this.runPromise) return this.runPromise;

    this.runPromise = this.recognizeImages(root, reportStatus).finally(() => {
      this.runPromise = null;
    });
    return this.runPromise;
  }

  destroy(): void {
    if (this.autoQueueTimer !== null) {
      window.clearTimeout(this.autoQueueTimer);
      this.autoQueueTimer = null;
    }
    this.rejectPendingRequests(new Error("OCR service destroyed."));
    this.ocrWorker?.terminate();
    this.ocrWorker = null;
    this.runPromise = null;
  }

  private async recognizeImages(root: HTMLElement, reportStatus: boolean): Promise<void> {
    const { images, skippedLarge } = this.collectCandidateImages(root);
    if (images.length === 0) {
      if (reportStatus && skippedLarge > 0) this.reportSkippedLarge(skippedLarge);
      return;
    }

    if (reportStatus) this.options.setStatus?.("Ejecutando OCR de imagenes...", "saving");

    let recognized = 0;
    for (const image of images) {
      image.setAttribute("data-hwe-ocr-state", "running");
      try {
        const text = this.normalizeText(await this.recognize(image.src));
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
      await this.yieldToBrowser();
    }

    if (reportStatus && recognized > 0) {
      this.options.setStatus?.(`OCR completado en ${recognized} imagen(es).`, "success");
    } else if (reportStatus && skippedLarge > 0) {
      this.reportSkippedLarge(skippedLarge);
    }
  }

  private collectCandidateImages(root: HTMLElement): {
    images: HTMLImageElement[];
    skippedLarge: number;
  } {
    const images: HTMLImageElement[] = [];
    let skippedLarge = 0;

    Array.from(root.querySelectorAll<HTMLImageElement>("img")).forEach((image) => {
      const src = image.getAttribute("src") ?? "";
      const state = image.getAttribute("data-hwe-ocr-state");
      if (
        !src.startsWith("data:image/") ||
        image.getAttribute("data-hwe-ocr-text") ||
        state === "running" ||
        (state !== null && TERMINAL_OCR_STATES.has(state))
      ) {
        return;
      }

      const skipReason = this.getSkipReason(image, src);
      if (skipReason) {
        image.setAttribute("data-hwe-ocr-state", skipReason);
        if (skipReason === "skipped-large") skippedLarge++;
        return;
      }

      images.push(image);
    });

    return { images, skippedLarge };
  }

  private getSkipReason(
    image: HTMLImageElement,
    src: string
  ): "skipped-large" | "skipped-unsupported" | null {
    const mimeType = this.getDataUrlMimeType(src);
    if (mimeType && !SUPPORTED_OCR_MIME_TYPES.has(mimeType)) {
      return "skipped-unsupported";
    }

    if (this.estimateDataUrlBytes(src) > MAX_OCR_IMAGE_BYTES) {
      return "skipped-large";
    }

    const pixels = image.naturalWidth * image.naturalHeight;
    if (pixels > MAX_OCR_IMAGE_PIXELS) {
      return "skipped-large";
    }

    return null;
  }

  private getDataUrlMimeType(src: string): string | null {
    const match = /^data:([^;,]+)/i.exec(src);
    return match?.[1]?.toLowerCase() ?? null;
  }

  private estimateDataUrlBytes(src: string): number {
    const commaIndex = src.indexOf(",");
    const payload = commaIndex >= 0 ? src.slice(commaIndex + 1) : src;
    if (/^data:[^;,]+;base64,/i.test(src)) {
      return Math.floor(payload.length * 0.75);
    }

    return payload.length;
  }

  private recognize(src: string): Promise<string> {
    const worker = this.getWorker();
    const id = `ocr-${++this.requestId}`;
    const request: OcrWorkerRequest = { id, src };

    return new Promise((resolve, reject) => {
      this.pendingRequests.set(id, { resolve, reject });
      try {
        worker.postMessage(request);
      } catch (error) {
        this.pendingRequests.delete(id);
        reject(error instanceof Error ? error : new Error(String(error)));
      }
    });
  }

  private getWorker(): Worker {
    if (!this.ocrWorker) {
      this.ocrWorker = new Worker(new URL("./OcrWorker.ts", import.meta.url), {
        type: "module",
      });
      this.ocrWorker.onmessage = (event: MessageEvent<OcrWorkerResponse>) => {
        this.handleWorkerMessage(event.data);
      };
      this.ocrWorker.onerror = (event) => {
        const message = event.message || "OCR worker error.";
        this.rejectPendingRequests(new Error(message));
        this.ocrWorker?.terminate();
        this.ocrWorker = null;
      };
    }

    return this.ocrWorker;
  }

  private handleWorkerMessage(response: OcrWorkerResponse): void {
    const pending = this.pendingRequests.get(response.id);
    if (!pending) return;

    this.pendingRequests.delete(response.id);
    if (response.ok) {
      pending.resolve(response.text ?? "");
      return;
    }

    pending.reject(new Error(response.error ?? "OCR failed."));
  }

  private rejectPendingRequests(error: Error): void {
    this.pendingRequests.forEach((pending) => pending.reject(error));
    this.pendingRequests.clear();
  }

  private reportSkippedLarge(count: number): void {
    this.options.setStatus?.(
      `OCR omitido en ${count} imagen(es) grandes para evitar bloquear el editor.`,
      "success"
    );
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

  private async yieldToBrowser(): Promise<void> {
    await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
  }
}
