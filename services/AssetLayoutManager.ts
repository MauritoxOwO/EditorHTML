import { hweDebugLog, hweDebugStart } from "../debug/DebugLogger";

const LARGE_DATA_IMAGE_BYTES = 2_500_000;
const LARGE_IMAGE_PIXELS = 1_500_000;
const IMAGE_WAIT_TIMEOUT_MS = 1200;
const DETACHED_IMAGE_WAIT_TIMEOUT_MS = 2500;

type DetachedImageResolver = (id: string) => string | undefined;
type DetachedImageReadyHandler = (image: HTMLImageElement) => Promise<void> | void;

export class AssetLayoutManager {
  async waitForStableLayout(root: HTMLElement): Promise<void> {
    const done = hweDebugStart("assets.waitForStableLayout", {
      images: root.querySelectorAll("img").length,
    });
    await Promise.all([this.waitForImages(root), this.waitForFonts()]);
    await this.waitFrames(2);
    done({
      images: root.querySelectorAll("img").length,
    });
  }

  async hydrateDetachedImages(
    root: HTMLElement,
    resolveSrc: DetachedImageResolver,
    onImageReady: DetachedImageReadyHandler
  ): Promise<void> {
    const images = Array.from(
      root.querySelectorAll<HTMLImageElement>("img[data-hwe-large-image-placeholder='true']")
    );
    const done = hweDebugStart("assets.hydrateDetachedImages", {
      images: images.length,
    });

    let restored = 0;
    let failed = 0;

    for (const image of images) {
      await this.waitForIdle();
      if (!root.isConnected || !image.isConnected) continue;

      const id = image.getAttribute("data-hwe-large-image-id") ?? "";
      const src = resolveSrc(id);
      if (!src) {
        failed++;
        hweDebugLog("assets.hydrateDetachedImages.missingSource", { id });
        continue;
      }

      const previousSrc = image.getAttribute("src") ?? "";
      image.setAttribute("data-hwe-large-image-state", "hydrating");
      const loaded = await this.loadImageSrc(image, src);
      if (!loaded) {
        failed++;
        image.setAttribute("src", previousSrc);
        image.removeAttribute("data-hwe-large-image-state");
        hweDebugLog("assets.hydrateDetachedImages.failed", {
          id,
          reason: image.getAttribute("data-hwe-large-image-reason") ?? "",
        });
        continue;
      }

      image.removeAttribute("data-hwe-large-image-id");
      image.removeAttribute("data-hwe-large-image-placeholder");
      image.removeAttribute("data-hwe-large-image-reason");
      image.removeAttribute("data-hwe-large-image-state");
      restored++;

      try {
        await onImageReady(image);
      } catch (error) {
        hweDebugLog("assets.hydrateDetachedImages.onReadyError", {
          message: error instanceof Error ? error.message : String(error),
        });
      }
    }

    done({
      failed,
      restored,
    });
  }

  private async waitForImages(root: HTMLElement): Promise<void> {
    const images = Array.from(root.querySelectorAll("img"));
    const pending = images.filter((img) => !img.complete || img.naturalWidth === 0);
    const done = hweDebugStart("assets.waitForImages", {
      images: images.length,
      pending: pending.length,
      largeDataImages: images.filter((img) => this.isLargeDataImage(img)).length,
      placeholderImages: images.filter(
        (img) => img.getAttribute("data-hwe-large-image-placeholder") === "true"
      ).length,
    });
    if (pending.length === 0) {
      done({ skipped: true });
      return;
    }

    await Promise.all(
      pending.map((img) => {
        if (this.isDataImage(img)) {
          return this.waitForImageEvent(img, IMAGE_WAIT_TIMEOUT_MS);
        }

        if (typeof img.decode === "function") {
          return Promise.race([
            img.decode().catch(() => undefined),
            this.waitForTimeout(IMAGE_WAIT_TIMEOUT_MS),
          ]);
        }

        return this.waitForImageEvent(img, IMAGE_WAIT_TIMEOUT_MS);
      })
    );
    done({
      pending: pending.length,
    });
  }

  private isLargeDataImage(img: HTMLImageElement): boolean {
    const src = img.getAttribute("src") ?? "";
    if (!src.startsWith("data:image/")) return false;
    if (this.estimateDataUrlBytes(src) > LARGE_DATA_IMAGE_BYTES) return true;

    const width = this.getDeclaredDimension(img, "width") || img.naturalWidth || 0;
    const height = this.getDeclaredDimension(img, "height") || img.naturalHeight || 0;
    return width * height > LARGE_IMAGE_PIXELS;
  }

  private isDataImage(img: HTMLImageElement): boolean {
    return (img.getAttribute("src") ?? "").startsWith("data:image/");
  }

  private getDeclaredDimension(img: HTMLImageElement, name: "height" | "width"): number {
    const raw = img.getAttribute(name);
    return raw ? Number.parseFloat(raw) || 0 : 0;
  }

  private estimateDataUrlBytes(src: string): number {
    const commaIndex = src.indexOf(",");
    const payload = commaIndex >= 0 ? src.slice(commaIndex + 1) : src;
    if (/^data:[^;,]+;base64,/i.test(src)) {
      return Math.floor(payload.length * 0.75);
    }

    return payload.length;
  }

  private waitForImageEvent(img: HTMLImageElement, timeoutMs: number): Promise<void> {
    if (img.complete && img.naturalWidth > 0) return Promise.resolve();

    return new Promise<void>((resolve) => {
      let settled = false;
      let timeout = 0;
      const done = () => {
        if (settled) return;
        settled = true;
        window.clearTimeout(timeout);
        img.removeEventListener("load", done);
        img.removeEventListener("error", done);
        resolve();
      };
      timeout = window.setTimeout(done, timeoutMs);
      img.addEventListener("load", done, { once: true });
      img.addEventListener("error", done, { once: true });
    });
  }

  private loadImageSrc(img: HTMLImageElement, src: string): Promise<boolean> {
    return new Promise<boolean>((resolve) => {
      let settled = false;
      let timeout = 0;

      const done = (loaded: boolean) => {
        if (settled) return;
        settled = true;
        window.clearTimeout(timeout);
        img.removeEventListener("load", onLoad);
        img.removeEventListener("error", onError);
        resolve(loaded);
      };
      const onLoad = () => done(img.naturalWidth > 0);
      const onError = () => done(false);

      timeout = window.setTimeout(
        () => done(img.complete && img.naturalWidth > 0),
        DETACHED_IMAGE_WAIT_TIMEOUT_MS
      );
      img.addEventListener("load", onLoad, { once: true });
      img.addEventListener("error", onError, { once: true });
      img.setAttribute("src", src);

      if (img.complete) {
        window.requestAnimationFrame(() => done(img.naturalWidth > 0));
      }
    });
  }

  private waitForIdle(): Promise<void> {
    return new Promise((resolve) => {
      const idleWindow = window as Window & {
        requestIdleCallback?: (
          callback: () => void,
          options?: { timeout?: number }
        ) => number;
      };

      if (idleWindow.requestIdleCallback) {
        idleWindow.requestIdleCallback(() => resolve(), { timeout: 250 });
        return;
      }

      window.setTimeout(resolve, 0);
    });
  }

  private async waitForFonts(): Promise<void> {
    const fonts = document.fonts;
    if (!fonts?.ready) return;

    await Promise.race([
      fonts.ready.then(() => undefined),
      new Promise<void>((resolve) => window.setTimeout(resolve, 2500)),
    ]);
  }

  private waitFrames(count: number): Promise<void> {
    return new Promise((resolve) => {
      let frame = 0;
      const tick = () => {
        frame++;
        if (frame >= count) {
          resolve();
          return;
        }
        requestAnimationFrame(tick);
      };
      requestAnimationFrame(tick);
    });
  }

  private waitForTimeout(timeoutMs: number): Promise<void> {
    return new Promise((resolve) => window.setTimeout(resolve, timeoutMs));
  }
}
