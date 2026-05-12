const LARGE_DATA_IMAGE_BYTES = 2_500_000;
const IMAGE_WAIT_TIMEOUT_MS = 1200;

export class AssetLayoutManager {
  async waitForStableLayout(root: HTMLElement): Promise<void> {
    await Promise.all([this.waitForImages(root), this.waitForFonts()]);
    await this.waitFrames(2);
  }

  private async waitForImages(root: HTMLElement): Promise<void> {
    const images = Array.from(root.querySelectorAll("img"));
    const pending = images.filter((img) => !img.complete || img.naturalWidth === 0);
    if (pending.length === 0) return;

    await Promise.all(
      pending.map((img) => {
        if (this.isLargeDataImage(img)) {
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
  }

  private isLargeDataImage(img: HTMLImageElement): boolean {
    const src = img.getAttribute("src") ?? "";
    return src.startsWith("data:image/") && this.estimateDataUrlBytes(src) > LARGE_DATA_IMAGE_BYTES;
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
