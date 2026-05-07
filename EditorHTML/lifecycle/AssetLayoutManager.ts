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
        if (typeof img.decode === "function") {
          return img.decode().catch(() => undefined);
        }

        return new Promise<void>((resolve) => {
          img.addEventListener("load", () => resolve(), { once: true });
          img.addEventListener("error", () => resolve(), { once: true });
        });
      })
    );
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
}