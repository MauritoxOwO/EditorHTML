export function addOcrTextLayers(root: HTMLElement): void {
  root.querySelectorAll<HTMLImageElement>("img[data-hwe-ocr-text]").forEach((image) => {
    const text = image.getAttribute("data-hwe-ocr-text")?.trim();
    const parent = image.parentNode;
    if (!text || !parent) return;

    const wrapper = document.createElement("span");
    wrapper.className = "hwe-ocr-wrapper";
    parent.insertBefore(wrapper, image);
    wrapper.appendChild(image);

    const textLayer = document.createElement("span");
    textLayer.className = "hwe-ocr-layer";
    textLayer.textContent = text;
    wrapper.appendChild(textLayer);
  });
}
