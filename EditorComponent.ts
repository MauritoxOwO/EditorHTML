
    page.appendChild(inner);
    return page;
  }

  private onPagesChanged(pages: HTMLElement[]): void {
    this.pages = pages;

    pages.forEach((page, index) => {
      if (!page.parentElement) {
        const previousPage = pages[index - 1];
        if (previousPage?.parentElement) {
          const divider = this.makePageDivider(index + 1);
          this.workspace.insertBefore(divider, previousPage.nextSibling);
          this.workspace.insertBefore(page, divider.nextSibling);
        } else {
          this.workspace.appendChild(page);
        }
      }
    });

    this.removeEmptyDividers();
    this.renumberDividers();
    this.updatePageCount();
  }

  private makePageDivider(pageNumber: number): HTMLElement {
    const divider = document.createElement("div");
    divider.className = "hwe-page-divider";
    divider.setAttribute("contenteditable", "false");

    const label = document.createElement("span");
    label.textContent = `Pagina ${pageNumber}`;
    divider.appendChild(label);

    return divider;
  }

  private removeEmptyDividers(): void {
    const dividers = Array.from(this.workspace.querySelectorAll(".hwe-page-divider"));
    dividers.forEach((divider) => {
      const next = divider.nextElementSibling;
      if (!next || !next.classList.contains("hwe-page")) divider.remove();
    });
  }

  private renumberDividers(): void {
    const dividers = this.workspace.querySelectorAll(".hwe-page-divider span");
    dividers.forEach((span, index) => {
      span.textContent = `Pagina ${index + 2}`;
    });
  }

  private pageOverflows(page: HTMLElement): boolean {
    const inner = this.getInner(page);
    if (!inner) return false;
    return inner.scrollHeight > page.clientHeight + 1;
  }

  private getLastMeaningfulChild(container: HTMLElement): ChildNode | null {
    const children = Array.from(container.childNodes).filter(
      (node) => !this.isEmptyNode(node)
    );
    return children.length > 1 ? children[children.length - 1] : null;
  }

  private isEmptyNode(node: ChildNode): boolean {
    if (node.nodeType === Node.TEXT_NODE) {
      return (node.textContent ?? "").trim() === "";
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return false;

    const el = node as HTMLElement;
    if (el.tagName === "BR") return true;

    return (
      el.tagName === "P" &&
      el.childNodes.length <= 1 &&
      (el.textContent ?? "").trim() === ""
    );
  }

  private getInner(page: HTMLElement): HTMLElement | null {
    return page.querySelector(".hwe-page-inner");
  }

  private waitFrames(n: number): Promise<void> {
    return new Promise((resolve) => {
      let count = 0;
      const tick = () => {
        count++;
        if (count >= n) {
          resolve();
          return;
        }
        requestAnimationFrame(tick);
      };
      requestAnimationFrame(tick);
    });
  }

  private onPageKeyDown(e: KeyboardEvent): void {
    if (e.ctrlKey && e.key.toLowerCase() === "s") {
      e.preventDefault();
      void this.save();
    }
  }

  private async save(): Promise<void> {
    const saveBtn = this.toolbar.getSaveButton();
    saveBtn.disabled = true;
    this.setStatus("Guardando...", "saving");

    try {
      const html = this.collectHtml();
      await saveHtmlToFileField(
        this.baseUrl,
        this.entityName,
        this.entityId,
        this.fieldName,
        html
      );

      this.isDirty = false;
      this.setStatus("Guardado correctamente", "success");
      window.setTimeout(() => this.setStatus("", ""), 3000);
    } catch (err) {
      this.setStatus(`Error al guardar: ${(err as Error).message}`, "error");
    } finally {
      saveBtn.disabled = false;
    }
  }

  private collectHtml(): string {
    return this.pages
      .map((page, index) => {
        const inner = this.getInner(page);
        const content = inner ? inner.innerHTML : page.innerHTML;
        if (index === 0) return content;
        return `<div style="page-break-before:always">${content}</div>`;
      })
      .join("\n");
  }

  private updatePageCount(): void {
    this.pageCountEl.textContent = `Paginas: ${this.pages.length}`;
  }

  private setStatus(message: string, type: StatusType): void {
    this.statusMsg.textContent = message;
    this.statusMsg.className = "hwe-status-msg" + (type ? ` ${type}` : "");
  }

  destroy(): void {
    this.paginator?.destroy();
    this.container.innerHTML = "";
  }
}

interface IInputs {
  htmlContent: ComponentFramework.PropertyTypes.StringProperty;
  entityName: ComponentFramework.PropertyTypes.StringProperty;
  fieldName: ComponentFramework.PropertyTypes.StringProperty;
}
