import { PageSetup } from "../pagination/PageGeometry";
import { DocumentSerializer } from "./DocumentSerializer";

export interface PrintHtmlExportContext {
  root: HTMLElement;
  pages: HTMLElement[];
  pageSetup: PageSetup;
  additionalCss: string;
}

export class PrintHtmlExportService {
  constructor(private readonly documentSerializer: DocumentSerializer) {}

  createPrintHtml(context: PrintHtmlExportContext): string {
    return this.documentSerializer.collectPdfHtml(
      context.root,
      context.pages,
      context.pageSetup,
      context.additionalCss
    );
  }
}
