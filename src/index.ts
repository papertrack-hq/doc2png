import { textToPdf } from "./converters/text-to-pdf";
import { docxToHtml } from "./converters/docx-to-pdf";
import { pptxToPdf } from "./converters/pptx-to-pdf";
import { renderExcelWorkbookToPdf } from "./converters/excel-processor";
import * as XLSX from "xlsx";
import { RenderedPage, RenderOptions, RenderResult } from "./types";

export type { RenderedPage, RenderOptions, RenderResult };
export { sanitizeWinAnsi } from "./converters/sanitize";

const DEFAULT_DPI = 150;

async function loadMupdf(): Promise<any> {
  // Dynamic import keeps this an optional/lazy dependency load and works
  // around mupdf being ESM-only in some environments.
  return Function("return import('mupdf')")();
}

async function rasterizeMupdfDoc(doc: any, dpi: number): Promise<RenderedPage[]> {
  const mupdf: any = await loadMupdf();
  const totalPages = doc.countPages();
  if (totalPages <= 0) {
    throw new Error("Document has no renderable pages");
  }

  const scale = dpi / 72;
  const matrix = mupdf.Matrix.scale(scale, scale);
  const pages: RenderedPage[] = [];

  for (let i = 0; i < totalPages; i++) {
    const pageNumber = i + 1;
    const page = doc.loadPage(i);
    try {
      const pixmap = page.toPixmap(matrix, mupdf.ColorSpace.DeviceRGB, false);
      try {
        const pngBytes = pixmap.asPNG();
        pages.push({ pageNumber, png: Buffer.from(pngBytes) });
      } finally {
        pixmap.destroy();
      }
    } finally {
      page.destroy();
    }
  }
  doc.destroy();

  return pages;
}

/**
 * Renders every page of a document — PDF, DOCX/DOC, PPTX/PPT, XLSX/XLS,
 * TXT/CSV, or a direct image (PNG/JPG/WEBP) — into an array of PNG page
 * images.
 *
 * Caller owns storage: this function returns in-memory buffers only, it
 * never writes to disk, blob storage, or a database.
 *
 * @param fileBuffer  Raw file bytes.
 * @param fileName    Original file name (used to detect the format by extension).
 * @param options     Render options (DPI, optional footer branding on Excel/PPTX previews).
 */
export async function renderDocumentToPages(
  fileBuffer: Buffer,
  fileName: string,
  options: RenderOptions = {}
): Promise<RenderResult> {
  const dpi = options.dpi ?? DEFAULT_DPI;
  const ext = (fileName.split(".").pop() || "").toLowerCase();

  // 1. Direct image formats — pass through as a single page.
  if (["png", "jpg", "jpeg", "webp"].includes(ext)) {
    return { totalPages: 1, pages: [{ pageNumber: 1, png: fileBuffer }] };
  }

  // 2. Excel spreadsheets — rendered as a grid preview, one page per sheet.
  if (["xlsx", "xls"].includes(ext)) {
    let workbook: XLSX.WorkBook;
    try {
      workbook = XLSX.read(fileBuffer, { type: "buffer" });
    } catch (_err) {
      workbook = { SheetNames: ["Sheet1"], Sheets: { Sheet1: {} } } as XLSX.WorkBook;
    }
    const { pdfBuffer } = await renderExcelWorkbookToPdf(workbook, fileName, {
      footerText: options.footerText,
    });
    const mupdf: any = await loadMupdf();
    const doc = mupdf.Document.openDocument(new Uint8Array(pdfBuffer), "application/pdf");
    const pages = await rasterizeMupdfDoc(doc, dpi);
    return { totalPages: pages.length, pages };
  }

  // 3. Everything else gets normalized to a PDF/HTML document, then rasterized.
  const mupdf: any = await loadMupdf();
  let doc: any;

  if (["docx", "doc"].includes(ext)) {
    const htmlContent = await docxToHtml(fileBuffer, fileName);
    doc = mupdf.Document.openDocument(new TextEncoder().encode(htmlContent), "text/html");
    if (typeof doc.layout === "function") {
      doc.layout(595.28, 841.89, 11); // A4 @ 11pt base font
    }
  } else if (["txt", "csv"].includes(ext)) {
    const pdfBuffer = await textToPdf(fileBuffer.toString("utf-8"), fileName);
    doc = mupdf.Document.openDocument(new Uint8Array(pdfBuffer), "application/pdf");
  } else if (["pptx", "ppt"].includes(ext)) {
    const pdfBuffer = await pptxToPdf(fileBuffer, fileName, { footerText: options.footerText });
    doc = mupdf.Document.openDocument(new Uint8Array(pdfBuffer), "application/pdf");
  } else {
    // Assume PDF (or a PDF-compatible format mupdf can open directly).
    doc = mupdf.Document.openDocument(new Uint8Array(fileBuffer), "application/pdf");
  }

  const pages = await rasterizeMupdfDoc(doc, dpi);
  return { totalPages: pages.length, pages };
}
