import * as XLSX from "xlsx";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import { sanitizeWinAnsi } from "./sanitize";

export interface ExcelRenderOptions {
  /** Text shown in the page footer, e.g. "Rendered by Acme Inc". Omit for no footer. */
  footerText?: string;
}

/**
 * Renders an Excel workbook preview (grid lines, column letters, row numbers,
 * zebra shading, sheet tabs) into a PDF buffer — one page per sheet.
 *
 * This is a visual approximation, not a faithful spreadsheet-engine render:
 * formulas show their last-calculated value (if present in the source file),
 * charts and conditional formatting are not reproduced.
 */
export async function renderExcelWorkbookToPdf(
  workbook: XLSX.WorkBook,
  fileName: string,
  options: ExcelRenderOptions = {}
): Promise<{ pdfBuffer: Buffer; sheetCount: number }> {
  const pdfDoc = await PDFDocument.create();
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  const sheetNames = workbook.SheetNames || ["Sheet1"];
  const sheetCount = Math.max(1, sheetNames.length);

  const pageWidth = 841.89; // landscape A4
  const pageHeight = 595.28;
  const margin = 30;

  for (let sIdx = 0; sIdx < sheetCount; sIdx++) {
    const sheetName = sheetNames[sIdx];
    const sheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1 }) || [];

    const page = pdfDoc.addPage([pageWidth, pageHeight]);

    // Title bar
    page.drawRectangle({
      x: margin,
      y: pageHeight - margin - 32,
      width: pageWidth - margin * 2,
      height: 32,
      color: rgb(0.12, 0.45, 0.25),
    });
    page.drawText(sanitizeWinAnsi(`${fileName} > [ ${sheetName} ]`), {
      x: margin + 12,
      y: pageHeight - margin - 22,
      size: 11,
      font: boldFont,
      color: rgb(1, 1, 1),
    });
    const sheetBadge = `Sheet ${sIdx + 1} of ${sheetCount}`;
    const badgeWidth = font.widthOfTextAtSize(sheetBadge, 9);
    page.drawText(sheetBadge, {
      x: pageWidth - margin - badgeWidth - 12,
      y: pageHeight - margin - 20,
      size: 9,
      font: boldFont,
      color: rgb(0.85, 0.95, 0.88),
    });

    // Table dimensions
    const maxRows = Math.min(24, Math.max(1, rawData.length));
    let maxCols = 1;
    for (let r = 0; r < maxRows; r++) {
      if (Array.isArray(rawData[r])) {
        maxCols = Math.max(maxCols, Math.min(10, rawData[r].length));
      }
    }

    const rowHeaderWidth = 35;
    const tableWidth = pageWidth - margin * 2;
    const colWidth = Math.max(50, (tableWidth - rowHeaderWidth) / maxCols);
    const rowHeight = 18;
    const tableTop = pageHeight - margin - 40;
    const colHeaderY = tableTop - rowHeight;

    // Corner box
    page.drawRectangle({
      x: margin,
      y: colHeaderY,
      width: rowHeaderWidth,
      height: rowHeight,
      color: rgb(0.9, 0.92, 0.94),
      borderColor: rgb(0.78, 0.82, 0.86),
      borderWidth: 0.75,
    });

    // Column letters
    for (let c = 0; c < maxCols; c++) {
      const colLetter = String.fromCharCode(65 + c);
      const colX = margin + rowHeaderWidth + c * colWidth;
      page.drawRectangle({
        x: colX,
        y: colHeaderY,
        width: colWidth,
        height: rowHeight,
        color: rgb(0.93, 0.95, 0.96),
        borderColor: rgb(0.78, 0.82, 0.86),
        borderWidth: 0.75,
      });
      const letterWidth = boldFont.widthOfTextAtSize(colLetter, 8);
      page.drawText(colLetter, {
        x: colX + (colWidth - letterWidth) / 2,
        y: colHeaderY + 5,
        size: 8,
        font: boldFont,
        color: rgb(0.4, 0.45, 0.5),
      });
    }

    // Rows
    for (let r = 0; r < maxRows; r++) {
      const rowY = colHeaderY - (r + 1) * rowHeight;
      if (rowY < margin + 25) break;

      const isHeaderRow = r === 0;
      const isEvenRow = r % 2 === 0;

      page.drawRectangle({
        x: margin,
        y: rowY,
        width: rowHeaderWidth,
        height: rowHeight,
        color: rgb(0.93, 0.95, 0.96),
        borderColor: rgb(0.78, 0.82, 0.86),
        borderWidth: 0.75,
      });
      const rowNumStr = `${r + 1}`;
      const numWidth = font.widthOfTextAtSize(rowNumStr, 8);
      page.drawText(rowNumStr, {
        x: margin + (rowHeaderWidth - numWidth) / 2,
        y: rowY + 5,
        size: 8,
        font: boldFont,
        color: rgb(0.4, 0.45, 0.5),
      });

      const rowData = Array.isArray(rawData[r]) ? rawData[r] : [];
      for (let c = 0; c < maxCols; c++) {
        const colX = margin + rowHeaderWidth + c * colWidth;
        const cellValue = rowData[c] !== undefined && rowData[c] !== null ? String(rowData[c]) : "";
        const cellText = sanitizeWinAnsi(cellValue);

        const cellBg = isHeaderRow
          ? rgb(0.94, 0.97, 0.95)
          : isEvenRow
          ? rgb(1, 1, 1)
          : rgb(0.98, 0.98, 0.99);

        page.drawRectangle({
          x: colX,
          y: rowY,
          width: colWidth,
          height: rowHeight,
          color: cellBg,
          borderColor: rgb(0.85, 0.87, 0.9),
          borderWidth: 0.5,
        });

        if (cellText) {
          const currentFont = isHeaderRow ? boldFont : font;
          const fontSize = isHeaderRow ? 8.5 : 8;
          let displayStr = cellText;
          while (
            displayStr.length > 0 &&
            currentFont.widthOfTextAtSize(displayStr + "...", fontSize) > colWidth - 8
          ) {
            displayStr = displayStr.slice(0, -1);
          }
          if (displayStr !== cellText && displayStr.length > 0) {
            displayStr += "...";
          }

          page.drawText(displayStr, {
            x: colX + 4,
            y: rowY + 5,
            size: fontSize,
            font: currentFont,
            color: isHeaderRow ? rgb(0.1, 0.2, 0.15) : rgb(0.15, 0.15, 0.15),
          });
        }
      }
    }

    // Footer — only drawn if the caller supplied footerText
    const footerLabel = options.footerText
      ? `${options.footerText} - ${sheetName} - ${rawData.length} rows`
      : `${sheetName} - ${rawData.length} rows`;
    page.drawText(sanitizeWinAnsi(footerLabel), {
      x: margin,
      y: margin + 5,
      size: 8,
      font,
      color: rgb(0.5, 0.55, 0.6),
    });
  }

  const pdfBytes = await pdfDoc.save();
  return { pdfBuffer: Buffer.from(pdfBytes), sheetCount };
}
