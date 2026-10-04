import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import { sanitizeWinAnsi } from "./sanitize";

/**
 * Converts raw plain text or CSV into a paginated PDF buffer.
 */
export async function textToPdf(textContent: string, title?: string): Promise<Buffer> {
  const pdfDoc = await PDFDocument.create();
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  const pageWidth = 595.28; // A4
  const pageHeight = 841.89;
  const margin = 45;
  const contentWidth = pageWidth - margin * 2;
  const fontSize = 10;
  const lineHeight = 14;

  const rawLines = textContent.split(/\r?\n/).map((l) => sanitizeWinAnsi(l));
  const wrappedLines: string[] = [];

  for (const rawLine of rawLines) {
    if (!rawLine.trim()) {
      wrappedLines.push("");
      continue;
    }

    const words = rawLine.split(" ");
    let currentLine = "";

    for (const word of words) {
      const testLine = currentLine ? `${currentLine} ${word}` : word;
      const testWidth = font.widthOfTextAtSize(testLine, fontSize);
      if (testWidth > contentWidth && currentLine) {
        wrappedLines.push(currentLine);
        currentLine = word;
      } else {
        currentLine = testLine;
      }
    }
    if (currentLine) {
      wrappedLines.push(currentLine);
    }
  }

  const linesPerPage = Math.floor((pageHeight - margin * 2 - 30) / lineHeight);
  const totalPages = Math.max(1, Math.ceil(wrappedLines.length / linesPerPage));

  for (let p = 0; p < totalPages; p++) {
    const page = pdfDoc.addPage([pageWidth, pageHeight]);
    let y = pageHeight - margin;

    if (title && p === 0) {
      page.drawText(sanitizeWinAnsi(title), {
        x: margin,
        y,
        size: 14,
        font: boldFont,
        color: rgb(0.1, 0.1, 0.1),
      });
      y -= 24;
    }

    const pageLines = wrappedLines.slice(p * linesPerPage, (p + 1) * linesPerPage);
    for (const line of pageLines) {
      if (line) {
        page.drawText(line, {
          x: margin,
          y,
          size: fontSize,
          font,
          color: rgb(0.15, 0.15, 0.15),
        });
      }
      y -= lineHeight;
    }

    const footerText = `Page ${p + 1} of ${totalPages}`;
    const footerWidth = font.widthOfTextAtSize(footerText, 8);
    page.drawText(footerText, {
      x: pageWidth - margin - footerWidth,
      y: margin / 2,
      size: 8,
      font,
      color: rgb(0.5, 0.5, 0.5),
    });
  }

  const pdfBytes = await pdfDoc.save();
  return Buffer.from(pdfBytes);
}
