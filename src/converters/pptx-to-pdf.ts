import JSZip from "jszip";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import { sanitizeWinAnsi } from "./sanitize";

export interface PptxRenderOptions {
  /** Text shown in the slide footer, e.g. "Rendered by Acme Inc". Omit for no footer text. */
  footerText?: string;
}

function extractBinaryStrings(buffer: Buffer): string[] {
  const raw = buffer.toString("latin1");
  const matches = raw.match(/[a-zA-Z0-9\s.,!?:;'"()\-_]{6,80}/g) || [];
  const ignoredWords = [
    "Root Entry",
    "Current User",
    "PowerPoint Document",
    "SummaryInformation",
    "DocumentSummaryInformation",
    "CompObj",
    "Times New Roman",
    "Calibri",
    "Arial",
  ];
  return matches.filter(
    (m) => !ignoredWords.some((w) => m.includes(w)) && /[a-zA-Z]{3,}/.test(m) && m.trim().length >= 6
  );
}

/**
 * Extracts slide text from PPTX/PPT and renders each slide as a simple,
 * modern card layout — a text-content approximation, not a faithful
 * rendering of the original slide design, images, or transitions.
 */
export async function pptxToPdf(
  pptxBuffer: Buffer,
  title?: string,
  options: PptxRenderOptions = {}
): Promise<Buffer> {
  let slideTextsArray: string[][] = [];

  try {
    const zip = await JSZip.loadAsync(pptxBuffer);
    const slideFiles = Object.keys(zip.files)
      .filter((f) => /^ppt\/slides\/slide\d+\.xml$/i.test(f))
      .sort((a, b) => {
        const numA = parseInt(a.match(/\d+/)![0], 10);
        const numB = parseInt(b.match(/\d+/)![0], 10);
        return numA - numB;
      });

    for (let i = 0; i < slideFiles.length; i++) {
      const slideXml = await zip.files[slideFiles[i]].async("text");
      const textMatches = Array.from(slideXml.matchAll(/<a:t[^>]*>(.*?)<\/a:t>/gi)).map((m) => m[1]);
      const slideTexts = textMatches.map((t) => sanitizeWinAnsi(t)).filter((t) => t.trim().length > 0);
      slideTextsArray.push(slideTexts);
    }
  } catch (_zipErr) {
    // Fallback for legacy binary .ppt (OLE2) formats
    const strings = extractBinaryStrings(pptxBuffer).map((s) => sanitizeWinAnsi(s));
    const maxSlides = 15;
    const chunkSize = Math.max(4, Math.ceil(strings.length / maxSlides));
    for (let i = 0; i < Math.min(strings.length, maxSlides * chunkSize); i += chunkSize) {
      const chunk = strings.slice(i, i + chunkSize);
      if (chunk.length > 0) slideTextsArray.push(chunk);
    }
  }

  const pdfDoc = await PDFDocument.create();
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  const slideWidth = 792; // 16:9
  const slideHeight = 445.5;
  const margin = 35;
  const totalSlides = Math.max(1, slideTextsArray.length);

  if (slideTextsArray.length === 0) {
    slideTextsArray.push([title || "Presentation Overview"]);
  }

  for (let i = 0; i < totalSlides; i++) {
    const slideLines = slideTextsArray[i] || [];
    const page = pdfDoc.addPage([slideWidth, slideHeight]);

    page.drawRectangle({ x: 0, y: 0, width: slideWidth, height: slideHeight, color: rgb(0.98, 0.98, 0.99) });
    page.drawRectangle({
      x: 0,
      y: slideHeight - 6,
      width: slideWidth,
      height: 6,
      color: rgb(0.18, 0.35, 0.95),
    });

    const badgeText = `SLIDE ${i + 1} OF ${totalSlides}`;
    page.drawRectangle({
      x: margin,
      y: slideHeight - margin - 20,
      width: 90,
      height: 18,
      color: rgb(0.9, 0.93, 1),
      borderColor: rgb(0.75, 0.83, 1),
      borderWidth: 0.75,
    });
    page.drawText(badgeText, {
      x: margin + 8,
      y: slideHeight - margin - 15,
      size: 7.5,
      font: boldFont,
      color: rgb(0.15, 0.35, 0.9),
    });

    const slideTitle = slideLines.length > 0 ? slideLines[0] : `Slide ${i + 1}`;
    const bodyLines = slideLines.length > 1 ? slideLines.slice(1) : [];

    page.drawText(sanitizeWinAnsi(slideTitle), {
      x: margin,
      y: slideHeight - margin - 48,
      size: 20,
      font: boldFont,
      color: rgb(0.08, 0.08, 0.12),
    });
    page.drawLine({
      start: { x: margin, y: slideHeight - margin - 56 },
      end: { x: slideWidth - margin, y: slideHeight - margin - 56 },
      thickness: 1,
      color: rgb(0.88, 0.9, 0.94),
    });

    const contentBoxY = slideHeight - margin - 68;
    const maxContentHeight = contentBoxY - margin - 35;

    if (bodyLines.length > 0) {
      let currentY = contentBoxY;
      const cardHeight = Math.min(50, Math.max(30, maxContentHeight / bodyLines.length - 8));

      for (let bIdx = 0; bIdx < Math.min(5, bodyLines.length); bIdx++) {
        const bulletText = bodyLines[bIdx];
        if (currentY - cardHeight < margin + 30) break;

        page.drawRectangle({
          x: margin,
          y: currentY - cardHeight,
          width: slideWidth - margin * 2,
          height: cardHeight,
          color: rgb(1, 1, 1),
          borderColor: rgb(0.88, 0.9, 0.94),
          borderWidth: 0.75,
        });
        page.drawRectangle({
          x: margin,
          y: currentY - cardHeight,
          width: 4,
          height: cardHeight,
          color: rgb(0.2, 0.45, 0.95),
        });
        page.drawText(`0${bIdx + 1}`, {
          x: margin + 14,
          y: currentY - cardHeight / 2 - 4,
          size: 10,
          font: boldFont,
          color: rgb(0.4, 0.5, 0.7),
        });

        const textY = currentY - cardHeight / 2 - 4;
        let lineStr = sanitizeWinAnsi(bulletText);
        if (font.widthOfTextAtSize(lineStr, 11) > slideWidth - margin * 2 - 60) {
          while (lineStr.length > 0 && font.widthOfTextAtSize(lineStr + "...", 11) > slideWidth - margin * 2 - 60) {
            lineStr = lineStr.slice(0, -1);
          }
          lineStr += "...";
        }
        page.drawText(lineStr, {
          x: margin + 40,
          y: textY,
          size: 11,
          font,
          color: rgb(0.15, 0.15, 0.2),
        });

        currentY -= cardHeight + 8;
      }
    } else {
      page.drawRectangle({
        x: margin,
        y: margin + 40,
        width: slideWidth - margin * 2,
        height: maxContentHeight,
        color: rgb(1, 1, 1),
        borderColor: rgb(0.88, 0.9, 0.94),
        borderWidth: 1,
      });
      page.drawText(sanitizeWinAnsi(title || "Presentation Slide Deck"), {
        x: margin + 30,
        y: margin + maxContentHeight / 2 + 10,
        size: 16,
        font: boldFont,
        color: rgb(0.12, 0.15, 0.25),
      });
    }

    if (options.footerText) {
      page.drawText(sanitizeWinAnsi(`${options.footerText} - ${title || "Presentation"}`), {
        x: margin,
        y: margin - 15,
        size: 8,
        font,
        color: rgb(0.5, 0.55, 0.6),
      });
    }

    const pageCountText = `${i + 1} / ${totalSlides}`;
    const pageNumWidth = font.widthOfTextAtSize(pageCountText, 8);
    page.drawText(pageCountText, {
      x: slideWidth - margin - pageNumWidth,
      y: margin - 15,
      size: 8,
      font: boldFont,
      color: rgb(0.4, 0.45, 0.5),
    });
  }

  const pdfBytes = await pdfDoc.save();
  return Buffer.from(pdfBytes);
}
