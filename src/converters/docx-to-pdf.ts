import * as mammoth from "mammoth";
import JSZip from "jszip";
import { textToPdf } from "./text-to-pdf";

const PAGE_BREAK_TOKEN = "__DOC2PNG_PAGE_BREAK__";

/**
 * Preprocesses DOCX XML to preserve page breaks, section breaks, and
 * page-break-before so rendered pages line up with the original document.
 */
async function preprocessDocxWithPageBreaks(buffer: Buffer): Promise<Buffer> {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const docXmlFile = zip.file("word/document.xml");
    if (!docXmlFile) return buffer;

    let docXml = await docXmlFile.async("text");

    // Explicit page breaks: <w:br w:type="page"/>
    docXml = docXml.replace(
      /<w:br[^>]*w:type=["']page["'][^>]*\/?>/gi,
      `<w:t>${PAGE_BREAK_TOKEN}</w:t>`
    );

    // Word-rendered page breaks: <w:lastRenderedPageBreak/>
    docXml = docXml.replace(
      /<w:lastRenderedPageBreak[^>]*\/?>/gi,
      `<w:t>${PAGE_BREAK_TOKEN}</w:t>`
    );

    // Paragraph page-break-before: <w:pageBreakBefore/>
    docXml = docXml.replace(
      /<w:pageBreakBefore[^>]*\/?>/gi,
      `</w:pPr><w:r><w:t>${PAGE_BREAK_TOKEN}</w:t></w:r><w:pPr>`
    );

    zip.file("word/document.xml", docXml);
    return (await zip.generateAsync({ type: "nodebuffer" })) as Buffer;
  } catch (_err) {
    return buffer;
  }
}

/**
 * Converts a DOCX/DOC buffer to styled A4 HTML with embedded images, tables,
 * headings, and page breaks matching the original Word document.
 */
export async function docxToHtml(docxBuffer: Buffer, title?: string): Promise<string> {
  try {
    const preprocessed = await preprocessDocxWithPageBreaks(docxBuffer);
    const result = await mammoth.convertToHtml(
      { buffer: preprocessed },
      {
        convertImage: mammoth.images.imgElement((image: any) => {
          return image.read("base64").then((imageBuffer: string) => {
            return { src: `data:${image.contentType};base64,${imageBuffer}` };
          });
        }),
      }
    );

    let htmlBody = result.value || `<p>${title || "Document"}</p>`;
    htmlBody = htmlBody.replace(
      new RegExp(`<p>\\s*${PAGE_BREAK_TOKEN}\\s*</p>|${PAGE_BREAK_TOKEN}`, "g"),
      '<div class="page-break"></div>'
    );

    return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
@page { size: 595.28pt 841.89pt; margin: 40pt 45pt 40pt 45pt; }
* { box-sizing: border-box; }
body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  font-size: 10.5pt;
  line-height: 1.55;
  color: #1a1a24;
  background-color: #ffffff;
  margin: 0;
  padding: 0;
}
h1 { font-size: 19pt; font-weight: 700; color: #0f172a; margin: 0 0 12pt; line-height: 1.3; }
h2 { font-size: 14pt; font-weight: 600; color: #1e293b; margin: 16pt 0 8pt; line-height: 1.35; }
h3 { font-size: 12pt; font-weight: 600; color: #334155; margin: 12pt 0 6pt; }
p { margin: 0 0 9pt; }
table { width: 100%; border-collapse: collapse; margin: 10pt 0 14pt; font-size: 9.5pt; }
th, td { border: 1px solid #cbd5e1; padding: 6pt 8pt; text-align: left; vertical-align: top; }
th { background-color: #f1f5f9; font-weight: 600; color: #0f172a; }
tr:nth-child(even) td { background-color: #f8fafc; }
ul, ol { margin: 0 0 9pt; padding-left: 18pt; }
li { margin-bottom: 3pt; }
img { max-width: 100%; height: auto; display: block; margin: 10pt 0; }
blockquote { margin: 10pt 0; padding-left: 12pt; border-left: 3pt solid #3b82f6; color: #475569; font-style: italic; }
.page-break { page-break-after: always; break-after: page; height: 0; margin: 0; padding: 0; }
</style>
</head>
<body>
${htmlBody}
</body>
</html>`;
  } catch (_err) {
    // Fallback for legacy binary .doc formats
    const raw = docxBuffer.toString("latin1");
    const matches = raw.match(/[\x20-\x7E\xA0-\xFF]{4,}/g) || [];
    const filtered = matches.filter(
      (m) =>
        !/^(Root Entry|WordDocument|SummaryInformation|DocumentSummaryInformation|CompObj|Table)/i.test(m) &&
        m.trim().length > 3
    );
    const paragraphs = filtered.map((p) => `<p>${p}</p>`).join("\n");
    return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>@page{size:595.28pt 841.89pt;margin:40pt;}body{font-family:sans-serif;font-size:10.5pt;line-height:1.5;}</style></head><body><h1>${title || "Document"}</h1>${paragraphs}</body></html>`;
  }
}

/**
 * Converts a DOCX/DOC buffer to a plain-text-based PDF (fallback path, no
 * layout fidelity — used only if HTML rendering is unavailable).
 */
export async function docxToPdf(docxBuffer: Buffer, title?: string): Promise<Buffer> {
  try {
    const result = await mammoth.extractRawText({ buffer: docxBuffer });
    const rawText = result.value || "Empty Document";
    return textToPdf(rawText, title);
  } catch (_err) {
    const raw = docxBuffer.toString("latin1");
    const matches = raw.match(/[\x20-\x7E]{4,}/g) || [];
    const filtered = matches.filter(
      (m) =>
        !/^(Root Entry|WordDocument|SummaryInformation|DocumentSummaryInformation|CompObj|Table)/i.test(m) &&
        m.trim().length > 3
    );
    const textContent = filtered.join("\n") || "Document";
    return textToPdf(textContent, title);
  }
}
