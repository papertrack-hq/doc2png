# doc2png

Render any document — PDF, DOCX/DOC, PPTX/PPT, XLSX/XLS, TXT/CSV, or an image
(PNG/JPG/WEBP) — into an array of page-by-page PNG images. No storage,
database, or cloud-provider dependency: you get buffers back, you decide
where they go.

Built by the team at [PaperTrack](https://papertrack.io) — a secure document
sharing and data-room tool — because we needed this for every document a
user uploads, and couldn't find a single package that handled all these
formats consistently.

## Install

```bash
npm install doc2png
```

## Usage

```ts
import { renderDocumentToPages } from "doc2png";
import { readFileSync } from "fs";

const buffer = readFileSync("./quarterly-report.pptx");
const { totalPages, pages } = await renderDocumentToPages(buffer, "quarterly-report.pptx");

for (const page of pages) {
  // page.pageNumber, page.png (a Buffer of PNG bytes)
  writeFileSync(`page-${page.pageNumber}.png`, page.png);
}
```

### Options

```ts
await renderDocumentToPages(buffer, fileName, {
  dpi: 150,                 // render resolution, default 150
  footerText: "Acme Inc",   // optional branding stamped on Excel/PPTX preview pages
});
```

## Supported formats

| Format | How it's rendered |
|---|---|
| PDF | Rasterized directly via MuPDF |
| DOCX / DOC | Converted to styled HTML (via `mammoth`), then rasterized. Page breaks, headings, tables, and inline images are preserved. Legacy binary `.doc` falls back to best-effort text extraction. |
| PPTX / PPT | Slide **text content** is extracted and re-laid-out into a clean card-style slide. This is **not** a faithful render of your original slide design — images, custom layouts, and transitions are not reproduced. Legacy binary `.ppt` falls back to best-effort text extraction. |
| XLSX / XLS | Rendered as a grid preview (column letters, row numbers, zebra striping) — one page per sheet, first 24 rows and 10 columns. This is a **visual approximation**, not a faithful spreadsheet-engine render: formulas show their last-calculated value if present in the file, and charts/conditional formatting are not reproduced. |
| PNG / JPG / WEBP | Passed through unchanged as a single page. |
| TXT / CSV | Paginated into a simple PDF, then rasterized. |

## ⚠️ Licensing: MuPDF is AGPL / commercial dual-licensed

This package uses [`mupdf`](https://www.npmjs.com/package/mupdf) (Artifex's
official Node bindings) to rasterize PDFs. MuPDF is distributed under the
**GNU AGPL v3**, with a paid commercial license available from Artifex for
anyone who doesn't want AGPL obligations.

**What this means for you:** if you ship a product that uses this package
(directly or as a dependency) and your product is offered as a network
service, AGPL's terms may require you to make your product's source
available to its users — unless you hold a commercial MuPDF license from
Artifex, or you aren't triggering AGPL's conditions in the first place.
This is a real legal question, not a formality — talk to a lawyer if you're
unsure, and see [mupdf.com/licensing](https://mupdf.com/licensing/) for
Artifex's own explanation.

If AGPL is a blocker for your use case, you'll need to swap the rasterization
step for an alternative PDF renderer with a license that fits your project —
`doc2png`'s MuPDF usage is isolated to `rasterizeMupdfDoc` in `src/index.ts`
and `renderExcelWorkbookToPdf`'s caller, so that swap is contained to a
couple of call sites.

## What this package deliberately does NOT do

- No file storage, blob upload, or database writes — you get `Buffer[]` back and own what happens next.
- No authentication, access control, or multi-tenant document ID handling — that's your application's job.
- No watermarking, analytics, or tracking — this package only rasterizes pages.

## License

MIT for this package's own code. See the MuPDF section above for the
rasterization dependency's license terms, which are separate from and not
overridden by this package's MIT license.
