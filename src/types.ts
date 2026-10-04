export interface RenderedPage {
  pageNumber: number;
  /** Raw PNG bytes for this page. */
  png: Buffer;
}

export interface RenderOptions {
  /**
   * Render resolution in DPI. Higher = sharper but larger PNGs.
   * @default 150
   */
  dpi?: number;
  /**
   * Optional footer text stamped onto rendered Excel/PPTX preview pages
   * (e.g. your product name). Left blank, no footer branding is added.
   */
  footerText?: string;
}

export interface RenderResult {
  totalPages: number;
  pages: RenderedPage[];
}
