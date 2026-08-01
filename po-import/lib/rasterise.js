'use strict';

const { createCanvas } = require('canvas');

/**
 * Pinned to pdfjs-dist v3 deliberately. v4 cannot render this PO under node-canvas 3.x: it
 * throws "Image or Canvas expected" from paintInlineImageXObject on the first embedded
 * image (the supplier's letterhead JPEG), because it expects browser image primitives that
 * node-canvas 3 no longer matches. Shimming HTMLElement/ImageData/OffscreenCanvas does not
 * fix it — the mismatch is deeper than the globals. v3 renders the same file cleanly.
 *
 * If pdfjs is ever upgraded here, re-run the fixture test before trusting it.
 */

/**
 * Render each PDF page to a PNG buffer.
 *
 * This step exists because of a specific defect in the supplier's PDFs: they carry a text
 * layer whose glyph codes have no /ToUnicode mapping, so the characters are unrecoverable
 * (a raw read yields "aXXNMM]"). Worse, Drive sees that layer, decides the file already has
 * text, and skips OCR entirely — returning the same garbage no matter what ocr flag is set.
 *
 * Rendering to pixels discards the poisoned layer, leaving Drive no option but to actually
 * read the page. Verified against PO AMB/26-27/PAC/00450: 0/6 key fields via direct upload,
 * 17/17 after rasterising.
 *
 * 300 DPI is the floor for the ~6pt digits in the rate column; below roughly 200 the
 * decimal points start dropping.
 */

const DEFAULT_DPI = 300;

async function rasterisePdf(pdfBuffer, { dpi = DEFAULT_DPI, onPage } = {}) {
  const pdfjs = require('pdfjs-dist/legacy/build/pdf.js');

  const doc = await pdfjs.getDocument({
    data: new Uint8Array(pdfBuffer),
    // The fonts are broken and unused once rasterised; skip them rather than warn per glyph.
    disableFontFace: true,
    isEvalSupported: false,
    verbosity: 0,
  }).promise;

  const scale = dpi / 72; // PDF user units are 1/72 inch
  const pages = [];

  for (let n = 1; n <= doc.numPages; n++) {
    const page = await doc.getPage(n);
    const viewport = page.getViewport({ scale });
    const canvas = createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
    const ctx = canvas.getContext('2d');

    // White ground: the canvas starts transparent, which OCR reads as noise.
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    await page.render({ canvasContext: ctx, viewport }).promise;

    const png = canvas.toBuffer('image/png');
    pages.push({ page: n, width: canvas.width, height: canvas.height, png });
    if (onPage) onPage(n, doc.numPages, png.length);
  }

  await doc.destroy();
  return pages;
}

module.exports = { rasterisePdf, DEFAULT_DPI };
