/* Printable QR stickers (name + serial) for the tools in the workbook.
 * A sticker on the tool makes recognition exact without any OCR or API:
 * the QR holds "TOOL:<serial>". Pages are drawn on a canvas (so Korean
 * names use the page's web font) and placed into an A4 PDF with jsPDF.
 */
(function () {
  "use strict";
  const TC = (window.TC = window.TC || {});

  const PX_PER_MM = 12; // ≈ 305 dpi
  const PAGE = { w: 210, h: 297 };
  const STICKER = { w: 48, h: 24 };
  const MARGIN = { x: 9, y: 16 };
  const COLS = 4;
  const ROWS = 11;

  function drawQr(ctx, text, x, y, size) {
    const qr = qrcode(0, "M");
    qr.addData(text);
    qr.make();
    const n = qr.getModuleCount();
    const quiet = 2;
    const cell = size / (n + quiet * 2);
    ctx.fillStyle = "#fff";
    ctx.fillRect(x, y, size, size);
    ctx.fillStyle = "#000";
    for (let r = 0; r < n; r++) {
      for (let c = 0; c < n; c++) {
        if (qr.isDark(r, c)) {
          ctx.fillRect(Math.floor(x + (c + quiet) * cell), Math.floor(y + (r + quiet) * cell), Math.ceil(cell), Math.ceil(cell));
        }
      }
    }
  }

  function fitText(ctx, text, maxWidth, size, weight, family) {
    let s = size;
    do {
      ctx.font = `${weight} ${s}px ${family}`;
      if (ctx.measureText(text).width <= maxWidth) break;
      s -= 2;
    } while (s > 8);
    return s;
  }

  /** Draw one sticker with its top-left corner at (x, y) in pixels. */
  function drawSticker(ctx, rec, x, y, k) {
    const w = STICKER.w * k;
    const h = STICKER.h * k;
    ctx.save();
    ctx.strokeStyle = "#b8bcb6";
    ctx.lineWidth = Math.max(1, k * 0.15);
    ctx.setLineDash([k * 1.2, k * 1.2]);
    ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
    ctx.setLineDash([]);
    const pad = 1.6 * k;
    const qrSize = h - 2 * pad;
    drawQr(ctx, TC.vision.QR_PREFIX + rec.serial, x + pad, y + pad, qrSize);
    const tx = x + pad + qrSize + 1.4 * k;
    const tw = x + w - pad - tx;
    const sans = '"IBM Plex Sans KR", "Noto Sans KR", "Malgun Gothic", sans-serif';
    const mono = '"IBM Plex Mono", ui-monospace, Menlo, Consolas, monospace';
    ctx.fillStyle = "#111";
    ctx.textBaseline = "alphabetic";
    // shared tools have no person in 'Asset #': show what follows the region
    const name = rec.names[0] || rec.asset.replace(/^[^_]*_[^_]*_/, "").replace(/_/g, " ") || rec.serial;
    const nameSize = fitText(ctx, name, tw, 5.2 * k, 700, sans);
    ctx.fillText(name, tx, y + pad + nameSize * 0.9);
    const serialSize = fitText(ctx, rec.serial, tw, 3.6 * k, 600, mono);
    ctx.fillText(rec.serial, tx, y + h * 0.62);
    ctx.fillStyle = "#444";
    const meta = `${rec.sheet} · ${[rec.manufacturer, rec.model].filter(Boolean).join(" ")}`;
    fitText(ctx, meta, tw, 2.4 * k, 400, sans);
    ctx.fillText(meta, tx, y + h - pad);
    ctx.restore();
    return serialSize;
  }

  function renderPage(records) {
    const k = PX_PER_MM;
    const c = TC.vision.makeCanvas(PAGE.w * k, PAGE.h * k);
    const ctx = c.getContext("2d");
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, c.width, c.height);
    records.forEach((rec, i) => {
      const col = i % COLS;
      const row = Math.floor(i / COLS);
      drawSticker(ctx, rec, (MARGIN.x + col * STICKER.w) * k, (MARGIN.y + row * STICKER.h) * k, k);
    });
    return c;
  }

  async function ensureFonts() {
    if (!document.fonts) return;
    await Promise.allSettled([
      document.fonts.load('700 40px "IBM Plex Sans KR"', "가나다"),
      document.fonts.load('400 30px "IBM Plex Sans KR"', "가나다"),
      document.fonts.load('600 30px "IBM Plex Mono"', "0123"),
    ]);
  }

  async function buildPdf(records) {
    await ensureFonts();
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait", compress: true });
    const perPage = COLS * ROWS;
    for (let p = 0; p * perPage < records.length; p++) {
      if (p) doc.addPage();
      const page = renderPage(records.slice(p * perPage, (p + 1) * perPage));
      doc.addImage(page.toDataURL("image/png"), "PNG", 0, 0, PAGE.w, PAGE.h, undefined, "FAST");
    }
    return doc.output("blob");
  }

  /** A single sticker for on-screen preview. */
  async function previewCanvas(rec) {
    await ensureFonts();
    const k = 8;
    const c = TC.vision.makeCanvas(STICKER.w * k, STICKER.h * k);
    const ctx = c.getContext("2d");
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, c.width, c.height);
    drawSticker(ctx, rec, 0, 0, k);
    return c;
  }

  TC.stickers = { buildPdf, previewCanvas, drawSticker, PER_PAGE: COLS * ROWS, STICKER };
})();
