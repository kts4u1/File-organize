/* Image analysis, all in the browser (no server, no API):
 *  - loading photos (EXIF orientation applied by the browser)
 *  - QR sticker decoding (jsQR)
 *  - calibration label (교정필증) detection, crop and deskew
 *  - finding handwritten lines on the sticky note
 *  - OCR with Tesseract.js: label dates, serial/name lines, SMAX screenshots
 */
(function () {
  "use strict";
  const TC = (window.TC = window.TC || {});
  const { parseDate, todayIso, daysBetween } = TC.xlsx;

  const MAX_SIDE = 3000; // originals are kept at most this large
  const WORK_SIDE = 1200; // detection runs on a copy this large
  const QR_PREFIX = "TOOL:";

  // ------------------------------------------------------------ canvas utils
  function makeCanvas(w, h) {
    const c = document.createElement("canvas");
    c.width = Math.max(1, Math.round(w));
    c.height = Math.max(1, Math.round(h));
    return c;
  }
  function scaled(src, maxSide) {
    const s = Math.min(1, maxSide / Math.max(src.width, src.height));
    if (s === 1) return { canvas: src, scale: 1 };
    const c = makeCanvas(src.width * s, src.height * s);
    const ctx = c.getContext("2d");
    ctx.imageSmoothingQuality = "high";
    ctx.drawImage(src, 0, 0, c.width, c.height);
    return { canvas: c, scale: c.width / src.width };
  }
  const pixels = (c) => c.getContext("2d", { willReadFrequently: true }).getImageData(0, 0, c.width, c.height);

  async function loadImage(file) {
    let bmp;
    try {
      bmp = await createImageBitmap(file, { imageOrientation: "from-image" });
    } catch (e) {
      const heic = /\.(heic|heif)$/i.test(file.name || "");
      throw new Error(heic
        ? "HEIC 사진은 이 브라우저에서 열 수 없습니다. 아이폰 설정 > 카메라 > 포맷 > '높은 호환성'으로 찍거나 JPG로 보내 주세요."
        : "이미지를 열 수 없습니다.");
    }
    const s = Math.min(1, MAX_SIDE / Math.max(bmp.width, bmp.height));
    const c = makeCanvas(bmp.width * s, bmp.height * s);
    const ctx = c.getContext("2d");
    ctx.imageSmoothingQuality = "high";
    ctx.drawImage(bmp, 0, 0, c.width, c.height);
    bmp.close?.();
    return c;
  }

  function rotateCanvas(src, degreesClockwise) {
    const d = ((degreesClockwise % 360) + 360) % 360;
    if (!d) return src;
    const swap = d === 90 || d === 270;
    const c = makeCanvas(swap ? src.height : src.width, swap ? src.width : src.height);
    const ctx = c.getContext("2d");
    ctx.translate(c.width / 2, c.height / 2);
    ctx.rotate((d * Math.PI) / 180);
    ctx.drawImage(src, -src.width / 2, -src.height / 2);
    return c;
  }

  function percentile(values, p) {
    const hist = new Uint32Array(256);
    for (let i = 0; i < values.length; i++) hist[values[i]]++;
    const target = values.length * p;
    let acc = 0;
    for (let v = 0; v < 256; v++) {
      acc += hist[v];
      if (acc >= target) return v;
    }
    return 255;
  }

  // Binary morphology with a w×h rectangle, separable, via running sums.
  function runFilter(mask, W, H, rx, ry, mode) {
    const tmp = new Uint8Array(W * H);
    const out = new Uint8Array(W * H);
    const pass = (src, dst, len, count, stride, step, r) => {
      for (let line = 0; line < count; line++) {
        const base = line * stride;
        let sum = 0;
        const at = (i) => src[base + i * step];
        for (let i = 0; i <= Math.min(r, len - 1); i++) sum += at(i);
        for (let i = 0; i < len; i++) {
          const lo = Math.max(0, i - r);
          const hi = Math.min(len - 1, i + r);
          const width = hi - lo + 1;
          dst[base + i * step] = mode === "dilate" ? (sum > 0 ? 1 : 0) : sum === width ? 1 : 0;
          if (i + r + 1 < len) sum += at(i + r + 1);
          if (i - r >= 0) sum -= at(i - r);
        }
      }
    };
    pass(mask, tmp, W, H, W, 1, rx); // rows
    pass(tmp, out, H, W, 1, W, ry); // columns
    return out;
  }
  const dilate = (m, W, H, rx, ry) => runFilter(m, W, H, rx, ry, "dilate");
  const erode = (m, W, H, rx, ry) => runFilter(m, W, H, rx, ry, "erode");

  function components(mask, W, H) {
    const labels = new Int32Array(W * H);
    const comps = [];
    const stack = new Int32Array(W * H);
    for (let start = 0; start < W * H; start++) {
      if (!mask[start] || labels[start]) continue;
      const id = comps.length + 1;
      let sp = 0;
      stack[sp++] = start;
      labels[start] = id;
      const c = { id, n: 0, x0: W, y0: H, x1: 0, y1: 0, sx: 0, sy: 0, sxx: 0, syy: 0, sxy: 0 };
      while (sp) {
        const p = stack[--sp];
        const x = p % W;
        const y = (p / W) | 0;
        c.n++;
        c.sx += x; c.sy += y; c.sxx += x * x; c.syy += y * y; c.sxy += x * y;
        if (x < c.x0) c.x0 = x;
        if (x > c.x1) c.x1 = x;
        if (y < c.y0) c.y0 = y;
        if (y > c.y1) c.y1 = y;
        const nb = [p - 1, p + 1, p - W, p + W];
        if (x === 0) nb[0] = -1;
        if (x === W - 1) nb[1] = -1;
        for (const q of nb) {
          if (q >= 0 && q < W * H && mask[q] && !labels[q]) {
            labels[q] = id;
            stack[sp++] = q;
          }
        }
      }
      comps.push(c);
    }
    return { labels, comps };
  }

  // ------------------------------------------------------------ label detection
  /** Returns {cx, cy, long, short, angle, box:{left,top,right,bottom}} in original pixels, or null. */
  function findLabel(src, hint) {
    const { canvas, scale } = scaled(src, WORK_SIDE);
    const W = canvas.width;
    const H = canvas.height;
    const data = pixels(canvas).data;
    const value = new Uint8Array(W * H);
    const chroma = new Uint8Array(W * H);
    for (let i = 0, j = 0; j < W * H; i += 4, j++) {
      const r = data[i], g = data[i + 1], b = data[i + 2];
      const mx = Math.max(r, g, b);
      value[j] = mx;
      chroma[j] = mx - Math.min(r, g, b);
    }
    const bright = percentile(value, 0.995);
    let mask = new Uint8Array(W * H);
    for (let j = 0; j < W * H; j++) mask[j] = chroma[j] < 0.09 * value[j] + 4 && value[j] > 0.55 * bright ? 1 : 0;
    const k = Math.max(3, Math.floor(Math.min(W, H) * 0.012)) | 1;
    mask = dilate(erode(mask, W, H, 1, 1), W, H, 1, 1); // open 3x3
    mask = erode(dilate(mask, W, H, k >> 1, k >> 1), W, H, k >> 1, k >> 1); // close k×k
    const { labels, comps } = components(mask, W, H);

    const cands = [];
    for (const c of comps) {
      const frac = c.n / (W * H);
      if (frac < 0.004 || frac > 0.6) continue;
      // fill holes (text inside the label) by flooding the outside within the bbox
      const bw = c.x1 - c.x0 + 1;
      const bh = c.y1 - c.y0 + 1;
      const seen = new Uint8Array(bw * bh);
      const st = [];
      const push = (x, y) => {
        const li = (y - c.y0) * bw + (x - c.x0);
        if (seen[li] || labels[y * W + x] === c.id) return;
        seen[li] = 1;
        st.push(li);
      };
      for (let x = c.x0; x <= c.x1; x++) { push(x, c.y0); push(x, c.y1); }
      for (let y = c.y0; y <= c.y1; y++) { push(c.x0, y); push(c.x1, y); }
      let outside = 0;
      while (st.length) {
        const li = st.pop();
        outside++;
        const x = (li % bw) + c.x0;
        const y = ((li / bw) | 0) + c.y0;
        if (x > c.x0) push(x - 1, y);
        if (x < c.x1) push(x + 1, y);
        if (y > c.y0) push(x, y - 1);
        if (y < c.y1) push(x, y + 1);
      }
      const filled = bw * bh - outside;
      // orientation from second moments
      const mx = c.sx / c.n;
      const my = c.sy / c.n;
      const cxx = c.sxx / c.n - mx * mx;
      const cyy = c.syy / c.n - my * my;
      const cxy = c.sxy / c.n - mx * my;
      const theta = 0.5 * Math.atan2(2 * cxy, cxx - cyy);
      const cos = Math.cos(theta);
      const sin = Math.sin(theta);
      let umin = Infinity, umax = -Infinity, vmin = Infinity, vmax = -Infinity;
      for (let y = c.y0; y <= c.y1; y++) {
        for (let x = c.x0; x <= c.x1; x++) {
          if (labels[y * W + x] !== c.id) continue;
          const u = (x - mx) * cos + (y - my) * sin;
          const v = -(x - mx) * sin + (y - my) * cos;
          if (u < umin) umin = u;
          if (u > umax) umax = u;
          if (v < vmin) vmin = v;
          if (v > vmax) vmax = v;
        }
      }
      const L = umax - umin + 1;
      const S = vmax - vmin + 1;
      const rectangularity = filled / (L * S);
      const longSide = Math.max(L, S);
      const shortSide = Math.min(L, S);
      const aspect = longSide / shortSide;
      if (rectangularity < 0.72 || aspect < 1.05 || aspect > 2.8) continue;
      let angle = (theta * 180) / Math.PI + (L >= S ? 0 : 90);
      while (angle > 90) angle -= 180;
      while (angle <= -90) angle += 180;
      const cu = (umin + umax) / 2;
      const cv = (vmin + vmax) / 2;
      const ccx = mx + cu * cos - cv * sin;
      const ccy = my + cu * sin + cv * cos;
      const aspectPref = aspect >= 1.2 && aspect <= 2.0 ? 1 : 0.7;
      cands.push({
        cx: ccx / scale, cy: ccy / scale, long: longSide / scale, short: shortSide / scale, angle,
        box: { left: c.x0 / scale, top: c.y0 / scale, right: (c.x1 + 1) / scale, bottom: (c.y1 + 1) / scale },
        score: Math.min(rectangularity, 1) * aspectPref * Math.pow(frac, 0.25),
      });
    }
    if (hint) for (const c of cands) c.score *= 0.3 + iou(c.box, hint);
    cands.sort((a, b) => b.score - a.score);
    return cands[0] || null;
  }

  function iou(a, b) {
    const ix = Math.max(0, Math.min(a.right, b.right) - Math.max(a.left, b.left));
    const iy = Math.max(0, Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top));
    const inter = ix * iy;
    const union = (a.right - a.left) * (a.bottom - a.top) + (b.right - b.left) * (b.bottom - b.top) - inter;
    return union > 0 ? inter / union : 0;
  }

  /** Straightened crop of the label with a margin (pad as a fraction of each side). */
  function cropLabel(src, label, pad = 0.12) {
    const straight = Math.abs(label.angle) <= 20 ? label.angle : 0;
    const landscape = Math.abs(label.angle) <= 45;
    const w = (landscape ? label.long : label.short) * (1 + 2 * pad);
    const h = (landscape ? label.short : label.long) * (1 + 2 * pad);
    const c = makeCanvas(w, h);
    const ctx = c.getContext("2d");
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, c.width, c.height);
    ctx.translate(c.width / 2, c.height / 2);
    ctx.rotate(landscape ? (-straight * Math.PI) / 180 : 0);
    ctx.translate(-label.cx, -label.cy);
    ctx.imageSmoothingQuality = "high";
    ctx.drawImage(src, 0, 0);
    return c;
  }

  // ------------------------------------------------------------ handwriting lines
  /** Boxes (original pixels) of dark ink text lines outside the label. */
  function findInkLines(src, labelBox) {
    const { canvas, scale } = scaled(src, WORK_SIDE);
    const W = canvas.width;
    const H = canvas.height;
    const data = pixels(canvas).data;
    const gray = new Uint8Array(W * H);
    for (let i = 0, j = 0; j < W * H; i += 4, j++) gray[j] = (0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2]) | 0;
    // local mean via integral image
    const integ = new Float64Array((W + 1) * (H + 1));
    for (let y = 0; y < H; y++) {
      let row = 0;
      for (let x = 0; x < W; x++) {
        row += gray[y * W + x];
        integ[(y + 1) * (W + 1) + x + 1] = integ[y * (W + 1) + x + 1] + row;
      }
    }
    const r = Math.max(8, Math.round(Math.min(W, H) * 0.03));
    const ink = new Uint8Array(W * H);
    const ex = labelBox ? {
      l: (labelBox.left * scale) - 6, t: (labelBox.top * scale) - 6, r: (labelBox.right * scale) + 6, b: (labelBox.bottom * scale) + 6,
    } : null;
    for (let y = 0; y < H; y++) {
      const y0 = Math.max(0, y - r), y1 = Math.min(H, y + r + 1);
      for (let x = 0; x < W; x++) {
        if (ex && x >= ex.l && x <= ex.r && y >= ex.t && y <= ex.b) continue;
        const x0 = Math.max(0, x - r), x1 = Math.min(W, x + r + 1);
        const sum = integ[y1 * (W + 1) + x1] - integ[y0 * (W + 1) + x1] - integ[y1 * (W + 1) + x0] + integ[y0 * (W + 1) + x0];
        const mean = sum / ((x1 - x0) * (y1 - y0));
        const g = gray[y * W + x];
        if (g < 125 && mean - g > 40) ink[y * W + x] = 1;
      }
    }
    const rx = Math.max(4, Math.round(W * 0.012));
    const merged = dilate(ink, W, H, rx, 2);
    const { comps, labels } = components(merged, W, H);
    const lines = [];
    for (const c of comps) {
      const bw = c.x1 - c.x0 + 1;
      const bh = c.y1 - c.y0 + 1;
      if (bh < H * 0.012 || bh > H * 0.16 || bw < bh * 1.8 || bw > W * 0.95) continue;
      let inkCount = 0;
      for (let y = c.y0; y <= c.y1; y++) for (let x = c.x0; x <= c.x1; x++) if (labels[y * W + x] === c.id && ink[y * W + x]) inkCount++;
      const density = inkCount / (bw * bh);
      if (density < 0.03 || density > 0.5) continue;
      lines.push({ left: c.x0 / scale, top: c.y0 / scale, right: (c.x1 + 1) / scale, bottom: (c.y1 + 1) / scale, weight: inkCount });
    }
    lines.sort((a, b) => b.weight - a.weight);
    return lines.slice(0, 8).sort((a, b) => a.top - b.top);
  }

  /** Grayscale, autocontrast, resized so the text is ~targetH px tall. */
  function prepareForOcr(src, box, targetH = 90, padRatio = 0.25) {
    const bh = box.bottom - box.top;
    const pad = bh * padRatio;
    const sx = Math.max(0, box.left - pad), sy = Math.max(0, box.top - pad);
    const sw = Math.min(src.width, box.right + pad) - sx;
    const sh = Math.min(src.height, box.bottom + pad) - sy;
    const k = Math.min(4, Math.max(1, targetH / bh));
    const c = makeCanvas(sw * k, sh * k);
    const ctx = c.getContext("2d", { willReadFrequently: true });
    ctx.imageSmoothingQuality = "high";
    ctx.drawImage(src, sx, sy, sw, sh, 0, 0, c.width, c.height);
    autocontrastGray(c);
    return c;
  }
  function autocontrastGray(c) {
    const ctx = c.getContext("2d", { willReadFrequently: true });
    const img = ctx.getImageData(0, 0, c.width, c.height);
    const d = img.data;
    const g = new Uint8Array(c.width * c.height);
    for (let i = 0, j = 0; j < g.length; i += 4, j++) g[j] = (0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]) | 0;
    const lo = percentile(g, 0.02);
    const hi = Math.max(lo + 1, percentile(g, 0.98));
    for (let i = 0, j = 0; j < g.length; i += 4, j++) {
      const v = Math.max(0, Math.min(255, ((g[j] - lo) * 255) / (hi - lo)));
      d[i] = d[i + 1] = d[i + 2] = v;
    }
    ctx.putImageData(img, 0, 0);
    return c;
  }

  const DIGITS = "0123456789";
  const LINE_VARIANTS = [["raw", 70], ["raw", 110], ["gray", 70], ["gray", 110], ["value", 90]];
  const cleanReading = (t) => String(t || "").toUpperCase().replace(/[^0-9A-Z]/g, "");

  /** Crop a text line with margin, scaled to `height` px of text, in one of three renderings. */
  function lineCrop(src, box, mode, height) {
    const bh = box.bottom - box.top;
    const pad = bh * 0.25;
    const sx = Math.max(0, box.left - pad), sy = Math.max(0, box.top - pad);
    const sw = Math.min(src.width, box.right + pad) - sx;
    const sh = Math.min(src.height, box.bottom + pad) - sy;
    const k = Math.min(5, Math.max(0.5, height / bh));
    const c = makeCanvas(sw * k, sh * k);
    const ctx = c.getContext("2d", { willReadFrequently: true });
    ctx.imageSmoothingQuality = "high";
    ctx.drawImage(src, sx, sy, sw, sh, 0, 0, c.width, c.height);
    if (mode === "gray") autocontrastGray(c);
    if (mode === "value") {
      // brightest channel: paper of any colour turns light, ink stays dark
      const img = ctx.getImageData(0, 0, c.width, c.height);
      const d = img.data;
      for (let i = 0; i < d.length; i += 4) d[i] = d[i + 1] = d[i + 2] = Math.max(d[i], d[i + 1], d[i + 2]);
      ctx.putImageData(img, 0, 0);
      autocontrastGray(c);
    }
    return c;
  }

  // ------------------------------------------------------------ QR
  function readQr(src) {
    if (typeof jsQR !== "function") return null;
    for (const side of [1200, 2000, MAX_SIDE]) {
      const { canvas } = scaled(src, side);
      const img = pixels(canvas);
      const code = jsQR(img.data, img.width, img.height, { inversionAttempts: "attemptBoth" });
      if (code && code.data) {
        const text = code.data.trim();
        return text.startsWith(QR_PREFIX) ? text.slice(QR_PREFIX.length).split("|")[0] : text;
      }
      if (Math.max(src.width, src.height) <= side) break;
    }
    return null;
  }

  // ------------------------------------------------------------ screenshots
  function looksLikeScreenshot(src) {
    const { canvas } = scaled(src, 600);
    const d = pixels(canvas).data;
    let white = 0, saturated = 0;
    const n = d.length / 4;
    for (let i = 0; i < d.length; i += 4) {
      const mx = Math.max(d[i], d[i + 1], d[i + 2]);
      const mn = Math.min(d[i], d[i + 1], d[i + 2]);
      if (mn > 238) white++;
      if (mx - mn > 60) saturated++;
    }
    return white / n > 0.6 && saturated / n < 0.08 && src.width / src.height > 1.4;
  }

  // ------------------------------------------------------------ OCR engine
  const SERIAL_CHARS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  let enginePromise = null;
  const engine = { eng: null, kor: null, error: null };

  async function assetBase() {
    const base = new URL("./", location.href).href;
    try {
      const r = await fetch(base + "tesseract/manifest.json", { cache: "no-store" });
      if (r.ok) return base + "tesseract/";
    } catch (e) { /* not published alongside: use the CDN */ }
    return null;
  }

  function withTimeout(promise, ms, message) {
    return Promise.race([promise, new Promise((_, rej) => setTimeout(() => rej(new Error(message)), ms))]);
  }

  function wasmAllowed() {
    try {
      new WebAssembly.Module(Uint8Array.of(0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00));
      return true;
    } catch (e) {
      return false;
    }
  }

  async function startEngine(onStatus) {
    if (typeof Tesseract === "undefined") throw new Error("Tesseract.js를 불러오지 못했습니다.");
    if (typeof WebAssembly === "undefined" || !wasmAllowed()) throw new Error("이 화면에서는 WebAssembly가 막혀 있습니다.");
    const local = await assetBase();
    const opts = local
      ? { workerPath: local + "worker-wrap.js", workerBlobURL: false, corePath: local + "core/", langPath: local + "lang", gzip: true, cacheMethod: "none" }
      : { cacheMethod: "none" };
    const logger = (m) => onStatus && m.status && onStatus(m.status, m.progress);
    engine.eng = await withTimeout(Tesseract.createWorker("eng", 1, { ...opts, logger }), 45000, "글자 인식 엔진 시작 시간 초과");
    engine.kor = await withTimeout(Tesseract.createWorker("kor", 1, { ...opts, logger }), 45000, "글자 인식 엔진 시작 시간 초과");
    return engine;
  }

  function ensureEngine(onStatus) {
    if (!enginePromise) {
      enginePromise = startEngine(onStatus).catch((e) => {
        engine.error = e;
        return null;
      });
    }
    return enginePromise;
  }

  async function ocr(worker, canvas, params) {
    await worker.setParameters({ tessedit_char_whitelist: "", ...params });
    const { data } = await worker.recognize(canvas);
    return data;
  }

  function datesIn(text) {
    const out = [];
    // 2025.12.29 / 2025-12-29 / 2025. 12. 29, and 20251229 when separators were lost
    const re = /(20\d{2})\s*[.\-/,]?\s*(\d{1,2})\s*[.\-/,]\s*(\d{1,2})|(?<!\d)(20\d{2})(\d{2})(\d{2})(?!\d)/g;
    let m;
    while ((m = re.exec(text))) {
      const iso = m[1] ? parseDate(`${m[1]}.${m[2]}.${m[3]}`) : parseDate(`${m[4]}.${m[5]}.${m[6]}`);
      if (iso && !out.includes(iso)) out.push(iso);
    }
    return out;
  }
  function tokensIn(text) {
    return (text.toUpperCase().match(/[0-9A-Z][0-9A-Z\-]{3,}[0-9A-Z]/g) || []).map((t) => t.replace(/-/g, ""));
  }

  /** Assign cal/next from the dates found on a label. */
  function splitDates(dates) {
    const sorted = [...dates].sort();
    if (sorted.length >= 2) return { calLast: sorted[0], calNext: sorted[sorted.length - 1] };
    if (sorted.length === 1) {
      return daysBetween(todayIso(), sorted[0]) > 3 ? { calLast: null, calNext: sorted[0] } : { calLast: sorted[0], calNext: null };
    }
    return { calLast: null, calNext: null };
  }

  // ------------------------------------------------------------ main analysis
  /**
   * Analyse one photo. Returns
   * {image, kind, labelCrop, reading:{source, serials, name, calLast, calNext, certNo, raw}, notes}
   */
  async function analyzePhoto(file, onStatus) {
    let image = await loadImage(file);
    const notes = [];
    const reading = { source: "ocr", serials: [], name: "", calLast: null, calNext: null, certNo: "", raw: [] };
    const screenshot = looksLikeScreenshot(image);
    const kind = screenshot ? "smax_screenshot" : "tool_photo";

    const qr = readQr(image);
    if (qr) {
      reading.source = "qr";
      reading.serials.push(qr);
    }

    let label = null;
    let labelCrop = null;
    if (!screenshot) {
      label = findLabel(image);
      if (label) labelCrop = cropLabel(image, label);
    }

    const eng = await ensureEngine(onStatus);
    if (!eng) {
      notes.push("글자 인식(OCR) 엔진을 쓸 수 없어 QR과 직접 선택만 사용합니다." + (engine.error ? ` (${engine.error.message})` : ""));
      return { image, kind, labelCrop, label, reading, notes };
    }

    if (screenshot) {
      const { canvas } = scaled(image, 2400);
      const copy = makeCanvas(canvas.width, canvas.height);
      copy.getContext("2d").drawImage(canvas, 0, 0);
      const data = await ocr(eng.eng, autocontrastGray(copy), { tessedit_pageseg_mode: "3" });
      reading.raw.push(data.text);
      const serialLine = data.text.match(/Serial\s*\/?\s*Asset\s*Number\s*[:\s]*([0-9A-Z\-]{4,})/i);
      if (serialLine && reading.source !== "qr") reading.serials.unshift(serialLine[1]);
      if (reading.source !== "qr") reading.serials.push(...tokensIn(data.text));
      const due = data.text.match(/Calibration\s*Due\s*Date\s*[:\s]*(\d{1,2}\/\d{1,2}\/\d{4})/i);
      if (due) reading.calNext = parseDate(due[1]);
      return { image, kind, labelCrop, label, reading, notes };
    }

    // 1) printed calibration label: dates and certificate number. A label
    //    photographed sideways or upside down is turned until its dates read.
    let turn = 0;
    if (labelCrop) {
      const inner = cropLabel(image, label, 0.02);
      const portrait = inner.height > inner.width;
      let best = await readLabel(eng.eng, portrait ? rotateCanvas(inner, 90) : inner);
      turn = portrait ? 90 : 0;
      if (best.dates.length < 2) {
        const alt = portrait ? 270 : 180;
        const other = await readLabel(eng.eng, rotateCanvas(inner, alt));
        if (other.dates.length > best.dates.length) {
          best = other;
          turn = alt;
        }
      }
      reading.raw.push(best.text);
      Object.assign(reading, splitDates(best.dates));
      reading.certNo = best.certNo;
      if (turn) {
        // continue on the upright photo so handwriting lines run horizontally
        image = rotateCanvas(image, turn);
        label = findLabel(image);
        labelCrop = label && cropLabel(image, label);
        notes.push("사진이 돌아가 있어 바로 세웠습니다.");
      }
    }

    // 2) handwritten note: serial digits (and the name, when legible). Each
    //    serial line is read in several preprocessings; all readings vote.
    if (reading.source !== "qr") {
      const lines = findInkLines(image, label && label.box);
      reading.votes = [];
      for (const line of lines) {
        const probe = cleanReading((await ocr(eng.eng, lineCrop(image, line, "raw", 110), { tessedit_pageseg_mode: "7", tessedit_char_whitelist: SERIAL_CHARS })).text);
        const digits = (probe.match(/\d/g) || []).length;
        if (probe.length >= 5 && digits >= 4) {
          reading.votes.push(probe);
          for (const [mode, height] of LINE_VARIANTS) {
            for (const wl of [DIGITS, SERIAL_CHARS]) {
              if (mode === "raw" && height === 110 && wl === SERIAL_CHARS) continue; // the probe
              const t = cleanReading((await ocr(eng.eng, lineCrop(image, line, mode, height), { tessedit_pageseg_mode: "7", tessedit_char_whitelist: wl })).text);
              if (t.length >= 4) reading.votes.push(t);
            }
          }
        } else if (!reading.name) {
          const k = await ocr(eng.kor, lineCrop(image, line, "raw", 110), { tessedit_pageseg_mode: "7" });
          const hangul = (k.text.match(/[가-힣]{2,4}/g) || [])[0];
          if (hangul) reading.name = hangul;
        }
      }
      const counts = new Map();
      for (const v of reading.votes) counts.set(v, (counts.get(v) || 0) + 1);
      reading.serials.push(...[...counts.entries()].sort((a, b) => b[1] - a[1]).map(([v]) => v));
      // 3) fallback: sparse text over the whole photo, outside the label
      if (!reading.serials.length) {
        const { canvas, scale } = scaled(image, 1600);
        const d = await ocr(eng.eng, canvas, { tessedit_pageseg_mode: "11", tessedit_char_whitelist: SERIAL_CHARS + "-." });
        for (const w of d.words || []) {
          const b = w.bbox;
          const box = { left: b.x0 / scale, top: b.y0 / scale, right: b.x1 / scale, bottom: b.y1 / scale };
          if (label && iou(box, label.box) > 0 && box.left >= label.box.left - 5 && box.right <= label.box.right + 5) continue;
          reading.serials.push(...tokensIn(w.text));
        }
        if (!reading.calLast && !reading.calNext) Object.assign(reading, splitDates(datesIn(d.text)));
      }
    }
    if (!labelCrop) notes.push("교정필증을 찾지 못했습니다.");
    return { image, kind, labelCrop, label, reading, notes };
  }

  async function readLabel(worker, inner) {
    const plain = scaled(inner, 1600).canvas;
    const k = Math.max(1, 1000 / plain.width);
    const prepared = prepareForOcr(plain, { left: 0, top: 0, right: plain.width, bottom: plain.height }, plain.height * k, 0);
    // Tesseract skips text boxed in by the label's table rules, so the later
    // passes erase them first. Stop as soon as both dates are found.
    const clean = removeLines(prepared);
    const passes = [[plain, "11"], [clean, "6"], [clean, "11"], [prepared, "11"]];
    const texts = [];
    for (const [canvas, psm] of passes) {
      texts.push((await ocr(worker, canvas, { tessedit_pageseg_mode: psm })).text);
      if (datesIn(texts.join("\n")).length >= 2) break;
    }
    const text = texts.join("\n");
    const cert = text.toUpperCase().match(/[A-Z]{1,3}\d{2}-\d{3,5}-\d{1,4}/);
    return { text, dates: datesIn(text), certNo: cert ? cert[0] : "" };
  }

  /** Copy of a grayscale canvas with long horizontal/vertical rules painted white. */
  function removeLines(src, frac = 0.3) {
    const c = makeCanvas(src.width, src.height);
    const ctx = c.getContext("2d", { willReadFrequently: true });
    ctx.drawImage(src, 0, 0);
    const img = ctx.getImageData(0, 0, c.width, c.height);
    const d = img.data;
    const W = c.width, H = c.height;
    const dark = new Uint8Array(W * H);
    for (let j = 0; j < W * H; j++) dark[j] = d[j * 4] < 110 ? 1 : 0;
    const kill = new Uint8Array(W * H);
    for (let y = 0; y < H; y++) {
      let run = 0;
      for (let x = 0; x <= W; x++) {
        if (x < W && dark[y * W + x]) run++;
        else {
          if (run > W * frac) for (let i = x - run; i < x; i++) kill[y * W + i] = 1;
          run = 0;
        }
      }
    }
    for (let x = 0; x < W; x++) {
      let run = 0;
      for (let y = 0; y <= H; y++) {
        if (y < H && dark[y * W + x]) run++;
        else {
          if (run > H * frac) for (let i = y - run; i < y; i++) kill[i * W + x] = 1;
          run = 0;
        }
      }
    }
    for (let y = 0; y < H; y++) {
      for (let x = 0; x < W; x++) {
        const j = y * W + x;
        if (kill[j] || (x > 0 && kill[j - 1]) || (x < W - 1 && kill[j + 1]) || (y > 0 && kill[j - W]) || (y < H - 1 && kill[j + W])) {
          d[j * 4] = d[j * 4 + 1] = d[j * 4 + 2] = 255;
        }
      }
    }
    ctx.putImageData(img, 0, 0);
    return c;
  }


  TC.vision = { loadImage, analyzePhoto, findLabel, cropLabel, findInkLines, readQr, rotateCanvas, makeCanvas, scaled, ensureEngine, engine, QR_PREFIX, datesIn, splitDates };
})();
