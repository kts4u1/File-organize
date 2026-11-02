# Third-party components

Bundled in `web/vendor/` and `web/tesseract/` so the desktop program works without CDN access.

| Component | Version | License | Files |
|---|---|---|---|
| JSZip | 3.10.1 | MIT (dual MIT / GPL-3.0-or-later) | `vendor/jszip.min.js` |
| jsPDF | 2.5.1 | MIT | `vendor/jspdf.umd.min.js` |
| jsQR | 1.4.0 | Apache-2.0 | `vendor/jsQR.js` |
| qrcode-generator | 1.4.4 | MIT | `vendor/qrcode.js` |
| Tesseract.js | 6.0.1 | Apache-2.0 | `vendor/tesseract.min.js`, `tesseract/worker.min.js` |
| Tesseract.js-core (Tesseract OCR, WebAssembly) | 6.1.2 | Apache-2.0 | `tesseract/core/*.wasm.js` |
| Tesseract traineddata `eng`, `kor` (tessdata, 4.0.0_best_int) | 4.0.0 | Apache-2.0 | `tesseract/lang-*.js` (base64 of the `.traineddata.gz`) |
