/* Tesseract.js worker wrapper (published as tesseract/worker-wrap.js).
 * The artifact host serves scripts but not .traineddata(.gz) files, so the
 * language data ships as lang-<code>.js (base64 of the .traineddata.gz) and
 * this wrapper answers the worker's request for it from those scripts. */
(function () {
  "use strict";
  const realFetch = self.fetch.bind(self);
  self.fetch = function (input, init) {
    const url = typeof input === "string" ? input : input.url;
    const m = /\/(eng|kor)\.traineddata(?:\.gz)?$/.exec(url);
    if (!m) return realFetch(input, init);
    const key = "TC_LANG_" + m[1];
    if (!self[key]) importScripts("lang-" + m[1] + ".js");
    const bin = atob(self[key]);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    self[key] = null; // free the string; each worker loads one language once
    return Promise.resolve(new Response(bytes, { status: 200 }));
  };
})();
importScripts("worker.min.js");
