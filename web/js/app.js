/* Page controller: file inputs, photo jobs, review cards, Excel output,
 * due-date overview and QR stickers. */
(function () {
  "use strict";
  const TC = window.TC;
  const { ToolWorkbook, todayIso, addMonths } = TC.xlsx;
  const M = TC.match;
  const V = TC.vision;

  const STATUS = {
    ok: { text: "자동 반영", cls: "ok" },
    review: { text: "확인 필요", cls: "review" },
    error: { text: "대상 없음", cls: "error" },
    skip: { text: "제외", cls: "skip" },
    pending: { text: "인식 중", cls: "pending" },
  };
  const KIND = { tool_photo: "툴 사진", smax_screenshot: "SMAX 캡쳐", other: "기타(제외)" };
  const ENGINE_STEPS = {
    "loading tesseract core": "글자 인식 엔진 불러오는 중",
    "initializing tesseract": "글자 인식 엔진 준비 중",
    "loading language traineddata": "언어 데이터 불러오는 중",
    "initializing api": "글자 인식 엔진 준비 중",
    "recognizing text": "글자 읽는 중",
  };

  const state = { excel: null, wb: null, jobs: [], busy: false, result: null, demo: false, regions: null, picked: new Set() };
  const $ = (sel) => document.querySelector(sel);

  function h(tag, attrs, ...children) {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v == null || v === false) continue;
      if (k === "class") el.className = v;
      else if (k.startsWith("on")) el.addEventListener(k.slice(2), v);
      else if (k === "checked" || k === "disabled" || k === "selected" || k === "value") el[k] = v;
      else el.setAttribute(k, v === true ? "" : v);
    }
    for (const c of children.flat()) if (c != null && c !== false) el.append(c.nodeType ? c : String(c));
    return el;
  }
  /** append() that leaves out optional parts given as null/false (plain append would print "null"). */
  const put = (el, ...parts) => el.append(...parts.flat().filter((p) => p != null && p !== false));

  // ------------------------------------------------------------ desktop program (툴검교정.exe)
  // The exe serves this page from 127.0.0.1 with a per-run token; SharePoint
  // download/upload go through it because only Edge on the PC is logged in.
  const DESKTOP_TOKEN = document.querySelector('meta[name="tc-token"]')?.content || "";
  const desktop = { on: !!DESKTOP_TOKEN, sharepoint: false, name: "", output: "" };

  async function api(path, options = {}) {
    const res = await fetch(path, { ...options, headers: { ...(options.headers || {}), "X-TC-Token": DESKTOP_TOKEN } });
    return res;
  }
  async function apiJson(path, options) {
    const res = await api(path, options);
    let body = null;
    try { body = await res.json(); } catch (e) { /* not JSON */ }
    return body || { ok: false, message: `오류 (${res.status})` };
  }

  // ------------------------------------------------------------ downloads
  const inViewer = !!(window.claude && typeof window.claude.use === "function");
  const downloadsReady = inViewer ? window.claude.use("downloads").catch(() => null) : Promise.resolve(null);

  async function saveFile(blob, filename) {
    if (desktop.on) {
      const r = await apiJson(`/api/save?name=${encodeURIComponent(filename)}`, { method: "POST", body: blob });
      return r.ok ? { ok: true, path: r.path } : { ok: false, message: r.message };
    }
    const ns = await downloadsReady;
    if (ns) {
      try {
        await ns.save({ filename, data: blob });
        return { ok: true };
      } catch (e) {
        if (e && e.code === "declined") return { ok: false, message: "저장을 취소했습니다." };
        if (e && e.code === "rate_limited") return { ok: false, message: "저장 창이 이미 열려 있습니다. 잠시 후 다시 누르세요." };
        return { ok: false, message: `파일을 저장할 수 없습니다 (${(e && e.code) || "오류"}).` };
      }
    }
    if (inViewer) return { ok: false, message: "이 화면에서는 파일 저장이 허용되지 않습니다." };
    const url = URL.createObjectURL(blob);
    const a = h("a", { href: url, download: filename });
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    return { ok: true };
  }

  function toast(text, kind) {
    const el = h("div", { class: `toast ${kind || ""}`, role: "status" }, text);
    $("#toasts").append(el);
    setTimeout(() => el.remove(), 5000);
  }

  // ------------------------------------------------------------ excel
  /** `bytes` may be a promise (still reading the file); photos chosen meanwhile wait for it. */
  function loadExcel(bytes, name, demo, source = demo ? "demo" : "file") {
    if (state.busy) {
      toast("사진을 읽는 중에는 엑셀을 바꿀 수 없습니다. 끝난 뒤 다시 선택하세요.", "bad");
      return Promise.resolve();
    }
    const loading = Promise.resolve(bytes)
      .then((b) => openExcel(b, name, demo, source))
      .finally(() => { if (state.loading === loading) state.loading = null; });
    state.loading = loading;
    return loading;
  }

  async function openExcel(bytes, name, demo, source) {
    try {
      const wb = await ToolWorkbook.open(bytes.slice(0), name);
      state.excel = { bytes, name, source };
      state.wb = wb;
      state.demo = !!demo;
      state.jobs = [];
      state.result = null;
      state.script = null;
      state.regions = null;
      state.picked = new Set(wb.records.filter((r) => !r.disposed).map((r) => r.key));
      renderAll();
    } catch (e) {
      toast(`엑셀을 읽을 수 없습니다: ${e.message}`, "bad");
    }
  }

  function excelSummary() {
    const box = $("#excel-summary");
    box.replaceChildren();
    if (!state.wb) {
      box.append(h("span", { class: "muted" }, "아직 선택하지 않았습니다"));
      return;
    }
    const counts = [...state.wb.sheets.keys()].map((s) => `${s} ${state.wb.records.filter((r) => r.sheet === s).length}`);
    box.append(h("strong", {}, state.excel.name), h("span", { class: "muted" }, ` · 툴 ${state.wb.records.length}개 (${counts.join(", ")})`));
    if (state.demo) box.append(h("span", { class: "chip demo" }, "예시 데이터"));
  }

  // ------------------------------------------------------------ photos
  async function processFiles(files) {
    if (state.loading) {
      setProgress("SharePoint 파일을 받는 중입니다. 다 받으면 사진을 읽습니다…");
      await state.loading;
      setProgress("");
    }
    if (!state.wb) {
      toast("Tool Verification 엑셀 파일을 먼저 선택하세요.", "bad");
      return;
    }
    if (state.busy) return;
    state.busy = true;
    state.result = null;
    showTab("update");
    const newJobs = files.map((file) => ({ id: Math.random().toString(36).slice(2), file, fileName: file.name || "photo.jpg", status: "pending", messages: [] }));
    state.jobs.push(...newJobs);
    renderJobs();
    let n = 0;
    for (const job of newJobs) {
      n++;
      setProgress(`${n}/${newJobs.length} — ${job.fileName}`);
      try {
        const res = await V.analyzePhoto(job.file, (status, p) => {
          const label = ENGINE_STEPS[status] || status;
          setProgress(`${n}/${newJobs.length} — ${job.fileName} · ${label}${p != null && p < 1 ? ` ${Math.round(p * 100)}%` : ""}`);
        });
        Object.assign(job, {
          image: res.image,
          labelCrop: res.labelCrop,
          kind: res.kind,
          reading: res.reading,
          notes: res.notes,
          calLast: res.kind === "tool_photo" ? res.reading.calLast : null,
          calNext: res.kind === "tool_photo" ? res.reading.calNext : null,
        });
        job.thumb = thumbnail(job.labelCrop && job.kind === "tool_photo" ? job.labelCrop : job.image);
        job.candidates = M.findCandidates(res.reading, state.wb.records);
        const d = M.decide(job.candidates, res.reading.source);
        job.target = d.candidate ? d.candidate.record : null;
        job.matchStatus = d.status;
        job.matchMessages = d.messages;
      } catch (e) {
        Object.assign(job, { error: e.message || String(e), kind: "tool_photo", reading: null, candidates: [], target: null, matchStatus: "error", matchMessages: [] });
      }
      M.evaluateAll(state.jobs.filter((j) => j.status !== "pending" || j === job));
      job.include = job.status === "ok";
      renderJobs();
    }
    state.busy = false;
    setProgress("");
    renderJobs();
  }

  function thumbnail(canvas) {
    const { canvas: small } = V.scaled(canvas, 520);
    return small.toDataURL("image/jpeg", 0.85);
  }

  function setProgress(text) {
    const el = $("#progress");
    el.textContent = text;
    el.hidden = !text;
  }

  function reevaluate() {
    M.evaluateAll(state.jobs.filter((j) => j.status !== "pending"));
    for (const job of state.jobs) if (job.status === "error" || job.status === "skip") job.include = false;
    renderJobs();
  }

  // ------------------------------------------------------------ job cards
  function targetSelect(job) {
    const sel = h("select", { id: `target-${job.id}`, "aria-label": "엑셀 대상 행" });
    sel.append(h("option", { value: "" }, "(대상 선택 안 함)"));
    const suggested = (job.candidates || []).slice(0, 4).map((c) => c.record);
    if (suggested.length) {
      const g = h("optgroup", { label: "추천" });
      for (const rec of suggested) g.append(h("option", { value: rec.key }, M.label(rec)));
      sel.append(g);
    }
    for (const sheet of state.wb.sheets.keys()) {
      const g = h("optgroup", { label: sheet });
      for (const rec of state.wb.records.filter((r) => r.sheet === sheet)) g.append(h("option", { value: "all:" + rec.key }, M.label(rec)));
      sel.append(g);
    }
    const current = job.target ? job.target.key : "";
    sel.value = suggested.some((r) => r.key === current) ? current : current ? "all:" + current : "";
    sel.addEventListener("change", () => {
      const key = sel.value.replace(/^all:/, "");
      job.target = state.wb.records.find((r) => r.key === key) || null;
      job.targetConfirmed = !!job.target;
      M.evaluateAll(state.jobs.filter((j) => j.status !== "pending"));
      job.include = job.status === "ok";
      reevaluate();
    });
    return sel;
  }

  function renderJob(job) {
    const st = STATUS[job.status] || STATUS.pending;
    const head = h("div", { class: "job-head" },
      h("span", { class: "job-file" }, job.fileName),
      h("span", { class: `pill ${st.cls}` }, st.text));
    if (job.status === "pending") {
      return h("article", { class: "job pending" }, head, h("div", { class: "job-wait" }, h("span", { class: "spinner", "aria-hidden": "true" }), "사진을 분석하고 있습니다"));
    }
    const r = job.reading;
    const read = h("dl", { class: "readout" },
      h("dt", {}, "인식 방법"), h("dd", {}, r ? (r.source === "qr" ? "QR 스티커" : "글자 인식") : "-"),
      h("dt", {}, "시리얼"), h("dd", { class: "mono" }, r && r.serials.length ? r.serials.slice(0, 3).join(" · ") : "-"),
      h("dt", {}, "이름"), h("dd", {}, (r && r.name) || "-"),
      h("dt", {}, "교정일"), h("dd", { class: "mono" }, (r && r.calLast) || "-"),
      h("dt", {}, "차기교정일"), h("dd", { class: "mono" }, (r && r.calNext) || "-"),
      h("dt", {}, "성적서번호"), h("dd", { class: "mono" }, (r && r.certNo) || "-"));

    const kind = h("select", { id: `kind-${job.id}`, "aria-label": "사진 종류" },
      ...Object.entries(KIND).map(([k, v]) => h("option", { value: k, selected: job.kind === k }, v)));
    kind.addEventListener("change", () => {
      job.kind = kind.value;
      reevaluate();
    });
    const controls = h("div", { class: "controls" },
      h("label", { class: "field" }, h("span", {}, "사진 종류"), kind),
      h("label", { class: "field" }, h("span", {}, "엑셀 대상 행"), targetSelect(job)));
    if (job.kind === "tool_photo") {
      const mk = (field, text) => {
        const input = h("input", { type: "date", id: `${field}-${job.id}`, value: job[field] || "" });
        input.addEventListener("change", () => {
          job[field] = input.value || null;
          reevaluate();
        });
        return h("label", { class: "field" }, h("span", {}, text), input);
      };
      controls.append(h("div", { class: "dates" }, mk("calLast", "교정일 (Cal Last)"), mk("calNext", "차기교정일 (Cal Next)")));
    }
    const inc = h("input", { type: "checkbox", id: `inc-${job.id}`, checked: !!job.include, disabled: job.status === "error" || job.status === "skip" });
    inc.addEventListener("change", () => {
      job.include = inc.checked;
      renderActions();
    });
    controls.append(h("label", { class: "include" }, inc, h("span", {}, "엑셀에 반영")));

    const figure = h("figure", { class: "thumb" },
      job.thumb ? h("img", { src: job.thumb, alt: "엑셀에 들어갈 사진" }) : h("div", { class: "noimg" }, "사진 없음"),
      h("figcaption", {}, job.kind === "tool_photo" ? (job.labelCrop ? "교정필증 자동 크롭" : "사진 전체") : "엑셀에 들어갈 사진"));
    const notes = [...(job.messages || []), ...(job.notes || [])];
    return h("article", { class: `job ${st.cls}` }, head,
      h("div", { class: "job-body" }, figure, read, controls),
      notes.length ? h("ul", { class: "notes" }, ...notes.map((m) => h("li", {}, m))) : null);
  }

  function renderJobs() {
    const list = $("#jobs");
    list.replaceChildren(...state.jobs.map(renderJob));
    $("#jobs-empty").hidden = state.jobs.length > 0 || !!state.result;
    const counts = {};
    for (const j of state.jobs) counts[j.status] = (counts[j.status] || 0) + 1;
    $("#job-counts").replaceChildren(...["ok", "review", "error", "skip"].filter((s) => counts[s]).map((s) =>
      h("span", { class: `pill ${STATUS[s].cls}` }, `${STATUS[s].text} ${counts[s]}`)));
    renderActions();
  }

  /** In the exe, a workbook downloaded from SharePoint is uploaded back right after it is built. */
  function uploadsDirectly() {
    return desktop.on && desktop.sharepoint && state.excel?.source === "sharepoint";
  }

  function renderActions() {
    const n = state.jobs.filter(M.applicable).length;
    const btn = $("#apply");
    btn.textContent = uploadsDirectly() ? `SharePoint에 반영 (${n}건)` : `엑셀 파일 만들기 (${n}건)`;
    btn.disabled = !n || state.busy || !!state.loading;
    const sb = $("#make-script");
    sb.hidden = desktop.on; // the exe offers the script only when SharePoint refuses the upload (locked)
    sb.textContent = `Excel 웹 스크립트 만들기 (${n}건)`;
    sb.disabled = !n || state.busy;
    $("#actions").hidden = !state.jobs.length;
    $("#clear").disabled = state.busy;
  }

  // ------------------------------------------------------------ apply
  async function canvasBytes(canvas, type, maxSide, quality) {
    const { canvas: c } = V.scaled(canvas, maxSide);
    const blob = await new Promise((res) => c.toBlob(res, type, quality));
    return { bytes: new Uint8Array(await blob.arrayBuffer()), size: [c.width, c.height] };
  }

  async function applyJobs() {
    const btn = $("#apply");
    btn.disabled = true;
    btn.textContent = "엑셀 만드는 중…";
    try {
      const wb = await ToolWorkbook.open(state.excel.bytes.slice(0), state.excel.name);
      const changes = [];
      const missing = [];
      const today = todayIso();
      const norm = M.normalizeSerial;
      for (const job of state.jobs) {
        if (!M.applicable(job)) continue;
        // the row may have moved if rows were added meanwhile: find the tool again by serial
        let rec = wb.findRecord(job.target.sheet, job.target.row);
        if (!rec || norm(rec.serial) !== norm(job.target.serial)) {
          rec = wb.records.find((r) => r.sheet === job.target.sheet && norm(r.serial) === norm(job.target.serial)) || null;
        }
        if (!rec) {
          missing.push(job);
          continue;
        }
        const before = { calLast: rec.calLast, calNext: rec.calNext };
        const notes = [];
        let replaced = 0;
        const columns = wb.sheets.get(rec.sheet).columns;
        if (job.kind === "tool_photo") {
          for (const [key, value] of [["cal_last", job.calLast], ["cal_next", job.calNext]]) {
            if ((await wb.writeDate(rec, key, value)) === "formula") notes.push(`${key === "cal_next" ? "Cal Next" : "Cal Last"} 열은 수식이라 그대로 두었습니다.`);
          }
          if (columns.label_photo) {
            const pic = await canvasBytes(job.labelCrop || job.image, "image/jpeg", 1000, 0.85);
            replaced = await wb.replacePicture(rec, "label_photo", pic.bytes, "jpeg", pic.size, `Label ${rec.serial}`, `${job.fileName} (${today})`);
          }
        } else if (columns.smax_photo) {
          const pic = await canvasBytes(job.image, "image/png", 1600);
          replaced = await wb.replacePicture(rec, "smax_photo", pic.bytes, "png", pic.size, `SMAX ${rec.serial}`, `${job.fileName} (${today})`);
        }
        changes.push({ job, rec, before, replaced, notes, after: { calLast: rec.calLast, calNext: rec.calNext } });
      }
      if (!changes.length) {
        toast("엑셀에서 대상 툴을 찾지 못해 반영할 것이 없습니다.", "bad");
        renderActions();
        return;
      }
      const blob = await wb.toBlob();
      const bytes = new Uint8Array(await blob.arrayBuffer());
      const direct = uploadsDirectly();
      state.result = { blob, changes, missing, name: state.excel.name, source: state.excel.source, testMode: false, upload: null };
      // the updated file becomes the working copy for the next batch
      state.excel = { bytes, name: state.excel.name, source: state.excel.source };
      state.wb = await ToolWorkbook.open(bytes.slice(0), state.excel.name);
      state.jobs = [];
      renderAll();
      await offerResult(); // keeps a copy (in the exe: %USERPROFILE%\ToolCal\output)
      if (direct) await uploadToSharePoint();
    } catch (e) {
      console.error(e);
      toast(`엑셀을 만들지 못했습니다: ${e.message}`, "bad");
      renderActions();
    }
  }

  // ------------------------------------------------------------ office script
  function makeScript() {
    showScript(state.jobs, false);
  }

  function showScript(jobs, testMode) {
    const script = TC.officeScript.build(jobs, testMode);
    state.script = { ...script, testMode };
    renderScript();
    $("#script-panel").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  /** The same changes as the built file, as a script: works while others have the file open. */
  function scriptFromResult() {
    const result = state.result;
    if (!result) return;
    showScript(result.changes.map((c) => c.job), result.testMode);
  }

  function renderScript() {
    const box = $("#script-panel");
    box.replaceChildren();
    box.hidden = !state.script;
    if (!state.script) return;
    const { text, count, size, testMode } = state.script;
    const area = h("textarea", { id: "script-text", readonly: true, "aria-label": "스크립트 내용" });
    area.value = text;
    const copy = async () => {
      try {
        await navigator.clipboard.writeText(text);
        toast("스크립트를 복사했습니다. Excel 웹의 새 스크립트에 붙여넣으세요.", "good");
      } catch (e) {
        details.open = true;
        area.focus();
        area.select();
        toast("자동 복사가 막혀 있어 내용을 선택해 두었습니다. Ctrl+C로 복사하세요.", "bad");
      }
    };
    const save = async () => {
      const r = await saveFile(new Blob([text], { type: "text/plain" }), `검교정_반영_스크립트_${todayIso()}.txt`);
      toast(r.ok ? "스크립트 파일을 저장했습니다." : r.message, r.ok ? "good" : "bad");
    };
    const details = h("details", {}, h("summary", {}, "스크립트 직접 보기"), area);
    box.append(...[  // optional parts are null: leave them out instead of printing "null"
      h("div", { class: "result-head" },
        h("h3", {}, `Excel 웹에서 실행할 스크립트를 만들었습니다 (${count}건 · ${Math.ceil(size / 1024)} KB)`),
        h("div", { class: "row-gap" },
          h("button", { class: "primary", onclick: copy }, "스크립트 복사"),
          h("button", { onclick: save }, ".txt로 저장"),
          desktop.on && desktop.url
            ? h("a", { class: "button-link", href: desktop.url, target: "_blank", rel: "noopener" }, "Excel 웹에서 파일 열기")
            : null,
          h("button", { onclick: () => { state.script = null; renderScript(); } }, "닫기"))),
      h("p", { class: "muted" }, testMode
        ? "지역 시트는 그대로 두고, 바뀔 행을 'TEST' 탭에 복사해 새 날짜와 사진을 넣습니다."
        : "지역 시트의 해당 행에 새 날짜를 쓰고, 그 행의 교정필증·SMAX 사진을 바꿉니다."),
      h("ol", {},
        h("li", {}, "SharePoint에서 엑셀을 브라우저(Excel 웹)로 엽니다."),
        h("li", {}, "위쪽 ", h("b", {}, "자동화"), " 탭 → ", h("b", {}, "새 스크립트"), "를 누릅니다."),
        h("li", {}, "편집기 내용을 모두 지우고(", h("kbd", {}, "Ctrl+A"), ") 붙여넣은 뒤(", h("kbd", {}, "Ctrl+V"), ") ", h("b", {}, "실행"), "을 누릅니다."),
        h("li", {}, "아래 출력에 '완료: ", String(count), "/", String(count), "건'이 나오면 끝입니다. 파일은 SharePoint에 바로 저장됩니다.")),
      h("p", { class: "muted small" }, "자동화 탭이 보이지 않으면 회사 설정에서 Office Scripts가 꺼져 있는 것입니다. 그때는 '엑셀 파일 만들기'로 받은 파일을 SharePoint에 올려 주세요."),
      size > 1.5 * 1024 * 1024
        ? h("p", { class: "muted small", style: "color: var(--review)" }, "사진이 많아 스크립트가 큽니다. 붙여넣기가 느리거나 실패하면 사진을 10장 정도씩 나눠서 만드세요.")
        : null,
      details].filter(Boolean));
  }

  async function offerResult() {
    if (!state.result) return;
    const r = await saveFile(state.result.blob, state.result.name);
    if (r.ok) toast(r.path ? `엑셀 사본을 저장했습니다: ${r.path}` : "엑셀 파일을 저장했습니다.", "good");
    else toast(r.message, "bad");
  }

  // ------------------------------------------------------------ SharePoint (desktop only)
  /** Show the exe's latest status line (e.g. "SharePoint에서 받는 중… 45.2 / 207.8 MB") while `promise` runs. */
  function followProgress(promise, show) {
    let stop = false;
    const tick = async () => {
      while (!stop) {
        const p = await apiJson("/api/progress").catch(() => null);
        if (!stop && p && p.message) show(p.message);
        await new Promise((r) => setTimeout(r, 1000));
      }
    };
    tick();
    return promise.finally(() => { stop = true; });
  }

  /** Download the latest workbook from SharePoint. Returns true when it was loaded. */
  function downloadFromSharePoint() {
    if (state.busy) {
      toast("사진을 읽는 중에는 엑셀을 바꿀 수 없습니다. 끝난 뒤 다시 받으세요.", "bad");
      return Promise.resolve(false);
    }
    if (state.loading) return state.loading.then(() => !!state.excel && state.excel.source === "sharepoint");
    const btn = $("#sp-download");
    const label = $("#sp-name");
    btn.disabled = true;
    btn.textContent = "SharePoint에서 받는 중…";
    label.textContent = "SharePoint 연결 중… (로그인 창이 뜨면 로그인해 주세요)";
    const run = (async () => {
      try {
        const res = await api("/api/sp/download", { method: "POST" });
        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          toast(err.message || `받지 못했습니다 (${res.status}).`, "bad");
          label.textContent = "받지 못했습니다. 버튼을 눌러 다시 받으세요.";
          return false;
        }
        const name = decodeURIComponent(res.headers.get("X-File-Name") || "Tool_Verification.xlsx");
        const modified = decodeURIComponent(res.headers.get("X-Modified") || "");
        label.textContent = "엑셀 여는 중…";
        await openExcel(new Uint8Array(await res.arrayBuffer()), name, false, "sharepoint");
        desktop.name = name;
        label.textContent = `SharePoint: ${name}${modified ? ` · 수정 ${modified.replace("T", " ").replace("Z", " UTC")}` : ""}`;
        toast("SharePoint 최신본을 받았습니다. 사진을 넣으면 됩니다.", "good");
        return true;
      } catch (e) {
        toast(`받지 못했습니다: ${e.message}`, "bad");
        label.textContent = "받지 못했습니다. 버튼을 눌러 다시 받으세요.";
        return false;
      } finally {
        btn.disabled = false;
        btn.textContent = "SharePoint에서 다시 받기";
      }
    })();
    const loading = followProgress(run, (m) => { if (/받는 중|여는 중|로그인/.test(m)) label.textContent = m; })
      .finally(() => { if (state.loading === loading) state.loading = null; renderActions(); });
    state.loading = loading;
    renderActions();
    return loading;
  }

  /** The file changed on SharePoint after our download: redo the same photos on the latest version. */
  async function reapplyOnLatest() {
    const jobs = state.result ? [...state.result.changes.map((c) => c.job), ...(state.result.missing || [])] : [];
    state.result = null;
    renderAll();
    if (!(await downloadFromSharePoint())) return;
    state.jobs = jobs; // targets are found again by serial in the new file
    renderJobs();
    await applyJobs();
  }

  async function uploadToSharePoint() {
    const result = state.result;
    if (!result) return;
    result.upload = { busy: true, message: `SharePoint에 올리는 중… (${(result.blob.size / 1e6).toFixed(0)} MB)` };
    renderResult();
    const r = await followProgress(apiJson("/api/sp/upload", { method: "POST", body: result.blob }), (m) => {
      if (result.upload && result.upload.busy && /올리는 중|확인 중/.test(m)) {
        result.upload.message = m;
        const el = document.getElementById("sp-progress");
        if (el) el.textContent = m;
      }
    });
    result.upload = r.ok ? { done: true, ...r } : { error: r.message || "올리지 못했습니다.", code: r.code };
    renderResult();
    toast(r.ok ? "SharePoint에 올렸습니다." : result.upload.error, r.ok ? "good" : "bad");
  }

  function sharePointBox(result) {
    if (!desktop.on || !desktop.sharepoint) return null;
    const up = result.upload || {};
    if (result.source !== "sharepoint") {
      // only results built on a SharePoint download may go back (demo data or an old local copy must not)
      const why = result.source === "demo"
        ? "예시 엑셀로 만든 결과라 SharePoint에는 올리지 않습니다."
        : "PC에 있던 엑셀로 만든 결과라 SharePoint에는 올리지 않습니다 (다른 사람의 최신 수정을 덮어쓸 수 있어서).";
      const fetchLatest = () => {
        state.result = null;
        renderAll();
        $("#excel-drop").scrollIntoView({ behavior: "smooth", block: "start" });
        downloadFromSharePoint();
      };
      return h("div", { class: "sp-box" },
        h("p", {}, why),
        h("p", { class: "muted" }, "SharePoint에서 최신 파일을 받아 사진을 반영하면, 이 자리에 'SharePoint에 올리기' 버튼이 나옵니다."),
        h("div", { class: "row-gap" }, h("button", { class: "primary", onclick: fetchLatest }, "SharePoint에서 최신 파일 받기")));
    }
    if (up.done) {
      return h("div", { class: "sp-box done" },
        h("strong", {}, `SharePoint에 올렸습니다: ${up.name}`),
        h("p", { class: "muted" }, `지금 SharePoint 파일의 시트: ${(up.sheets || []).join(", ")}`,
          up.modified ? ` · 수정 시각 ${String(up.modified).replace("T", " ").replace("Z", " UTC")}` : ""),
        h("p", { class: "muted small" }, "이전 버전은 SharePoint의 '버전 기록'에서 되돌릴 수 있습니다."));
    }
    if (up.busy) {
      return h("div", { class: "sp-box" },
        h("strong", {}, `SharePoint의 '${desktop.name || result.name}'에 올리는 중`),
        h("p", { id: "sp-progress" }, up.message || "SharePoint에 올리는 중…"),
        h("p", { class: "muted small" }, "파일이 크면 1~2분 걸립니다. 이 창을 닫지 마세요."));
    }
    return h("div", { class: `sp-box ${up.error ? "bad" : ""}` },
      h("strong", {}, `SharePoint의 '${desktop.name || result.name}'에 올리기`),
      h("p", {}, `지역 시트의 ${result.changes.length}건(날짜·사진)이 바뀝니다. 받은 뒤 다른 사람이 파일을 바꿨거나 지금 편집 중이면 올리지 않고 알려 드립니다.`),
      up.error ? h("p", { class: "sp-error" }, up.error) : null,
      up.code === "locked"
        ? h("p", { class: "muted" }, "Excel 웹 스크립트는 열려 있는 파일 안에서 바뀐 부분만 넣기 때문에, 다른 사람이 파일을 열어 두어도 반영됩니다.")
        : null,
      h("div", { class: "row-gap" },
        up.code === "locked"
          ? h("button", { class: "primary", id: "sp-script", onclick: scriptFromResult }, "Excel 웹 스크립트로 반영")
          : null,
        up.code === "changed"
          ? h("button", { class: "primary", id: "sp-reapply", onclick: reapplyOnLatest }, "최신 파일로 다시 반영")
          : h("button", { class: up.code === "locked" ? "" : "primary", id: "sp-upload", onclick: uploadToSharePoint },
            up.error ? "다시 올리기" : "SharePoint에 올리기")));
  }

  function renderResult() {
    const box = $("#result");
    box.replaceChildren();
    box.hidden = !state.result;
    if (!state.result) return;
    const { changes, missing } = state.result;
    const direct = desktop.on && desktop.sharepoint && state.result.source === "sharepoint";
    const rows = changes.map((c) => h("tr", {},
      h("td", {}, c.job.fileName),
      h("td", {}, M.label(c.rec)),
      h("td", {}, KIND[c.job.kind]),
      h("td", { class: "mono" }, c.job.kind === "tool_photo" ? `${c.before.calLast || "-"} → ${c.after.calLast}` : ""),
      h("td", { class: "mono" }, c.job.kind === "tool_photo" ? `${c.before.calNext || "-"} → ${c.after.calNext}` : ""),
      h("td", {}, c.replaced ? `교체 (${c.replaced}장 삭제)` : "추가", c.notes.length ? h("div", { class: "muted" }, c.notes.join(" ")) : null)));
    put(box,
      h("div", { class: "result-head" },
        h("h3", {}, `${changes.length}건을 반영했습니다`),
        h("div", { class: "row-gap" },
          h("button", { onclick: offerResult }, desktop.on ? "PC에 사본 저장" : "엑셀 다시 저장"),
          h("button", { onclick: () => { state.result = null; renderAll(); } }, "닫기"))),
      h("p", { class: "muted" }, direct
        ? `파일 이름: ${state.result.name} — 아래에서 SharePoint에 올린 결과를 확인하세요. 다음 사진은 이 파일에 이어서 반영됩니다.`
        : `파일 이름: ${state.result.name} — 회사 공유 폴더의 원본을 이 파일로 바꿔 주세요. 다음 사진은 이 파일에 이어서 반영됩니다.`),
      missing && missing.length
        ? h("p", { class: "sp-error" }, `엑셀에서 툴을 찾지 못해 빠진 사진: ${missing.map((j) => j.fileName).join(", ")}`)
        : null,
      h("div", { class: "table-wrap" }, h("table", {},
        h("thead", {}, h("tr", {}, ...["사진", "대상", "종류", "교정일", "차기교정일", "사진"].map((t) => h("th", {}, t)))),
        h("tbody", {}, ...rows))));
    const sp = sharePointBox(state.result);
    if (sp) box.append(sp);
  }

  // ------------------------------------------------------------ due dates
  function renderStatus() {
    const box = $("#status-body");
    box.replaceChildren();
    if (!state.wb) {
      box.append(h("p", { class: "muted" }, "엑셀 파일을 선택하면 지역별 만료·임박 툴을 보여줍니다."));
      return;
    }
    const warn = +($("#warn-days").value || 30);
    const regions = [...state.wb.sheets.keys()];
    if (!state.regions) state.regions = new Set(regions);
    const chips = h("div", { class: "chips", role: "group", "aria-label": "지역" }, ...regions.map((r) =>
      h("button", { class: `chip ${state.regions.has(r) ? "on" : ""}`, "aria-pressed": String(state.regions.has(r)), onclick: () => {
        state.regions.has(r) ? state.regions.delete(r) : state.regions.add(r);
        renderStatus();
      } }, r)));
    const rows = M.dueStatus(state.wb.records.filter((r) => state.regions.has(r.sheet)), warn);
    const count = (s) => rows.filter((r) => r.state === s).length;
    const stats = h("div", { class: "stats" },
      ...[["만료", count("만료"), "error"], [`${warn}일 이내`, count("임박"), "review"], ["정상", count("정상"), "ok"], ["전체", rows.length, ""]].map(([t, n, c]) =>
        h("div", { class: `stat ${c}` }, h("span", { class: "stat-n" }, n), h("span", { class: "stat-t" }, t))));
    const cls = { 만료: "error", 임박: "review", 정상: "ok", "날짜 없음": "skip", 폐기: "skip" };
    const table = h("table", {},
      h("thead", {}, h("tr", {}, ...["상태", "남은 일수", "차기교정일", "교정일", "지역", "행", "담당", "모델", "Serial"].map((t) => h("th", {}, t)))),
      h("tbody", {}, ...rows.map(({ state: s, days, rec }) => h("tr", { class: `due-${cls[s]}` },
        h("td", {}, h("span", { class: `pill ${cls[s]}` }, s)),
        h("td", { class: "mono num" }, days == null ? "" : `${days}일`),
        h("td", { class: "mono" }, rec.calNext || ""),
        h("td", { class: "mono" }, rec.calLast || ""),
        h("td", {}, rec.sheet),
        h("td", { class: "mono num" }, rec.row),
        h("td", {}, rec.names.join(", ") || rec.asset),
        h("td", {}, `${rec.manufacturer} ${rec.model}`.trim()),
        h("td", { class: "mono" }, rec.serial)))));
    box.append(chips, stats, h("div", { class: "table-wrap" }, table), h("p", { class: "muted small" }, `기준일 ${todayIso()} · 엑셀 'Cal Next' 열 기준`));
  }

  // ------------------------------------------------------------ stickers
  async function renderStickers() {
    const box = $("#sticker-body");
    box.replaceChildren();
    if (!state.wb) {
      box.append(h("p", { class: "muted" }, "엑셀 파일을 선택하면 툴마다 이름·시리얼·QR이 들어간 스티커를 만들 수 있습니다."));
      return;
    }
    const groups = [...state.wb.sheets.keys()].map((sheet) => {
      const recs = state.wb.records.filter((r) => r.sheet === sheet);
      const all = h("input", { type: "checkbox", id: `st-all-${sheet}`, checked: recs.every((r) => state.picked.has(r.key)) });
      all.addEventListener("change", () => {
        for (const r of recs) all.checked ? state.picked.add(r.key) : state.picked.delete(r.key);
        renderStickers();
      });
      return h("fieldset", { class: "st-group" },
        h("legend", {}, h("label", {}, all, ` ${sheet} (${recs.length})`)),
        ...recs.map((r) => {
          const cb = h("input", { type: "checkbox", id: `st-${r.key}`, checked: state.picked.has(r.key) });
          cb.addEventListener("change", () => {
            cb.checked ? state.picked.add(r.key) : state.picked.delete(r.key);
            updateStickerButton();
          });
          return h("label", { class: "st-item" }, cb,
            h("span", {}, r.names[0] || r.asset), h("span", { class: "mono muted" }, r.serial),
            h("span", { class: "muted small" }, `${r.manufacturer} ${r.model}${r.disposed ? " · 폐기" : ""}`));
        }));
    });
    const button = h("button", { class: "primary", id: "make-stickers", onclick: makeStickers });
    const preview = h("div", { class: "st-preview" });
    box.append(h("div", { class: "st-top" }, preview, h("div", {},
      h("p", {}, "스티커를 출력해 툴에 붙이면, 사진의 QR로 툴을 정확히 찾습니다 (손글씨보다 확실합니다). 교정필증과 함께 한 장에 나오게 찍어 주세요."),
      h("p", { class: "muted small" }, `A4 한 장에 ${TC.stickers.PER_PAGE}개 · 스티커 ${TC.stickers.STICKER.w}×${TC.stickers.STICKER.h}mm · PDF를 '실제 크기(100%)'로 인쇄`),
      button)), ...groups);
    updateStickerButton();
    const first = state.wb.records.find((r) => state.picked.has(r.key)) || state.wb.records[0];
    if (first) {
      const c = await TC.stickers.previewCanvas(first);
      preview.replaceChildren(h("img", { src: c.toDataURL("image/png"), alt: `스티커 예시: ${first.serial}` }));
    }
  }

  function updateStickerButton() {
    const b = $("#make-stickers");
    if (!b) return;
    const n = state.wb.records.filter((r) => state.picked.has(r.key)).length;
    b.textContent = `스티커 PDF 만들기 (${n}개)`;
    b.disabled = !n;
  }

  async function makeStickers() {
    const b = $("#make-stickers");
    b.disabled = true;
    b.textContent = "PDF 만드는 중…";
    try {
      const recs = state.wb.records.filter((r) => state.picked.has(r.key));
      const blob = await TC.stickers.buildPdf(recs);
      const r = await saveFile(blob, "툴_QR스티커.pdf");
      toast(r.ok ? "스티커 PDF를 저장했습니다." : r.message, r.ok ? "good" : "bad");
    } catch (e) {
      toast(`PDF를 만들지 못했습니다: ${e.message}`, "bad");
    }
    updateStickerButton();
  }

  // ------------------------------------------------------------ demo
  async function startDemo() {
    const bin = atob(TC.DEMO_XLSX_B64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    await loadExcel(bytes, "예시_Tool_Verification.xlsx", true);
    const photos = await TC.demoPhotos(state.wb.records);
    await processFiles(photos);
  }

  // ------------------------------------------------------------ tabs & wiring
  function showTab(name) {
    for (const t of document.querySelectorAll("[role=tab]")) {
      const on = t.dataset.tab === name;
      t.setAttribute("aria-selected", String(on));
      t.tabIndex = on ? 0 : -1;
    }
    for (const p of document.querySelectorAll("[role=tabpanel]")) p.hidden = p.id !== `panel-${name}`;
    try { localStorage.setItem("tc-tab", name); } catch (e) { /* storage unavailable */ }
  }

  function renderAll() {
    excelSummary();
    renderJobs();
    renderResult();
    renderStatus();
    renderStickers();
    $("#photo-drop").classList.toggle("disabled", !state.wb);
  }

  function wireDrop(zone, input, onFiles) {
    zone.addEventListener("dragover", (e) => {
      e.preventDefault();
      zone.classList.add("over");
    });
    zone.addEventListener("dragleave", () => zone.classList.remove("over"));
    zone.addEventListener("drop", (e) => {
      e.preventDefault();
      zone.classList.remove("over");
      onFiles([...e.dataTransfer.files]);
    });
    input.addEventListener("change", () => {
      onFiles([...input.files]);
      input.value = "";
    });
  }

  function init() {
    wireDrop($("#excel-drop"), $("#excel-input"), (files) => {
      const f = files.find((x) => /\.xlsx$/i.test(x.name));
      if (!f) return toast("xlsx 파일을 선택하세요.", "bad");
      return loadExcel(f.arrayBuffer().then((b) => new Uint8Array(b)), f.name, false);
    });
    wireDrop($("#photo-drop"), $("#photo-input"), (files) => {
      const imgs = files.filter((f) => /^image\//.test(f.type) || /\.(jpe?g|png|heic|heif|webp|bmp)$/i.test(f.name));
      if (imgs.length) processFiles(imgs);
    });
    $("#demo").addEventListener("click", startDemo);
    $("#apply").addEventListener("click", applyJobs);
    $("#make-script").addEventListener("click", makeScript);
    $("#clear").addEventListener("click", () => {
      state.jobs = [];
      renderJobs();
    });
    $("#warn-days").addEventListener("change", renderStatus);
    const tabs = [...document.querySelectorAll("[role=tab]")];
    tabs.forEach((t, i) => {
      t.addEventListener("click", () => showTab(t.dataset.tab));
      t.addEventListener("keydown", (e) => {
        const d = e.key === "ArrowRight" ? 1 : e.key === "ArrowLeft" ? -1 : 0;
        if (!d) return;
        const next = tabs[(i + d + tabs.length) % tabs.length];
        showTab(next.dataset.tab);
        next.focus();
      });
    });
    let tab = "update";
    try { tab = localStorage.getItem("tc-tab") || tab; } catch (e) { /* storage unavailable */ }
    showTab(["update", "status", "stickers"].includes(tab) ? tab : "update");
    renderAll();
    if (desktop.on) {
      apiJson("/api/status").then((st) => {
        desktop.sharepoint = !!st.sharepoint;
        desktop.name = st.name || "";
        desktop.output = st.output || "";
        desktop.url = st.url || "";
        if (!desktop.sharepoint) return;
        $("#sp-row").hidden = false;
        $("#sp-name").textContent = desktop.name ? `SharePoint: ${desktop.name}` : "SharePoint 파일";
        renderActions();
        downloadFromSharePoint(); // start right away: the user only has to add photos
      });
      $("#sp-download").addEventListener("click", downloadFromSharePoint);
    }
    // warm up the OCR engine in the background so the first photo is quicker
    setTimeout(async () => {
      const el = $("#engine-status");
      el.textContent = "글자 인식 엔진 준비 중…";
      const ok = await V.ensureEngine();
      el.textContent = ok
        ? "글자 인식 준비됨 · QR 스티커와 손글씨 메모 모두 읽습니다"
        : `글자 인식을 쓸 수 없어 QR 스티커와 직접 선택만 사용합니다${V.engine.error ? ` (${V.engine.error.message})` : ""}`;
      el.classList.toggle("bad", !ok);
    }, 800);
  }

  TC.app = { state, processFiles, loadExcel, startDemo };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
