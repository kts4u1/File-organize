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

  // ------------------------------------------------------------ downloads
  const inViewer = !!(window.claude && typeof window.claude.use === "function");
  const downloadsReady = inViewer ? window.claude.use("downloads").catch(() => null) : Promise.resolve(null);

  async function saveFile(blob, filename) {
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
  async function loadExcel(bytes, name, demo) {
    try {
      const wb = await ToolWorkbook.open(bytes.slice(0), name);
      state.excel = { bytes, name };
      state.wb = wb;
      state.demo = !!demo;
      state.jobs = [];
      state.result = null;
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

  function renderActions() {
    const n = state.jobs.filter(M.applicable).length;
    const btn = $("#apply");
    btn.textContent = `엑셀 파일 만들기 (${n}건)`;
    btn.disabled = !n || state.busy;
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
      const today = todayIso();
      for (const job of state.jobs) {
        if (!M.applicable(job)) continue;
        const rec = wb.findRecord(job.target.sheet, job.target.row);
        if (!rec || M.normalizeSerial(rec.serial) !== M.normalizeSerial(job.target.serial)) continue;
        const before = { calLast: rec.calLast, calNext: rec.calNext };
        const notes = [];
        let replaced = 0;
        if (job.kind === "tool_photo") {
          for (const [key, value] of [["cal_last", job.calLast], ["cal_next", job.calNext]]) {
            if ((await wb.writeDate(rec, key, value)) === "formula") notes.push(`${key === "cal_next" ? "Cal Next" : "Cal Last"} 열은 수식이라 그대로 두었습니다.`);
          }
          if (wb.sheets.get(rec.sheet).columns.label_photo) {
            const pic = await canvasBytes(job.labelCrop || job.image, "image/jpeg", 1000, 0.85);
            replaced = await wb.replacePicture(rec, "label_photo", pic.bytes, "jpeg", pic.size, `Label ${rec.serial}`, `${job.fileName} (${today})`);
          }
        } else if (wb.sheets.get(rec.sheet).columns.smax_photo) {
          const pic = await canvasBytes(job.image, "image/png", 1600);
          replaced = await wb.replacePicture(rec, "smax_photo", pic.bytes, "png", pic.size, `SMAX ${rec.serial}`, `${job.fileName} (${today})`);
        }
        changes.push({ job, rec, before, after: { calLast: rec.calLast, calNext: rec.calNext }, replaced, notes });
      }
      const blob = await wb.toBlob();
      const bytes = new Uint8Array(await blob.arrayBuffer());
      state.result = { blob, changes, name: state.excel.name };
      // the updated file becomes the working copy for the next batch
      state.excel = { bytes, name: state.excel.name };
      state.wb = await ToolWorkbook.open(bytes.slice(0), state.excel.name);
      state.jobs = [];
      renderAll();
      await offerResult();
    } catch (e) {
      console.error(e);
      toast(`엑셀을 만들지 못했습니다: ${e.message}`, "bad");
      renderActions();
    }
  }

  async function offerResult() {
    if (!state.result) return;
    const r = await saveFile(state.result.blob, state.result.name);
    if (r.ok) toast("엑셀 파일을 저장했습니다.", "good");
    else toast(r.message, "bad");
  }

  function renderResult() {
    const box = $("#result");
    box.replaceChildren();
    box.hidden = !state.result;
    if (!state.result) return;
    const { changes } = state.result;
    const rows = changes.map((c) => h("tr", {},
      h("td", {}, c.job.fileName),
      h("td", {}, M.label(c.rec)),
      h("td", {}, KIND[c.job.kind]),
      h("td", { class: "mono" }, c.job.kind === "tool_photo" ? `${c.before.calLast || "-"} → ${c.after.calLast}` : ""),
      h("td", { class: "mono" }, c.job.kind === "tool_photo" ? `${c.before.calNext || "-"} → ${c.after.calNext}` : ""),
      h("td", {}, c.replaced ? `교체 (${c.replaced}장 삭제)` : "추가", c.notes.length ? h("div", { class: "muted" }, c.notes.join(" ")) : null)));
    box.append(
      h("div", { class: "result-head" },
        h("h3", {}, `${changes.length}건을 반영한 엑셀을 만들었습니다`),
        h("div", { class: "row-gap" },
          h("button", { class: "primary", onclick: offerResult }, "엑셀 다시 저장"),
          h("button", { onclick: () => { state.result = null; renderAll(); } }, "닫기"))),
      h("p", { class: "muted" }, `파일 이름: ${state.result.name} — 회사 공유 폴더의 원본을 이 파일로 바꿔 주세요. 다음 사진은 이 파일에 이어서 반영됩니다.`),
      h("div", { class: "table-wrap" }, h("table", {},
        h("thead", {}, h("tr", {}, ...["사진", "대상", "종류", "교정일", "차기교정일", "사진"].map((t) => h("th", {}, t)))),
        h("tbody", {}, ...rows))));
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
    wireDrop($("#excel-drop"), $("#excel-input"), async (files) => {
      const f = files.find((x) => /\.xlsx$/i.test(x.name));
      if (!f) return toast("xlsx 파일을 선택하세요.", "bad");
      await loadExcel(new Uint8Array(await f.arrayBuffer()), f.name, false);
    });
    wireDrop($("#photo-drop"), $("#photo-input"), (files) => {
      const imgs = files.filter((f) => /^image\//.test(f.type) || /\.(jpe?g|png|heic|heif|webp|bmp)$/i.test(f.name));
      if (imgs.length) processFiles(imgs);
    });
    $("#demo").addEventListener("click", startDemo);
    $("#apply").addEventListener("click", applyJobs);
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
