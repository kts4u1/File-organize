/* Match what was read from a photo to a tool row, and validate the dates.
 * Browser port of toolcal/matcher.py and the checks in toolcal/pipeline.py.
 *
 * Offline OCR reads handwriting imperfectly, so serial readings are matched
 * against the known serials of the workbook (closed set): a reading that is
 * one or two characters off still finds its tool, and the user confirms.
 */
(function () {
  "use strict";
  const TC = (window.TC = window.TC || {});
  const { addMonths, daysBetween, todayIso } = TC.xlsx;

  const CONFUSABLE = { O: "0", Q: "0", D: "0", I: "1", L: "1", S: "5", B: "8", Z: "2", G: "6" };
  const normalizeSerial = (v) => String(v || "").toUpperCase().replace(/[^0-9A-Z]/g, "");
  const loose = (v) => normalizeSerial(v).replace(/[OQDILSBZG]/g, (c) => CONFUSABLE[c]).replace(/^0+/, "");

  function levenshtein(a, b) {
    if (a === b) return 0;
    let prev = Array.from({ length: b.length + 1 }, (_, i) => i);
    for (let i = 1; i <= a.length; i++) {
      const cur = [i];
      for (let j = 1; j <= b.length; j++) {
        cur.push(Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] !== b[j - 1] ? 1 : 0)));
      }
      prev = cur;
    }
    return prev[b.length];
  }

  /** 0 = same, 0.5 = same up to look-alike characters / leading zeros, n = edits. */
  function serialDistance(reading, serial) {
    const a = normalizeSerial(reading);
    const b = normalizeSerial(serial);
    if (!a || !b) return 99;
    if (a === b) return 0;
    const la = loose(a);
    const lb = loose(b);
    if (la === lb) return 0.5;
    return levenshtein(la, lb);
  }

  function nameRelation(name, rec) {
    name = String(name || "").replace(/\s+/g, "");
    if (!/[가-힣]/.test(name) || !rec.names.length) return "unknown";
    for (const known of rec.names) if (name === known || name.includes(known) || known.includes(name)) return "exact";
    for (const known of rec.names) {
      if (known.length === name.length) {
        let same = 0;
        for (let i = 0; i < name.length; i++) if (known[i] === name[i]) same++;
        if (same >= name.length - 1) return "partial";
      }
    }
    return "mismatch";
  }

  /**
   * reading: {source:'qr'|'ocr'|'manual', serials:[...], name, calLast, calNext}
   * Returns candidates sorted by score.
   */
  function findCandidates(reading, records) {
    const serials = [];
    for (const s of reading.serials || []) if (normalizeSerial(s).length >= 4 && !serials.includes(s)) serials.push(s);
    const votes = (reading.votes && reading.votes.length ? reading.votes : serials).filter((v) => normalizeSerial(v).length >= 4);
    const out = [];
    for (const rec of records) {
      let relation = nameRelation(reading.name, rec);
      // Offline OCR of handwritten Hangul is unreliable: a match helps, a miss says nothing.
      if (reading.source === "ocr" && relation === "mismatch") relation = "unknown";
      let best = null;
      serials.forEach((s, i) => {
        const d = serialDistance(s, rec.serial);
        const limit = normalizeSerial(rec.serial).length >= 7 ? 2 : 1;
        if (d > limit) return;
        let base = d === 0 ? 100 : d === 0.5 ? 95 : d === 1 ? 70 : 45;
        if (i > 0 && reading.source === "qr") base -= 5;
        if (!best || base > best.base) best = { base, d, s };
      });
      if (!best) {
        if (relation !== "exact" || serials.length) continue;
        best = { base: 30, d: 99, s: "" }; // no serial read: offer the engineer's tools
      }
      let score = best.base + { exact: 25, partial: 10, mismatch: -30, unknown: 0 }[relation];
      // Agreement across several readings of the same line (0..20 points).
      if (votes.length > 1 && best.d < 99) {
        let sum = 0;
        for (const v of votes) sum += Math.max(0, 3 - Math.min(3, serialDistance(v, rec.serial)));
        score += (sum / (3 * votes.length)) * 20;
      }
      // A label whose interval matches this tool type is weak extra evidence.
      if (reading.calLast && reading.calNext && rec.intervalMonths) {
        if (Math.abs(daysBetween(addMonths(reading.calLast, rec.intervalMonths), reading.calNext)) <= 3) score += 5;
      }
      if (rec.disposed) score -= 40;
      out.push({ record: rec, score, distance: best.d, reading: best.s, name: relation });
    }
    out.sort((a, b) => b.score - a.score);
    return out;
  }

  function describe(c) {
    let serial;
    if (c.distance === 0) serial = "시리얼 일치";
    else if (c.distance < 1) serial = `시리얼 유사(${c.reading}→${c.record.serial})`;
    else if (c.distance < 99) serial = `시리얼 ${c.distance}자 차이(${c.reading}→${c.record.serial})`;
    else serial = "시리얼 미확인";
    const name = { exact: "이름 일치", partial: "이름 유사", mismatch: "이름 불일치", unknown: "이름 확인 안 함" }[c.name];
    return `${serial}, ${name}`;
  }

  function label(rec) {
    return `[${rec.sheet} ${rec.row}행] ${rec.asset} · ${rec.model} · S/N ${rec.serial}`;
  }

  function decide(candidates, source) {
    if (!candidates.length) {
      return { candidate: null, status: "error", messages: ["엑셀에서 일치하는 툴을 찾지 못했습니다. 대상 행을 직접 선택하세요."] };
    }
    const top = candidates[0];
    const second = candidates[1];
    const msgs = [`매칭: ${label(top.record)} (${describe(top)})`];
    const clearLead = !second || second.score <= top.score - 20;
    if (top.record.disposed) return { candidate: top, status: "review", messages: msgs.concat("폐기 처리된 툴입니다. 확인 후 반영하세요.") };
    if (top.distance <= 0.5 && top.name !== "mismatch" && clearLead) {
      if (source === "qr") msgs[0] += " · QR";
      if (top.distance === 0.5) msgs.push("시리얼 표기 차이(0/O, 앞자리 0 등)는 자동 보정했습니다.");
      return { candidate: top, status: "ok", messages: msgs };
    }
    if (top.distance === 1 && top.name === "exact" && clearLead) {
      msgs.push("시리얼 한 글자가 다르지만 이름이 일치해 자동 보정했습니다.");
      return { candidate: top, status: "ok", messages: msgs };
    }
    if (top.name === "mismatch") msgs.push("사진의 이름과 엑셀 담당자가 다릅니다.");
    if (top.distance >= 99) msgs.push("시리얼을 읽지 못해 이름으로만 찾았습니다.");
    else if (top.distance >= 1) msgs.push("손글씨 시리얼이 정확히 읽히지 않아 가장 가까운 툴을 골랐습니다. 사진과 비교해 확인하세요.");
    const others = candidates.slice(1, 3).filter((c) => c.distance < 99);
    if (others.length && (!clearLead || top.distance >= 1)) msgs.push("다른 후보: " + others.map((c) => label(c.record)).join(" / "));
    return { candidate: top, status: "review", messages: msgs };
  }

  /** (Re)compute job.status / job.messages from its current target and dates. */
  function evaluate(job, today) {
    today = today || todayIso();
    const msgs = [];
    let status = "ok";
    const review = (m) => {
      msgs.push(m);
      if (status === "ok") status = "review";
    };
    if (job.error) {
      msgs.push(job.error);
      status = job.targetConfirmed ? "review" : "error";
    }
    if (job.kind === "other") {
      job.status = "skip";
      job.messages = msgs.concat("툴 사진이나 SMAX 캡쳐가 아니어서 제외했습니다.");
      return;
    }
    if (job.targetConfirmed && job.target) msgs.push(`대상(직접 선택): ${label(job.target)}`);
    else {
      msgs.push(...job.matchMessages);
      if (job.matchStatus === "error") status = "error";
      else if (job.matchStatus === "review") review("매칭 결과를 확인하세요.");
    }
    const rec = job.target;
    if (!rec) {
      job.status = "error";
      job.messages = msgs;
      return;
    }
    if (job.kind === "tool_photo") {
      const iv = rec.intervalMonths;
      if (job.calLast && !job.calNext && iv) {
        job.calNext = addMonths(job.calLast, iv);
        msgs.push(`차기교정일을 읽지 못해 교정주기(${rec.intervalRaw})로 계산했습니다.`);
      } else if (job.calNext && !job.calLast && iv) {
        job.calLast = addMonths(job.calNext, -iv);
        msgs.push(`교정일을 읽지 못해 교정주기(${rec.intervalRaw})로 계산했습니다.`);
      }
      if (!(job.calLast && job.calNext)) review("교정일/차기교정일을 입력하세요.");
      else {
        if (job.calNext <= job.calLast) review("차기교정일이 교정일보다 빠릅니다.");
        if (daysBetween(today, job.calLast) > 3) review("교정일이 미래 날짜입니다.");
        if (iv) {
          const expected = addMonths(job.calLast, iv);
          if (Math.abs(daysBetween(expected, job.calNext)) > 3) review(`교정주기(${rec.intervalRaw})와 차기교정일이 맞지 않습니다 (예상 ${expected}).`);
        }
        if (rec.calLast && job.calLast < rec.calLast) review(`엑셀의 기존 교정일(${rec.calLast})보다 이전 날짜입니다. 예전 사진인지 확인하세요.`);
        if (rec.calLast === job.calLast && rec.calNext === job.calNext) msgs.push("날짜는 엑셀과 같습니다 (사진만 교체).");
        else msgs.push(`날짜 변경: ${rec.calLast || "-"} → ${job.calLast}, 차기 ${rec.calNext || "-"} → ${job.calNext}`);
      }
      if (!job.labelCrop) msgs.push("교정필증을 찾지 못해 사진 전체를 넣습니다.");
    } else if (job.kind === "smax_screenshot") {
      msgs.push("SMAX 캡쳐를 'SMAX Tool Photo' 열에 넣습니다.");
    }
    job.status = status;
    job.messages = msgs;
  }

  function evaluateAll(jobs) {
    for (const job of jobs) evaluate(job);
    // same tool twice in one batch
    const groups = new Map();
    for (const job of jobs) {
      if (!job.target || job.kind === "other") continue;
      const k = `${job.target.key}|${job.kind}`;
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(job);
    }
    for (const group of groups.values()) {
      if (group.length < 2) continue;
      const names = group.map((j) => j.fileName).join(", ");
      for (const job of group) {
        job.messages.push(`같은 툴에 대한 사진이 여러 장입니다: ${names}`);
        if (job.status === "ok") job.status = "review";
      }
    }
    // SMAX screenshot showing a due date other than the one being written
    const planned = new Map();
    for (const j of jobs) if (j.kind === "tool_photo" && j.target && j.calNext && j.status !== "error") planned.set(j.target.key, j.calNext);
    for (const j of jobs) {
      if (j.kind !== "smax_screenshot" || !j.target || !j.reading?.calNext) continue;
      const expected = planned.get(j.target.key) || j.target.calNext;
      if (expected && expected !== j.reading.calNext) {
        j.messages.push(`참고: SMAX의 Due Date(${j.reading.calNext})가 엑셀에 들어갈 차기교정일(${expected})과 다릅니다. SMAX를 먼저 수정한 뒤 캡쳐했는지 확인하세요.`);
      }
    }
  }

  function applicable(job) {
    if (!job.include || !job.target || job.status === "error" || job.status === "skip") return false;
    if (job.kind === "tool_photo") return !!(job.calLast && job.calNext);
    return job.kind === "smax_screenshot";
  }

  function dueStatus(records, warnDays, today) {
    today = today || todayIso();
    const rows = records.map((rec) => {
      let state;
      let days = null;
      if (rec.disposed) state = "폐기";
      else if (!rec.calNext) state = "날짜 없음";
      else {
        days = daysBetween(today, rec.calNext);
        state = days < 0 ? "만료" : days <= warnDays ? "임박" : "정상";
      }
      return { state, days, rec };
    });
    const order = { 만료: 0, 임박: 1, 정상: 2, "날짜 없음": 3, 폐기: 4 };
    rows.sort((a, b) => order[a.state] - order[b.state] || (a.days ?? 1e6) - (b.days ?? 1e6));
    return rows;
  }

  TC.match = { normalizeSerial, serialDistance, nameRelation, findCandidates, decide, evaluate, evaluateAll, applicable, dueStatus, label };
})();
