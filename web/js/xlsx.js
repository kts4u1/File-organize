/* Tool Verification workbook: read tool rows, write dates and pictures.
 *
 * Browser port of toolcal/workbook.py + xlsx_package.py. The .xlsx is edited
 * at the XML level with JSZip: only the parts that change are re-serialised,
 * everything else (SharePoint customXml, filters, other pictures) is kept.
 * Dates are ISO strings "YYYY-MM-DD" throughout.
 */
(function () {
  "use strict";
  const TC = (window.TC = window.TC || {});

  const NS = {
    main: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    rel: "http://schemas.openxmlformats.org/package/2006/relationships",
    ct: "http://schemas.openxmlformats.org/package/2006/content-types",
    xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  };
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const REL_WORKSHEET = REL + "/worksheet";
  const REL_DRAWING = REL + "/drawing";
  const REL_IMAGE = REL + "/image";
  const REL_STYLES = REL + "/styles";
  const REL_SST = REL + "/sharedStrings";
  const CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml";
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

  const HEADER_ALIASES = {
    cal_next: ["cal next", "next cal", "cal due", "due date", "차기교정일", "차기 교정일"],
    cal_last: ["cal last", "last cal", "cal date", "교정일"],
    manufacturer: ["manufacturer", "제조사"],
    model: ["model", "모델"],
    description: ["description", "설명", "품명"],
    serial: ["serial", "serial #", "serial no", "s/n", "시리얼", "시리얼번호"],
    asset: ["asset #", "asset", "asset no", "자산번호", "담당자"],
    interval: ["cal interval", "interval", "교정주기"],
    cycle: ["cal cycle", "cycle", "주기단위"],
    label_photo: ["label photo", "라벨 사진", "교정필증"],
    smax_photo: ["smax tool photo", "smax photo", "smax"],
  };
  const REQUIRED = ["cal_next", "cal_last", "serial", "asset"];
  const NON_NAME_WORDS = new Set(("공통 등록 메디아나 사무실 폐기 예비 서울 부산 대구 인천 광주 대전 울산 세종 경기 " +
    "강원 충북 충남 전북 전남 경북 경남 제주").split(" "));
  const BUILTIN_DATE_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47,
    27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58]);

  const MIN_PICTURE_ROW_PT = 60;
  const RAISED_ROW_PT = 100;
  const PAD_PX = 3;
  const EMU_PX = 9525;
  const EMU_PT = 12700;
  const MDW = 7; // max digit width (px) of Calibri 11

  // ------------------------------------------------------------ helpers
  const posix = {
    dirname: (p) => p.slice(0, Math.max(0, p.lastIndexOf("/"))),
    basename: (p) => p.slice(p.lastIndexOf("/") + 1),
    normalize(p) {
      const out = [];
      for (const seg of p.split("/")) {
        if (!seg || seg === ".") continue;
        if (seg === "..") out.pop();
        else out.push(seg);
      }
      return out.join("/");
    },
    resolve(fromPart, target) {
      if (target.startsWith("/")) return posix.normalize(target);
      return posix.normalize(posix.dirname(fromPart) + "/" + target);
    },
    relative(fromPart, toPart) {
      const from = posix.dirname(fromPart).split("/").filter(Boolean);
      const to = toPart.split("/");
      let i = 0;
      while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
      return [...Array(from.length - i).fill(".."), ...to.slice(i)].join("/");
    },
  };

  function colToIndex(letters) {
    let n = 0;
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }
  function indexToCol(n) {
    let s = "";
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }
  const kids = (el, ns, name) => Array.from(el.children).filter((c) => c.namespaceURI === ns && c.localName === name);
  const kid = (el, ns, name) => kids(el, ns, name)[0] || null;

  function isDateFormat(code) {
    if (!code) return false;
    const stripped = code.replace(/"[^"]*"/g, "").replace(/\[[^\]]*\]/g, "").replace(/\\./g, "");
    if (/general/i.test(stripped)) return false;
    return /[dmyhs]/i.test(stripped);
  }

  // ------------------------------------------------------------ dates
  const DAY = 86400000;
  const EPOCH_1900 = Date.UTC(1899, 11, 30);
  const EPOCH_1904 = Date.UTC(1904, 0, 1);
  function isoFromParts(y, m, d) {
    const t = Date.UTC(y, m - 1, d);
    const dt = new Date(t);
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
    return dt.toISOString().slice(0, 10);
  }
  function parseDate(value) {
    if (value == null) return null;
    const text = String(value).trim();
    let m = text.match(/(\d{4})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})/);
    if (m) return isoFromParts(+m[1], +m[2], +m[3]);
    m = text.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); // M/D/YYYY (SMAX)
    if (m) return isoFromParts(+m[3], +m[1], +m[2]);
    m = text.match(/(?:^|\D)(\d{2})[.\-](\d{1,2})[.\-](\d{1,2})(?:\D|$)/); // YY.MM.DD
    if (m) return isoFromParts(2000 + +m[1], +m[2], +m[3]);
    return null;
  }
  function addMonths(iso, months) {
    const [y, m, d] = iso.split("-").map(Number);
    const total = y * 12 + (m - 1) + months;
    const ny = Math.floor(total / 12);
    const nm = (total % 12) + 1;
    const last = new Date(Date.UTC(ny, nm, 0)).getUTCDate();
    return isoFromParts(ny, nm, Math.min(d, last));
  }
  function daysBetween(a, b) {
    return Math.round((Date.parse(b) - Date.parse(a)) / DAY);
  }
  function todayIso() {
    const n = new Date();
    return isoFromParts(n.getFullYear(), n.getMonth() + 1, n.getDate());
  }

  function intervalMonths(interval, cycle) {
    const n = parseInt(String(interval ?? "").trim(), 10);
    if (!Number.isFinite(n)) return null;
    const unit = String(cycle || "MONTHS").trim().toUpperCase();
    if (unit.startsWith("Y") || unit.startsWith("년")) return n * 12;
    if (unit.startsWith("M") || unit.startsWith("개월") || unit.startsWith("월")) return n;
    return null;
  }

  function extractNames(asset, exclude) {
    const names = [];
    for (let token of String(asset).match(/[가-힣]{2,4}/g) || []) {
      if (token.endsWith("님")) token = token.slice(0, -1);
      if (token.length >= 2 && !NON_NAME_WORDS.has(token) && !exclude.has(token) && !names.includes(token)) names.push(token);
    }
    return names;
  }

  // ------------------------------------------------------------ workbook
  class ToolWorkbook {
    static async open(data, fileName) {
      const wb = new ToolWorkbook();
      wb.fileName = fileName || "Tool_Verification.xlsx";
      wb.zip = await JSZip.loadAsync(data);
      wb.docs = new Map();
      wb.dirty = new Set();
      await wb._load();
      return wb;
    }

    // ---------------------------------------------------- package access
    has(name) {
      return !!this.zip.file(name);
    }
    async xml(name) {
      if (!this.docs.has(name)) {
        const text = await this.zip.file(name).async("string");
        const doc = new DOMParser().parseFromString(text, "application/xml");
        if (doc.getElementsByTagName("parsererror").length) throw new Error(`XML 오류: ${name}`);
        this.docs.set(name, doc);
      }
      return this.docs.get(name);
    }
    touch(name) {
      this.dirty.add(name);
    }
    relsName(part) {
      return `${posix.dirname(part)}/_rels/${posix.basename(part)}.rels`.replace(/^\//, "");
    }
    async rels(part, create) {
      const name = this.relsName(part);
      if (!this.has(name) && !this.docs.has(name)) {
        if (!create) return null;
        const doc = new DOMParser().parseFromString(`<Relationships xmlns="${NS.rel}"/>`, "application/xml");
        this.docs.set(name, doc);
        this.touch(name);
        return doc;
      }
      return this.xml(name);
    }
    async relTargets(part) {
      const out = new Map();
      const doc = await this.rels(part, false);
      if (!doc) return out;
      for (const rel of kids(doc.documentElement, NS.rel, "Relationship")) {
        if (rel.getAttribute("TargetMode") === "External") continue;
        out.set(rel.getAttribute("Id"), { type: rel.getAttribute("Type"), part: posix.resolve(part, rel.getAttribute("Target")) });
      }
      return out;
    }
    async addRel(part, type, target) {
      const doc = await this.rels(part, true);
      const used = new Set(kids(doc.documentElement, NS.rel, "Relationship").map((r) => r.getAttribute("Id")));
      let n = 1;
      while (used.has(`rId${n}`)) n++;
      const rel = doc.createElementNS(NS.rel, "Relationship");
      rel.setAttribute("Id", `rId${n}`);
      rel.setAttribute("Type", type);
      rel.setAttribute("Target", posix.relative(part, target));
      doc.documentElement.appendChild(rel);
      this.touch(this.relsName(part));
      return `rId${n}`;
    }
    async removeRel(part, rid) {
      const doc = await this.rels(part, false);
      if (!doc) return;
      for (const rel of kids(doc.documentElement, NS.rel, "Relationship")) {
        if (rel.getAttribute("Id") === rid) {
          rel.remove();
          this.touch(this.relsName(part));
        }
      }
    }
    async isReferenced(target) {
      for (const name of Object.keys(this.zip.files).concat([...this.docs.keys()])) {
        if (!name.endsWith(".rels") || this.zip.files[name]?.dir) continue;
        if (!this.has(name) && !this.docs.has(name)) continue;
        const dir = posix.dirname(name);
        const source = `${posix.dirname(dir)}/${posix.basename(name).slice(0, -5)}`.replace(/^\//, "");
        for (const { part } of (await this.relTargets(source)).values()) if (part === target) return true;
      }
      return false;
    }
    async ensureDefaultContentType(ext, type) {
      const doc = await this.xml("[Content_Types].xml");
      const root = doc.documentElement;
      for (const el of kids(root, NS.ct, "Default")) {
        if ((el.getAttribute("Extension") || "").toLowerCase() === ext) return;
      }
      const el = doc.createElementNS(NS.ct, "Default");
      el.setAttribute("Extension", ext);
      el.setAttribute("ContentType", type);
      const firstOverride = kid(root, NS.ct, "Override");
      root.insertBefore(el, firstOverride);
      this.touch("[Content_Types].xml");
    }
    async addOverride(part, type) {
      const doc = await this.xml("[Content_Types].xml");
      const el = doc.createElementNS(NS.ct, "Override");
      el.setAttribute("PartName", "/" + part);
      el.setAttribute("ContentType", type);
      doc.documentElement.appendChild(el);
      this.touch("[Content_Types].xml");
    }
    uniquePartName(dir, stem, ext) {
      let highest = 0;
      const re = new RegExp(`^${stem}(\\d+)\\.`, "i");
      for (const name of Object.keys(this.zip.files)) {
        if (posix.dirname(name) !== dir) continue;
        const m = posix.basename(name).match(re);
        if (m) highest = Math.max(highest, +m[1]);
      }
      return `${dir}/${stem}${highest + 1}.${ext}`;
    }

    // ---------------------------------------------------- read
    async _load() {
      const wbDoc = await this.xml("xl/workbook.xml");
      const pr = wbDoc.getElementsByTagNameNS(NS.main, "workbookPr")[0];
      this.date1904 = pr && ["1", "true"].includes(pr.getAttribute("date1904"));
      const targets = await this.relTargets("xl/workbook.xml");
      this.stylesPart = null;
      let sstPart = null;
      for (const t of targets.values()) {
        if (t.type === REL_STYLES) this.stylesPart = t.part;
        if (t.type === REL_SST) sstPart = t.part;
      }
      this.sharedStrings = [];
      if (sstPart && this.has(sstPart)) {
        const sst = await this.xml(sstPart);
        for (const si of kids(sst.documentElement, NS.main, "si")) {
          let text = "";
          const walk = (el) => {
            for (const c of el.children) {
              if (c.localName === "rPh") continue; // phonetic guide text
              if (c.localName === "t") text += c.textContent;
              else walk(c);
            }
          };
          walk(si);
          this.sharedStrings.push(text);
        }
        this.docs.delete(sstPart); // read-only; free memory
      }
      await this._loadStyles();

      this.sheets = new Map();
      this.records = [];
      for (const sheet of wbDoc.getElementsByTagNameNS(NS.main, "sheet")) {
        const t = targets.get(sheet.getAttributeNS(NS.r, "id"));
        if (!t || t.type !== REL_WORKSHEET || !this.has(t.part)) continue;
        await this._readSheet(sheet.getAttribute("name"), t.part);
      }
      if (!this.sheets.size) {
        throw new Error("툴 목록 시트를 찾지 못했습니다. 헤더에 'Serial', 'Asset #', 'Cal Next', 'Cal Last' 열이 있어야 합니다.");
      }
      const sheetNames = new Set(this.sheets.keys());
      for (const rec of this.records) rec.names = extractNames(rec.asset, sheetNames);
      for (const rec of this.records) {
        const info = this.sheets.get(rec.sheet);
        rec.pictures = {};
        for (const key of ["label_photo", "smax_photo"]) {
          if (info.columns[key]) rec.pictures[key] = (await this._anchorsAt(info, rec.row, info.columns[key])).length;
        }
      }
    }

    async _loadStyles() {
      this.xfNumFmt = [];
      this.customFmts = new Map();
      if (!this.stylesPart) return;
      const doc = await this.xml(this.stylesPart);
      const numFmts = doc.getElementsByTagNameNS(NS.main, "numFmts")[0];
      if (numFmts) for (const nf of kids(numFmts, NS.main, "numFmt")) this.customFmts.set(+nf.getAttribute("numFmtId"), nf.getAttribute("formatCode"));
      const xfs = doc.getElementsByTagNameNS(NS.main, "cellXfs")[0];
      if (xfs) this.xfNumFmt = kids(xfs, NS.main, "xf").map((xf) => +(xf.getAttribute("numFmtId") || 0));
    }

    styleIsDate(s) {
      if (s == null || s === "") return false;
      const fmt = this.xfNumFmt[+s];
      if (fmt == null) return false;
      if (BUILTIN_DATE_IDS.has(fmt)) return true;
      return this.customFmts.has(fmt) && isDateFormat(this.customFmts.get(fmt));
    }

    cellValue(c) {
      const t = c.getAttribute("t");
      const v = kid(c, NS.main, "v");
      if (t === "s") return v ? this.sharedStrings[+v.textContent] ?? "" : "";
      if (t === "inlineStr") {
        const is = kid(c, NS.main, "is");
        return is ? Array.from(is.getElementsByTagNameNS(NS.main, "t")).map((x) => x.textContent).join("") : "";
      }
      if (t === "str" || t === "e") return v ? v.textContent : "";
      if (t === "b") return v ? v.textContent === "1" : null;
      if (!v) return null;
      const num = Number(v.textContent);
      if (this.styleIsDate(c.getAttribute("s"))) {
        const base = this.date1904 ? EPOCH_1904 : EPOCH_1900;
        return { date: new Date(base + Math.round(num) * DAY).toISOString().slice(0, 10) };
      }
      return num;
    }

    async _readSheet(name, part) {
      const doc = await this.xml(part);
      const rows = [];
      for (const row of doc.getElementsByTagNameNS(NS.main, "row")) {
        const values = {};
        for (const c of kids(row, NS.main, "c")) {
          const m = (c.getAttribute("r") || "").match(/^([A-Z]+)/);
          if (m) values[m[1]] = this.cellValue(c);
        }
        rows.push({ r: +row.getAttribute("r"), values });
      }
      let header = null;
      let columns = null;
      for (const row of rows.slice(0, 10)) {
        const found = {};
        for (const [letter, value] of Object.entries(row.values)) {
          const h = String(value ?? "").replace(/\s+/g, " ").trim().toLowerCase();
          for (const [key, aliases] of Object.entries(HEADER_ALIASES)) {
            if (!found[key] && aliases.includes(h)) found[key] = letter;
          }
        }
        if (REQUIRED.every((k) => found[k])) {
          header = row.r;
          columns = found;
          break;
        }
      }
      if (!header) return;
      this.sheets.set(name, { name, part, headerRow: header, columns });
      const text = (v) => {
        if (v == null) return "";
        if (typeof v === "object" && v.date) return v.date;
        if (typeof v === "number") return Number.isInteger(v) ? String(v) : String(v);
        return String(v).trim();
      };
      for (const row of rows) {
        if (row.r <= header) continue;
        const get = (key) => (columns[key] ? row.values[columns[key]] : null);
        const serial = text(get("serial"));
        const asset = text(get("asset"));
        if (!serial && !asset) continue;
        const nextRaw = get("cal_next");
        const lastRaw = get("cal_last");
        const asDate = (v) => (v && typeof v === "object" && v.date ? v.date : typeof v === "string" ? parseDate(v) : null);
        this.records.push({
          sheet: name,
          row: row.r,
          key: `${name}!${row.r}`,
          calNext: asDate(nextRaw),
          calNextRaw: nextRaw,
          calLast: asDate(lastRaw),
          manufacturer: text(get("manufacturer")),
          model: text(get("model")),
          description: text(get("description")),
          serial,
          asset,
          intervalMonths: intervalMonths(text(get("interval")), text(get("cycle"))),
          intervalRaw: `${text(get("interval"))} ${text(get("cycle"))}`.trim(),
          disposed: typeof nextRaw === "string" && nextRaw.includes("폐기"),
          names: [],
        });
      }
    }

    findRecord(sheet, row) {
      return this.records.find((r) => r.sheet === sheet && r.row === row) || null;
    }

    // ---------------------------------------------------- cell writes
    async _rowEl(info, row, create) {
      const doc = await this.xml(info.part);
      const sheetData = doc.getElementsByTagNameNS(NS.main, "sheetData")[0];
      for (const el of kids(sheetData, NS.main, "row")) {
        const r = +el.getAttribute("r");
        if (r === row) return el;
        if (r > row) {
          if (!create) return null;
          const n = doc.createElementNS(NS.main, "row");
          n.setAttribute("r", String(row));
          sheetData.insertBefore(n, el);
          return n;
        }
      }
      if (!create) return null;
      const n = doc.createElementNS(NS.main, "row");
      n.setAttribute("r", String(row));
      sheetData.appendChild(n);
      return n;
    }

    async _cellEl(info, letter, row) {
      const rowEl = await this._rowEl(info, row, true);
      const doc = rowEl.ownerDocument;
      const target = colToIndex(letter);
      for (const c of kids(rowEl, NS.main, "c")) {
        const m = (c.getAttribute("r") || "").match(/^([A-Z]+)/);
        const idx = m ? colToIndex(m[1]) : 0;
        if (idx === target) return c;
        if (idx > target) {
          const n = doc.createElementNS(NS.main, "c");
          n.setAttribute("r", `${letter}${row}`);
          rowEl.insertBefore(n, c);
          return n;
        }
      }
      const n = doc.createElementNS(NS.main, "c");
      n.setAttribute("r", `${letter}${row}`);
      rowEl.appendChild(n);
      return n;
    }

    async _dateStyleFor(info, letter) {
      const doc = await this.xml(info.part);
      const re = new RegExp(`^${letter}\\d+$`);
      for (const c of doc.getElementsByTagNameNS(NS.main, "c")) {
        if (re.test(c.getAttribute("r") || "") && this.styleIsDate(c.getAttribute("s"))) return c.getAttribute("s");
      }
      for (let i = 0; i < this.xfNumFmt.length; i++) if (this.styleIsDate(String(i))) return String(i);
      const styles = await this.xml(this.stylesPart);
      const xfs = styles.getElementsByTagNameNS(NS.main, "cellXfs")[0];
      const xf = styles.createElementNS(NS.main, "xf");
      for (const [k, v] of [["numFmtId", "14"], ["fontId", "0"], ["fillId", "0"], ["borderId", "0"], ["xfId", "0"], ["applyNumberFormat", "1"]]) xf.setAttribute(k, v);
      xfs.appendChild(xf);
      xfs.setAttribute("count", String(kids(xfs, NS.main, "xf").length));
      this.xfNumFmt.push(14);
      this.touch(this.stylesPart);
      return String(this.xfNumFmt.length - 1);
    }

    /** Returns "written" | "unchanged" | "formula". */
    async writeDate(rec, key, iso) {
      const info = this.sheets.get(rec.sheet);
      const letter = info.columns[key];
      const c = await this._cellEl(info, letter, rec.row);
      if (kid(c, NS.main, "f")) {
        const wb = await this.xml("xl/workbook.xml");
        let calc = wb.getElementsByTagNameNS(NS.main, "calcPr")[0];
        if (!calc) {
          calc = wb.createElementNS(NS.main, "calcPr");
          wb.documentElement.appendChild(calc);
        }
        calc.setAttribute("fullCalcOnLoad", "1");
        this.touch("xl/workbook.xml");
        return "formula";
      }
      const [y, m, d] = iso.split("-").map(Number);
      const serial = String(Math.round((Date.UTC(y, m - 1, d) - (this.date1904 ? EPOCH_1904 : EPOCH_1900)) / DAY));
      const v = kid(c, NS.main, "v");
      if (!c.hasAttribute("t") && v && v.textContent === serial && this.styleIsDate(c.getAttribute("s"))) return "unchanged";
      while (c.firstChild) c.removeChild(c.firstChild);
      c.removeAttribute("t");
      if (!this.styleIsDate(c.getAttribute("s"))) c.setAttribute("s", await this._dateStyleFor(info, letter));
      const nv = c.ownerDocument.createElementNS(NS.main, "v");
      nv.textContent = serial;
      c.appendChild(nv);
      this.touch(info.part);
      if (key === "cal_next") {
        rec.calNext = iso;
        rec.calNextRaw = { date: iso };
        rec.disposed = false;
      } else if (key === "cal_last") rec.calLast = iso;
      return "written";
    }

    // ---------------------------------------------------- geometry
    async _colWidthPx(info, colIdx) {
      const doc = await this.xml(info.part);
      const cols = doc.getElementsByTagNameNS(NS.main, "cols")[0];
      if (cols) {
        for (const col of kids(cols, NS.main, "col")) {
          if (+col.getAttribute("min") <= colIdx && colIdx <= +col.getAttribute("max")) {
            if (["1", "true"].includes(col.getAttribute("hidden"))) return 0;
            const w = parseFloat(col.getAttribute("width") || "8.43");
            return Math.floor(((256 * w + Math.floor(128 / MDW)) / 256) * MDW);
          }
        }
      }
      const fmt = doc.getElementsByTagNameNS(NS.main, "sheetFormatPr")[0];
      if (fmt && fmt.getAttribute("defaultColWidth")) return Math.floor(parseFloat(fmt.getAttribute("defaultColWidth")) * MDW + 5);
      return 64;
    }
    async _defaultRowPt(info) {
      const doc = await this.xml(info.part);
      const fmt = doc.getElementsByTagNameNS(NS.main, "sheetFormatPr")[0];
      return fmt && fmt.getAttribute("defaultRowHeight") ? parseFloat(fmt.getAttribute("defaultRowHeight")) : 15;
    }
    async _rowHeightPt(info, row) {
      const el = await this._rowEl(info, row, false);
      if (el) {
        if (["1", "true"].includes(el.getAttribute("hidden"))) return 0;
        if (el.getAttribute("ht")) return parseFloat(el.getAttribute("ht"));
      }
      return this._defaultRowPt(info);
    }
    async _absOffsetEmu(info, colIdx, row) {
      let x = 0;
      for (let c = 1; c < colIdx; c++) x += await this._colWidthPx(info, c);
      const doc = await this.xml(info.part);
      const byRow = new Map();
      for (const el of doc.getElementsByTagNameNS(NS.main, "row")) byRow.set(+el.getAttribute("r"), el);
      const def = await this._defaultRowPt(info);
      let y = 0;
      for (let r = 1; r < row; r++) {
        const el = byRow.get(r);
        if (el && ["1", "true"].includes(el.getAttribute("hidden"))) continue;
        y += el && el.getAttribute("ht") ? parseFloat(el.getAttribute("ht")) : def;
      }
      return [x * EMU_PX, Math.round(y * EMU_PT)];
    }

    // ---------------------------------------------------- pictures
    async _drawingPart(info, create) {
      const doc = await this.xml(info.part);
      const drawing = doc.getElementsByTagNameNS(NS.main, "drawing")[0];
      if (drawing) {
        const t = (await this.relTargets(info.part)).get(drawing.getAttributeNS(NS.r, "id"));
        if (t && t.type === REL_DRAWING) return t.part;
      }
      if (!create) return null;
      const part = this.uniquePartName("xl/drawings", "drawing", "xml");
      const dd = new DOMParser().parseFromString(`<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}"/>`, "application/xml");
      this.docs.set(part, dd);
      this.zip.file(part, "", { createFolders: false }); // placeholder so the name counts as taken
      this.touch(part);
      await this.addOverride(part, CT_DRAWING);
      const rid = await this.addRel(info.part, REL_DRAWING, part);
      const el = doc.createElementNS(NS.main, "drawing");
      el.setAttributeNS(NS.r, "r:id", rid);
      const follows = ["legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];
      const before = Array.from(doc.documentElement.children).find((c) => follows.includes(c.localName));
      doc.documentElement.insertBefore(el, before || null);
      this.touch(info.part);
      return part;
    }

    async _anchorsAt(info, row, letter) {
      const part = await this._drawingPart(info, false);
      if (!part) return [];
      const doc = await this.xml(part);
      const col0 = colToIndex(letter) - 1;
      const row0 = row - 1;
      return Array.from(doc.documentElement.children).filter((a) => {
        if (a.namespaceURI !== NS.xdr || !["twoCellAnchor", "oneCellAnchor"].includes(a.localName)) return false;
        if (!kid(a, NS.xdr, "pic")) return false;
        const from = kid(a, NS.xdr, "from");
        return from && +kid(from, NS.xdr, "col").textContent === col0 && +kid(from, NS.xdr, "row").textContent === row0;
      });
    }

    /** Replace the pictures in the row's `key` column. Returns how many were removed. */
    async replacePicture(rec, key, bytes, ext, size, name, descr) {
      const info = this.sheets.get(rec.sheet);
      const letter = info.columns[key];
      if (!letter) throw new Error(`'${rec.sheet}' 시트에 사진 열이 없습니다.`);
      const drawing = await this._drawingPart(info, true);
      const dd = await this.xml(drawing);
      const removed = await this._anchorsAt(info, rec.row, letter);
      const oldRids = new Set();
      for (const anchor of removed) {
        for (const blip of anchor.getElementsByTagNameNS(NS.a, "blip")) {
          const rid = blip.getAttributeNS(NS.r, "embed");
          if (rid) oldRids.add(rid);
        }
        anchor.remove();
      }
      if (removed.length) {
        this.touch(drawing);
        const stillUsed = new Set();
        for (const el of dd.getElementsByTagName("*")) {
          for (const attr of ["embed", "link", "id"]) {
            const v = el.getAttributeNS(NS.r, attr);
            if (v) stillUsed.add(v);
          }
        }
        const targets = await this.relTargets(drawing);
        for (const rid of oldRids) {
          if (stillUsed.has(rid)) continue;
          const media = targets.get(rid)?.part;
          await this.removeRel(drawing, rid);
          if (media && this.has(media) && !(await this.isReferenced(media))) this.zip.remove(media);
        }
      }

      const media = this.uniquePartName("xl/media", "image", ext);
      this.zip.file(media, bytes, { createFolders: false });
      await this.ensureDefaultContentType(ext, ext === "jpeg" || ext === "jpg" ? "image/jpeg" : `image/${ext}`);
      const rid = await this.addRel(drawing, REL_IMAGE, media);

      const colIdx = colToIndex(letter);
      let rowPt = await this._rowHeightPt(info, rec.row);
      if (rowPt < MIN_PICTURE_ROW_PT) {
        const el = await this._rowEl(info, rec.row, true);
        el.setAttribute("ht", String(RAISED_ROW_PT));
        el.setAttribute("customHeight", "1");
        this.touch(info.part);
        rowPt = RAISED_ROW_PT;
      }
      const boxW = Math.max((await this._colWidthPx(info, colIdx)) - 2 * PAD_PX, 10);
      const boxH = Math.max(Math.floor((rowPt * 96) / 72) - 2 * PAD_PX, 10);
      const scale = Math.min(boxW / size[0], boxH / size[1]);
      const cx = Math.floor(size[0] * scale) * EMU_PX;
      const cy = Math.floor(size[1] * scale) * EMU_PX;
      const pad = PAD_PX * EMU_PX;
      const [ax, ay] = await this._absOffsetEmu(info, colIdx, rec.row);
      let shapeId = 1;
      for (const el of dd.getElementsByTagNameNS(NS.xdr, "cNvPr")) shapeId = Math.max(shapeId, +el.getAttribute("id") || 0);

      const X = (tag) => dd.createElementNS(NS.xdr, `xdr:${tag}`);
      const A = (tag) => dd.createElementNS(NS.a, `a:${tag}`);
      const add = (parent, child, text) => {
        if (text != null) child.textContent = String(text);
        parent.appendChild(child);
        return child;
      };
      const anchor = X("twoCellAnchor");
      anchor.setAttribute("editAs", "oneCell");
      for (const [tag, co, ro] of [["from", pad, pad], ["to", pad + cx, pad + cy]]) {
        const pos = add(anchor, X(tag));
        add(pos, X("col"), colIdx - 1);
        add(pos, X("colOff"), co);
        add(pos, X("row"), rec.row - 1);
        add(pos, X("rowOff"), ro);
      }
      const pic = add(anchor, X("pic"));
      const nv = add(pic, X("nvPicPr"));
      const cnv = add(nv, X("cNvPr"));
      cnv.setAttribute("id", String(shapeId + 1));
      cnv.setAttribute("name", name);
      if (descr) cnv.setAttribute("descr", descr);
      add(add(nv, X("cNvPicPr")), A("picLocks")).setAttribute("noChangeAspect", "1");
      const fill = add(pic, X("blipFill"));
      const blip = add(fill, A("blip"));
      blip.setAttributeNS(NS.r, "r:embed", rid); // the serializer declares xmlns:r where needed
      add(add(fill, A("stretch")), A("fillRect"));
      const spPr = add(pic, X("spPr"));
      const xfrm = add(spPr, A("xfrm"));
      const off = add(xfrm, A("off"));
      off.setAttribute("x", String(ax + pad));
      off.setAttribute("y", String(ay + pad));
      const extEl = add(xfrm, A("ext"));
      extEl.setAttribute("cx", String(cx));
      extEl.setAttribute("cy", String(cy));
      const geom = add(spPr, A("prstGeom"));
      geom.setAttribute("prst", "rect");
      add(geom, A("avLst"));
      add(anchor, X("clientData"));
      dd.documentElement.appendChild(anchor);
      this.touch(drawing);
      rec.pictures[key] = 1;
      return removed.length;
    }

    // ---------------------------------------------------- save
    async toBlob() {
      for (const name of this.dirty) {
        const doc = this.docs.get(name);
        if (!doc) continue;
        // Chrome keeps the parsed <?xml ... standalone="yes"?> declaration; add one only if missing.
        const text = new XMLSerializer().serializeToString(doc);
        this.zip.file(name, text.startsWith("<?xml") ? text : XML_DECL + text, { createFolders: false });
      }
      this.dirty.clear();
      return this.zip.generateAsync({
        type: "blob",
        compression: "DEFLATE",
        mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
    }
  }

  TC.xlsx = { ToolWorkbook, parseDate, addMonths, daysBetween, todayIso, indexToCol, colToIndex, isDateFormat };
})();
