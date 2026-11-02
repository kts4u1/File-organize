/* Office Script export: apply the reviewed updates inside Excel for the web.
 *
 * The generated script runs on the live SharePoint workbook (Automate tab),
 * so nothing is downloaded or replaced and co-workers' edits are kept.
 * It finds each tool by serial in its regional sheet, then either copies the
 * row to a TEST sheet and updates the copy (TEST mode) or updates the row
 * itself and swaps its pictures (direct mode).
 */
(function () {
  "use strict";
  const TC = (window.TC = window.TC || {});

  const DAY = 86400000;
  const EPOCH = Date.UTC(1899, 11, 30);
  const excelSerial = (iso) => {
    if (!iso) return 0;
    const [y, m, d] = iso.split("-").map(Number);
    return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / DAY);
  };

  function base64Of(canvas, type, maxSide, quality) {
    const { canvas: c } = TC.vision.scaled(canvas, maxSide);
    return c.toDataURL(type, quality).split(",")[1];
  }

  /** One item per tool: a photo and a SMAX capture of the same tool are merged. */
  function collect(jobs, today) {
    const items = new Map();
    for (const job of jobs) {
      if (!TC.match.applicable(job)) continue;
      const rec = job.target;
      if (!items.has(rec.key)) {
        items.set(rec.key, {
          sheet: rec.sheet, row: rec.row, serial: rec.serial,
          calLast: 0, calNext: 0, label: "", smax: "",
          memo: `원본: ${rec.sheet} ${rec.row}행 · ${today} 시험 반영`,
        });
      }
      const it = items.get(rec.key);
      if (job.kind === "tool_photo") {
        it.calLast = excelSerial(job.calLast);
        it.calNext = excelSerial(job.calNext);
        it.label = base64Of(job.labelCrop || job.image, "image/jpeg", 800, 0.82);
      } else if (job.kind === "smax_screenshot") {
        it.smax = base64Of(job.image, "image/png", 1400);
      }
    }
    return [...items.values()];
  }

  function build(jobs, testMode) {
    const today = TC.xlsx.todayIso();
    const items = collect(jobs, today);
    const mode = testMode ? "TEST" : "DIRECT";
    const list = items.map((it) => "    " + JSON.stringify(it)).join(",\n");
    const text = `/**
 * 툴 검교정 정리 — Excel 웹 반영 스크립트 (${today}, ${items.length}건)
 * 방식: ${testMode ? "지역 시트는 그대로 두고 'TEST' 탭에 복사본으로 반영" : "지역 시트의 해당 행에 바로 반영 (사진 교체)"}
 * 사용법: Excel 웹에서 이 파일 열기 → [자동화] → [새 스크립트] → 내용을 모두 지우고 붙여넣기 → [실행]
 * 결과는 스크립트 창 아래 '출력'에 행별로 표시됩니다.
 */
function main(workbook: ExcelScript.Workbook) {
  const MODE: string = "${mode}"; // "TEST" 또는 "DIRECT"
  const items: Item[] = [
${list}
  ];
  let test: ExcelScript.Worksheet | undefined = undefined;
  let nextRow = 0;
  let done = 0;
  for (const it of items) {
    const src = workbook.getWorksheet(it.sheet);
    if (!src) {
      console.log(\`건너뜀: '\${it.sheet}' 시트가 없습니다 (S/N \${it.serial})\`);
      continue;
    }
    const layout = readLayout(src);
    if (!layout) {
      console.log(\`건너뜀: '\${it.sheet}' 시트에서 Serial/Cal Last/Cal Next 헤더를 찾지 못했습니다\`);
      continue;
    }
    const srcRow = findRow(src, layout, it);
    if (srcRow < 0) {
      console.log(\`건너뜀: '\${it.sheet}' 시트에서 S/N \${it.serial}을 찾지 못했습니다\`);
      continue;
    }
    let dest: ExcelScript.Worksheet = src;
    let destRow = srcRow;
    if (MODE === "TEST") {
      if (!test) {
        const prepared = prepareTest(workbook, src, layout);
        test = prepared.sheet;
        nextRow = prepared.nextRow;
      }
      test.getRangeByIndexes(nextRow, 0, 1, layout.width)
        .copyFrom(src.getRangeByIndexes(srcRow, 0, 1, layout.width), ExcelScript.RangeCopyType.all, false, false);
      test.getCell(nextRow, 0).getFormat().setRowHeight(src.getCell(srcRow, 0).getFormat().getRowHeight());
      test.getCell(nextRow, layout.width).setValue(it.memo);
      dest = test;
      destRow = nextRow;
      nextRow++;
    }
    if (it.calLast > 0) writeDate(dest.getCell(destRow, layout.calLast), it.calLast);
    if (it.calNext > 0) writeDate(dest.getCell(destRow, layout.calNext), it.calNext);
    const replace = MODE !== "TEST";
    if (it.label !== "" && layout.label >= 0) putPicture(dest, dest.getCell(destRow, layout.label), it.label, \`Label \${it.serial}\`, replace);
    if (it.smax !== "" && layout.smax >= 0) putPicture(dest, dest.getCell(destRow, layout.smax), it.smax, \`SMAX \${it.serial}\`, replace);
    done++;
    console.log(\`반영: \${dest.getName()} \${destRow + 1}행 ← \${it.sheet} \${srcRow + 1}행 (S/N \${it.serial})\`);
  }
  if (test) test.activate();
  console.log(\`완료: \${done}/\${items.length}건\`);
}

interface Item {
  sheet: string; // 지역 시트 이름
  row: number; // 원본 행 번호 (1부터)
  serial: string;
  calLast: number; // 엑셀 날짜 숫자, 0이면 바꾸지 않음
  calNext: number;
  label: string; // 교정필증 사진 (base64 JPEG), 없으면 ""
  smax: string; // SMAX 캡쳐 (base64 PNG), 없으면 ""
  memo: string;
}

interface Layout {
  header: number; // 0부터 센 헤더 행
  width: number; // A열부터 마지막 헤더 열까지의 열 수
  serial: number;
  calLast: number;
  calNext: number;
  label: number;
  smax: number;
}

function norm(v: string | number | boolean): string {
  return String(v).toUpperCase().replace(/[^0-9A-Z]/g, "");
}

function isBlank(range: ExcelScript.Range | undefined): boolean {
  return !range || (range.getRowCount() === 1 && range.getColumnCount() === 1 && String(range.getValue()) === "");
}

function readLayout(ws: ExcelScript.Worksheet): Layout | undefined {
  const used = ws.getUsedRange(true);
  if (!used || isBlank(used)) return undefined;
  const values = used.getValues();
  const r0 = used.getRowIndex();
  const c0 = used.getColumnIndex();
  const find = (row: (string | number | boolean)[], names: string[]): number => {
    for (let c = 0; c < row.length; c++) {
      const h = String(row[c]).replace(/\\s+/g, " ").trim().toLowerCase();
      if (names.indexOf(h) >= 0) return c0 + c;
    }
    return -1;
  };
  for (let i = 0; i < Math.min(values.length, 10); i++) {
    const row = values[i];
    const serial = find(row, ["serial", "serial #", "serial no", "s/n", "시리얼", "시리얼번호"]);
    const calLast = find(row, ["cal last", "last cal", "cal date", "교정일"]);
    const calNext = find(row, ["cal next", "next cal", "cal due", "due date", "차기교정일", "차기 교정일"]);
    if (serial >= 0 && calLast >= 0 && calNext >= 0) {
      return {
        header: r0 + i,
        width: c0 + row.length,
        serial: serial,
        calLast: calLast,
        calNext: calNext,
        label: find(row, ["label photo", "라벨 사진", "교정필증"]),
        smax: find(row, ["smax tool photo", "smax photo", "smax"]),
      };
    }
  }
  return undefined;
}

function findRow(ws: ExcelScript.Worksheet, layout: Layout, it: Item): number {
  const want = norm(it.serial);
  const at = it.row - 1;
  if (at > layout.header && norm(ws.getCell(at, layout.serial).getValue()) === want) return at;
  const used = ws.getUsedRange(true);
  if (!used) return -1;
  const values = used.getValues();
  const r0 = used.getRowIndex();
  const c0 = used.getColumnIndex();
  for (let i = 0; i < values.length; i++) {
    const r = r0 + i;
    if (r > layout.header && norm(values[i][layout.serial - c0]) === want) return r;
  }
  return -1;
}

function prepareTest(workbook: ExcelScript.Workbook, src: ExcelScript.Worksheet, layout: Layout): { sheet: ExcelScript.Worksheet; nextRow: number } {
  const existing = workbook.getWorksheet("TEST");
  if (existing) {
    const used = existing.getUsedRange(true);
    if (used && !isBlank(used)) return { sheet: existing, nextRow: used.getRowIndex() + used.getRowCount() };
  }
  const ws = existing ? existing : workbook.addWorksheet("TEST");
  ws.getRangeByIndexes(0, 0, 1, layout.width)
    .copyFrom(src.getRangeByIndexes(layout.header, 0, 1, layout.width), ExcelScript.RangeCopyType.all, false, false);
  ws.getCell(0, layout.width).setValue("TEST 메모");
  for (let c = 0; c < layout.width; c++) {
    ws.getCell(0, c).getFormat().setColumnWidth(src.getCell(0, c).getFormat().getColumnWidth());
  }
  ws.getCell(0, layout.width).getFormat().setColumnWidth(240);
  ws.getFreezePanes().freezeRows(1);
  return { sheet: ws, nextRow: 1 };
}

function writeDate(cell: ExcelScript.Range, serial: number) {
  cell.setValue(serial);
  if (!/[ymd]/i.test(cell.getNumberFormat())) cell.setNumberFormat("yyyy-mm-dd");
}

function putPicture(ws: ExcelScript.Worksheet, cell: ExcelScript.Range, base64: string, name: string, replace: boolean) {
  const left = cell.getLeft();
  const top = cell.getTop();
  const width = cell.getWidth();
  const height = cell.getHeight();
  if (replace) {
    const old: ExcelScript.Shape[] = [];
    for (const s of ws.getShapes()) {
      if (s.getType() !== ExcelScript.ShapeType.image) continue;
      const x = s.getLeft();
      const y = s.getTop();
      if (x >= left - 2 && x < left + width && y >= top - 2 && y < top + height) old.push(s);
    }
    for (const s of old) s.delete();
  }
  const pic = ws.addImage(base64);
  pic.setName(name);
  const scale = Math.min((width - 4) / pic.getWidth(), (height - 4) / pic.getHeight());
  const w = pic.getWidth() * scale;
  const h = pic.getHeight() * scale;
  pic.setLockAspectRatio(false);
  pic.setWidth(w);
  pic.setHeight(h);
  pic.setLockAspectRatio(true);
  pic.setLeft(left + 2);
  pic.setTop(top + 2);
  pic.setPlacement(ExcelScript.Placement.oneCell);
}
`;
    return { text, count: items.length, size: new Blob([text]).size };
  }

  TC.officeScript = { build };
})();
