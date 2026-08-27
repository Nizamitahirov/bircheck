import * as XLSX from "xlsx";

export type ProgressFn = (step: string, pct: number) => void;
export type Cell = string | number | boolean | Date | null | undefined;
type Row = Cell[];

export interface Order {
  type: string;
  start: number; // excel serial
  end: number; // excel serial
}

export interface CrossPair {
  code: string;
  a: Order;
  b: Order;
}

export interface CrossGroup {
  code: string;
  pairs: CrossPair[];
  orders: Order[]; // distinct orders involved in any crossing, sorted by start
}

export interface CrossResult {
  workbook: XLSX.WorkBook;
  sheetName: string;
  groups: CrossGroup[];
  pairs: CrossPair[];
  totalRecords: number;
  employeesWithCross: number;
  durationMs: number;
}

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;

function dateFromSerial(s: number): Date {
  return new Date(EXCEL_EPOCH_MS + Math.round(s) * MS_PER_DAY);
}

function excelSerialFromDate(d: Date): number {
  const utc = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  return Math.round((utc - EXCEL_EPOCH_MS) / MS_PER_DAY);
}

// Robust date parser: number → serial; Date → serial; "dd.mm.yyyy" (or / -) → serial.
function asSerial(v: Cell): number | null {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return Math.round(v);
  if (v instanceof Date) return excelSerialFromDate(v);
  const s = String(v).trim();
  if (!s) return null;
  const m = s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/);
  if (m) {
    const day = Number(m[1]);
    const month = Number(m[2]);
    const year = Number(m[3]);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return Math.round((Date.UTC(year, month - 1, day) - EXCEL_EPOCH_MS) / MS_PER_DAY);
    }
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : excelSerialFromDate(d);
}

function toKey(v: Cell): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return String(Math.trunc(v));
  return String(v).trim();
}

export function formatSerial(s: number): string {
  const d = dateFromSerial(s);
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  return `${dd}.${mm}.${d.getUTCFullYear()}`;
}

function sheetToRows(ws: XLSX.WorkSheet): Row[] {
  return XLSX.utils.sheet_to_json<Row>(ws, {
    header: 1,
    raw: true,
    defval: null,
    blankrows: false,
  });
}

// Two intervals cross when they overlap or touch on a boundary day. Since data
// is grouped per employee and sorted by start, a.start <= b.start, so the test
// reduces to b.start <= a.end — but we keep the symmetric form for safety.
function crosses(a: Order, b: Order): boolean {
  return a.start <= b.end && b.start <= a.end;
}

export async function processCrossCheck(
  file: File,
  onProgress: ProgressFn = () => {},
): Promise<CrossResult> {
  const t0 = performance.now();

  onProgress("Excel faylı oxunur...", 5);
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: false });

  // Prefer a "Dates" sheet (as in the macro); otherwise use the first sheet
  // that actually has data rows.
  let sheetName = wb.SheetNames.find((n) => n.toLowerCase() === "dates") ?? "";
  if (!sheetName) {
    for (const n of wb.SheetNames) {
      const rows = sheetToRows(wb.Sheets[n]);
      if (rows.length > 1) {
        sheetName = n;
        break;
      }
    }
  }
  if (!sheetName) throw new Error("Faylda məlumat olan sheet tapılmadı.");

  const rows = sheetToRows(wb.Sheets[sheetName]);
  if (rows.length < 2) throw new Error(`"${sheetName}" sheet-də məlumat yoxdur.`);

  onProgress("Məlumatlar oxunur...", 25);

  // Group orders per employee. A=code(0), B=type(1), C=start(2), D=end(3).
  const byCode = new Map<string, Order[]>();
  let totalRecords = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    const code = toKey(r[0]);
    if (!code) continue;
    let start = asSerial(r[2]);
    let end = asSerial(r[3]);
    if (start === null && end === null) continue;
    if (start === null) start = end!;
    if (end === null) end = start;
    if (start > end) [start, end] = [end, start];
    const type = r[1] === null || r[1] === undefined ? "" : String(r[1]).trim();
    const arr = byCode.get(code) ?? [];
    arr.push({ type, start, end });
    byCode.set(code, arr);
    totalRecords++;
  }

  onProgress("Kəsişən əmrlər axtarılır...", 55);

  // Detect crossing pairs per employee, sorted by badge then start.
  const groups: CrossGroup[] = [];
  const allPairs: CrossPair[] = [];
  const codes = Array.from(byCode.keys()).sort((x, y) =>
    x.localeCompare(y, undefined, { numeric: true }),
  );

  for (const code of codes) {
    const orders = byCode.get(code)!;
    // Sort by start asc, then end asc (matches filter_data: key1 A, key2 C).
    orders.sort((a, b) => a.start - b.start || a.end - b.end);
    const pairs: CrossPair[] = [];
    const involved = new Map<string, Order>();
    for (let i = 0; i < orders.length; i++) {
      for (let t = i + 1; t < orders.length; t++) {
        if (crosses(orders[i], orders[t])) {
          const pair: CrossPair = { code, a: orders[i], b: orders[t] };
          pairs.push(pair);
          allPairs.push(pair);
          involved.set(`${orders[i].start}|${orders[i].end}|${orders[i].type}`, orders[i]);
          involved.set(`${orders[t].start}|${orders[t].end}|${orders[t].type}`, orders[t]);
        }
      }
    }
    if (pairs.length) {
      const distinct = Array.from(involved.values()).sort(
        (a, b) => a.start - b.start || a.end - b.end,
      );
      groups.push({ code, pairs, orders: distinct });
    }
  }

  onProgress("Nəticə hazırlanır...", 85);

  wb.Sheets["Result"] = buildResultSheet(groups);
  if (!wb.SheetNames.includes("Result")) wb.SheetNames.push("Result");

  onProgress("Tamamlandı", 100);

  return {
    workbook: wb,
    sheetName,
    groups,
    pairs: allPairs,
    totalRecords,
    employeesWithCross: groups.length,
    durationMs: performance.now() - t0,
  };
}

// Result sheet: one row per unique Employee Badge.
//   A = Emp Badge
//   B = Kəsişən əmrlər və tarixləri (one order per line, in a single cell)
//   C = Status (aligned, one status per line)
function buildResultSheet(groups: CrossGroup[]): XLSX.WorkSheet {
  const aoa: Row[] = [["Emp Badge", "Kəsişən əmrlər və tarixləri", "Status"]];
  for (const g of groups) {
    const dateLines = g.orders
      .map((o) => `${formatSerial(o.start)} – ${formatSerial(o.end)}`)
      .join("\n");
    const statusLines = g.orders.map((o) => o.type || "—").join("\n");
    aoa.push([g.code, dateLines, statusLines]);
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 16 }, { wch: 34 }, { wch: 30 }];
  // Enable wrap so the multi-line cells display as separate lines in Excel.
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1");
  for (let r = 1; r <= range.e.r; r++) {
    for (const c of [1, 2]) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell) cell.s = { alignment: { wrapText: true, vertical: "top" } };
    }
  }
  return ws;
}

export function exportCrossCheck(
  result: CrossResult,
  filename: string,
  override?: CrossGroup[],
) {
  const groups = override ?? result.groups;
  result.workbook.Sheets["Result"] = buildResultSheet(groups);
  if (!result.workbook.SheetNames.includes("Result")) result.workbook.SheetNames.push("Result");
  XLSX.writeFile(result.workbook, filename, { compression: true });
}
