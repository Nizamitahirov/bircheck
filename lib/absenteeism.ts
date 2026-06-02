import * as XLSX from "xlsx";

export type ProgressFn = (step: string, pct: number) => void;
export type Cell = string | number | boolean | Date | null | undefined;
type Row = Cell[];

export interface NeticeRow {
  code: string;
  name: string;
  total: number;
}

export interface MuqayiseRow {
  code: string;
  name: string;
  dateSerial: number;
  type: number;
  weight: number;
}

export interface AbsenteeismResult {
  workbook: XLSX.WorkBook;
  netice: NeticeRow[];
  muqayise: MuqayiseRow[];
  periodStart: number;
  periodEnd: number;
  totalEmployees: number;
  durationMs: number;
}

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;

function excelSerialFromDate(d: Date): number {
  const utc = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  return Math.round((utc - EXCEL_EPOCH_MS) / MS_PER_DAY);
}

function dateFromSerial(s: number): Date {
  return new Date(EXCEL_EPOCH_MS + Math.round(s) * MS_PER_DAY);
}

function todaySerial(): number {
  return excelSerialFromDate(new Date());
}

function asSerial(v: Cell): number | null {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return Math.round(v);
  if (v instanceof Date) return excelSerialFromDate(v);
  const s = String(v).trim();
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : excelSerialFromDate(d);
}

function toKey(v: Cell): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return String(Math.trunc(v));
  return String(v).trim();
}

function toNumber(v: Cell): number {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const n = Number(String(v).replace(",", "."));
  return Number.isFinite(n) ? n : 0;
}

function sheetToRows(ws: XLSX.WorkSheet): Row[] {
  return XLSX.utils.sheet_to_json<Row>(ws, {
    header: 1,
    raw: true,
    defval: null,
    blankrows: false,
  });
}

export function formatExcelDate(v: Cell): string {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number") {
    if (v > 30000 && v < 80000) {
      const d = dateFromSerial(v);
      const dd = String(d.getUTCDate()).padStart(2, "0");
      const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
      return `${dd}.${mm}.${d.getUTCFullYear()}`;
    }
    return String(v);
  }
  if (v instanceof Date) {
    const dd = String(v.getUTCDate()).padStart(2, "0");
    const mm = String(v.getUTCMonth() + 1).padStart(2, "0");
    return `${dd}.${mm}.${v.getUTCFullYear()}`;
  }
  return String(v);
}

interface DataRow {
  code: string;
  name: string; // best-effort B column
  type: number;
  start: number;
  end: number;
  raw: Row;
}

export async function processAbsenteeism(
  file: File,
  onProgress: ProgressFn = () => {},
): Promise<AbsenteeismResult> {
  const t0 = performance.now();

  onProgress("Excel faylı oxunur...", 5);
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: false });

  if (!wb.SheetNames.includes("Data")) {
    throw new Error('Faylda "Data" sheet-i tapılmadı.');
  }

  const dataRowsRaw = sheetToRows(wb.Sheets["Data"]);
  if (dataRowsRaw.length < 2) {
    throw new Error('"Data" sheet boşdur — heç bir məlumat yoxdur.');
  }
  const header = dataRowsRaw[0];

  const today = todaySerial();
  const periodStart = today - 63;

  onProgress("Tarix aralığı dəqiqləşdirilir...", 15);

  // Step 1: clip / drop rows based on the [today-63, today] window.
  const adjustedRaw: Row[] = [header];
  const items: DataRow[] = [];
  for (let i = 1; i < dataRowsRaw.length; i++) {
    const src = dataRowsRaw[i];
    if (!src) continue;
    const code = toKey(src[0]);
    if (!code) continue;

    let start = asSerial(src[4]);
    let end = asSerial(src[5]);
    if (start === null || end === null) continue;
    if (start > end) [start, end] = [end, start];

    // VBA conditions, simplified to interval clip vs drop:
    if (end < periodStart) continue; // entirely before window
    if (start > today) continue; // entirely after window
    if (start < periodStart) start = periodStart;
    if (end > today) end = today;

    const adjustedRow = [...src];
    adjustedRow[4] = start;
    adjustedRow[5] = end;
    adjustedRaw.push(adjustedRow);

    items.push({
      code,
      name: src[1] === null || src[1] === undefined ? "" : String(src[1]).trim(),
      type: toNumber(src[2]),
      start,
      end,
      raw: adjustedRow,
    });
  }

  // Sort items by code asc, start asc (matches VBA initial sort).
  items.sort((a, b) => (a.code < b.code ? -1 : a.code > b.code ? 1 : a.start - b.start));

  onProgress("Personal kodları toplanır...", 30);

  // Step 2: unique personal codes for Netice.
  const namesByCode = new Map<string, string>();
  for (const it of items) {
    if (it.name && !namesByCode.get(it.code)) namesByCode.set(it.code, it.name);
  }
  const codes = Array.from(new Set(items.map((i) => i.code))).sort();

  onProgress("Tarixlər genişləndirilir...", 50);

  // Step 3: expand each interval into individual day rows.
  type Day = { code: string; name: string; date: number; type: number; D: number };
  let days: Day[] = [];
  for (const it of items) {
    for (let p = 0; p <= it.end - it.start; p++) {
      days.push({
        code: it.code,
        name: it.name,
        date: it.end - p,
        type: it.type,
        D: 0,
      });
    }
  }

  // Sort: A asc, B asc, C desc (so dedup keeps the highest C per (A, B))
  days.sort(
    (a, b) =>
      (a.code < b.code ? -1 : a.code > b.code ? 1 : 0) ||
      a.date - b.date ||
      b.type - a.type,
  );

  // Drop type 42 and 29 first, then dedup on (code, date) keeping first.
  days = days.filter((d) => d.type !== 42 && d.type !== 29);
  const seen = new Set<string>();
  days = days.filter((d) => {
    const k = `${d.code}|${d.date}`;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });

  onProgress("Müddətlər hesablanır...", 70);

  // Step 4: compute D weight per row.
  for (let j = 0; j < days.length; j++) {
    if (j === 0 || days[j].code !== days[j - 1].code) {
      days[j].D = 1;
      continue;
    }
    const c = days[j].type;
    const gap = days[j].date - days[j - 1].date;
    if (c === 30 || c === 32) {
      days[j].D = gap > 3 ? 4 : -1;
    } else {
      days[j].D = gap;
    }
  }

  // Step 5: chain filter — when D > 3 we reached a "fresh chain", so set the
  // anchor's D to 1 and drop every earlier consecutive same-person row.
  let l = days.length - 1;
  while (l >= 1) {
    if (days[l].D > 3) {
      const code = days[l].code;
      let k = l - 1;
      while (k >= 0 && days[k].code === code) {
        days[l].D = 1;
        days.splice(k, 1);
        l--;
        k--;
      }
    }
    l--;
  }

  onProgress("Yekun hesablanır...", 88);

  // Step 6: aggregate D into per-employee totals.
  const totals = new Map<string, number>();
  for (const d of days) {
    totals.set(d.code, (totals.get(d.code) ?? 0) + d.D);
  }

  const netice: NeticeRow[] = codes.map((code) => ({
    code,
    name: namesByCode.get(code) ?? "",
    total: totals.get(code) ?? 0,
  }));

  const muqayise: MuqayiseRow[] = days.map((d) => ({
    code: d.code,
    name: d.name,
    dateSerial: d.date,
    type: d.type,
    weight: d.D,
  }));

  onProgress("Excel fayl yaradılır...", 96);

  // Replace Data with the adjusted version; build Netice/Muqayise sheets.
  wb.Sheets["Data"] = XLSX.utils.aoa_to_sheet(adjustedRaw);

  const neticeAoa: Row[] = [["Personal kod", "Müddət"], ...netice.map((n) => [n.code, n.total])];
  const neticeWs = XLSX.utils.aoa_to_sheet(neticeAoa);
  neticeWs["!cols"] = [{ wch: 16 }, { wch: 12 }];
  wb.Sheets["Netice"] = neticeWs;

  const muqayiseAoa: Row[] = [
    ["Personal kod", "Tarix", "Növ", "Müddət"],
    ...muqayise.map((m) => [m.code, m.dateSerial, m.type, m.weight]),
  ];
  const muqayiseWs = XLSX.utils.aoa_to_sheet(muqayiseAoa);
  muqayiseWs["!cols"] = [{ wch: 16 }, { wch: 14 }, { wch: 10 }, { wch: 10 }];
  // Format Tarix column as date
  for (let r = 1; r <= muqayise.length; r++) {
    const addr = XLSX.utils.encode_cell({ r, c: 1 });
    const cell = muqayiseWs[addr];
    if (cell && typeof cell.v === "number") {
      cell.t = "n";
      cell.z = "dd.mm.yyyy";
    }
  }
  wb.Sheets["Muqayise"] = muqayiseWs;

  for (const name of ["Netice", "Muqayise"]) {
    if (!wb.SheetNames.includes(name)) wb.SheetNames.push(name);
  }

  onProgress("Tamamlandı", 100);

  return {
    workbook: wb,
    netice,
    muqayise,
    periodStart,
    periodEnd: today,
    totalEmployees: codes.length,
    durationMs: performance.now() - t0,
  };
}

export function exportAbsenteeism(
  result: AbsenteeismResult,
  filename: string,
  override?: { netice: NeticeRow[]; muqayise: MuqayiseRow[] },
) {
  const netice = override?.netice ?? result.netice;
  const muqayise = override?.muqayise ?? result.muqayise;
  const wb = result.workbook;

  const neticeAoa: Row[] = [["Personal kod", "Müddət"], ...netice.map((n) => [n.code, n.total])];
  const neticeWs = XLSX.utils.aoa_to_sheet(neticeAoa);
  neticeWs["!cols"] = [{ wch: 16 }, { wch: 12 }];
  wb.Sheets["Netice"] = neticeWs;

  const muqayiseAoa: Row[] = [
    ["Personal kod", "Tarix", "Növ", "Müddət"],
    ...muqayise.map((m) => [m.code, m.dateSerial, m.type, m.weight]),
  ];
  const muqayiseWs = XLSX.utils.aoa_to_sheet(muqayiseAoa);
  muqayiseWs["!cols"] = [{ wch: 16 }, { wch: 14 }, { wch: 10 }, { wch: 10 }];
  for (let r = 1; r <= muqayise.length; r++) {
    const addr = XLSX.utils.encode_cell({ r, c: 1 });
    const cell = muqayiseWs[addr];
    if (cell && typeof cell.v === "number") {
      cell.t = "n";
      cell.z = "dd.mm.yyyy";
    }
  }
  wb.Sheets["Muqayise"] = muqayiseWs;
  for (const name of ["Netice", "Muqayise"]) {
    if (!wb.SheetNames.includes(name)) wb.SheetNames.push(name);
  }

  XLSX.writeFile(wb, filename, { compression: true });
}
