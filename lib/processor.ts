import * as XLSX from "xlsx";

export type ProgressFn = (step: string, pct: number) => void;

export type Cell = string | number | boolean | Date | null | undefined;
export type Row = Cell[];

export interface ProblemsRow {
  id: string;
  cells: Cell[];
}

export interface ProcessResult {
  workbook: XLSX.WorkBook;
  header: Cell[];
  rows: ProblemsRow[];
  totalEmployees: number;
  diffCount: number;
  durationMs: number;
}

// Tanınan məzuniyyət növləri. Hesablamada `From_HRB_Otpusk` H sütununun
// cəmi alınır — növdən asılı olmayaraq bütün bu yazılar avtomatik daxil olur.
export const VACATION_TYPES = [
  "Əmək məzuniyyəti",
  "Xəstəlik vərəqəsi",
  "HİK məzuniyyəti",
  "Öz hesabına (İşə gəlməmə)",
  "Ödənişsiz məzuniyyət",
  "Təhsil məzuniyyəti",
  "Kənarlaşma",
  "Uşağa qulluğa görə məzuniyyət (Q. Sosial məzuniyyət)",
  "Sertifikasiyaya görə məzuniyyət",
  "Məzuniyyətdən geri çağırılma",
  "Analıq məzuniyyəti (Hamiləlik və doğuşa görə məzuniyyət)",
  "Uşağa qulluğa görə məzuniyyətin dayandırılması",
  "Orta əmək haqqı saxlamaqla işə gəlməmə",
];

const SHEETS = {
  problems: "Problems",
  baza: "From_Excel_Baza",
  mhmd: "From_MHMD",
  otpusk: "From_HRB_Otpusk",
} as const;

const DEFAULT_HEADER: Cell[] = [
  "Personal kod",
  "Soyad",
  "Ad",
  "Ata adı",
  "Vəzifə",
  "Departament",
  "Filial",
  "İşə qəbul tarixi",
  "İş qrafiki",
  "MHMD iş günü",
  "İş günü Bizdə olan",
  "Fərq",
  "Fərqin səbəbi",
];

function sheetToRows(ws: XLSX.WorkSheet): Row[] {
  return XLSX.utils.sheet_to_json<Row>(ws, {
    header: 1,
    raw: true,
    defval: null,
    blankrows: false,
  });
}

function toNumber(v: Cell): number {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (v instanceof Date) return excelSerialFromDate(v);
  const n = Number(String(v).replace(",", "."));
  return Number.isFinite(n) ? n : 0;
}

function toKey(v: Cell): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return String(Math.trunc(v));
  return String(v).trim();
}

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;

function dateFromExcelSerial(serial: number): Date {
  return new Date(EXCEL_EPOCH_MS + Math.round(serial) * MS_PER_DAY);
}

function excelSerialFromDate(d: Date): number {
  const utc = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  return Math.round((utc - EXCEL_EPOCH_MS) / MS_PER_DAY);
}

function asExcelSerial(v: Cell): number | null {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return Math.round(v);
  if (v instanceof Date) return excelSerialFromDate(v);
  const s = String(v).trim();
  if (!s) return null;
  const d = new Date(s);
  if (!isNaN(d.getTime())) return excelSerialFromDate(d);
  return null;
}

// Excel NETWORKDAYS.INTL(start, end, weekend=11, holidays). weekend=11 → only Sunday.
function networkDaysIntl(
  startSerial: number,
  endSerial: number,
  holidays: Set<number>,
): number {
  if (endSerial < startSerial) return -networkDaysIntl(endSerial, startSerial, holidays);
  let count = 0;
  for (let s = startSerial; s <= endSerial; s++) {
    const d = new Date(EXCEL_EPOCH_MS + s * MS_PER_DAY);
    if (d.getUTCDay() === 0) continue;
    if (holidays.has(s)) continue;
    count++;
  }
  return count;
}

function getCell(rows: Row[], r: number, c: number): Cell {
  return rows[r]?.[c] ?? null;
}

function findLastRow(rows: Row[], col: number): number {
  for (let i = rows.length - 1; i >= 0; i--) {
    const v = rows[i]?.[col];
    if (v !== null && v !== undefined && v !== "") return i;
  }
  return -1;
}

export function formatExcelDate(v: Cell): string {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number") {
    if (v > 30000 && v < 80000) {
      const d = dateFromExcelSerial(v);
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

export async function processWorkbook(
  file: File,
  onProgress: ProgressFn = () => {},
): Promise<ProcessResult> {
  const t0 = performance.now();

  onProgress("Excel faylı oxunur...", 5);
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: false });

  const missing = (Object.values(SHEETS) as string[]).filter(
    (n) => !wb.SheetNames.includes(n),
  );
  if (missing.length) {
    throw new Error(
      `Faylda bu sheet(lər) yoxdur: ${missing.join(", ")}. Tələb olunan sheet adları: ${Object.values(
        SHEETS,
      ).join(", ")}.`,
    );
  }

  const bazaRows = sheetToRows(wb.Sheets[SHEETS.baza]);
  const mhmdRowsRaw = sheetToRows(wb.Sheets[SHEETS.mhmd]);
  const otpuskRows = sheetToRows(wb.Sheets[SHEETS.otpusk]);

  const periodStart = asExcelSerial(getCell(otpuskRows, 1, 9)); // J2
  const periodEnd = asExcelSerial(getCell(otpuskRows, 2, 9)); // J3
  const totalWorkdays = toNumber(getCell(otpuskRows, 1, 10)); // K2

  if (periodStart === null || periodEnd === null) {
    throw new Error(
      "From_HRB_Otpusk sheet-də J2 (başlama tarixi) və J3 (bitmə tarixi) doldurulmalıdır.",
    );
  }
  if (!totalWorkdays) {
    throw new Error("From_HRB_Otpusk sheet-də K2 (ümumi iş günü) doldurulmalıdır.");
  }

  const holidaysDefault = new Set<number>();
  const holidays6Day = new Set<number>();
  for (let r = 1; r <= 28; r++) {
    const vI = asExcelSerial(getCell(otpuskRows, r, 8)); // I column
    if (vI !== null) holidaysDefault.add(vI);
    const vL = asExcelSerial(getCell(otpuskRows, r, 11)); // L column — "Qeyri-iş günləri cari ay - 6 günlük"
    if (vL !== null) holidays6Day.add(vL);
  }

  onProgress("MHMD məlumatları təmizlənir...", 15);

  // Drop MHMD rows outside the period
  const mhmdRows: Row[] = [];
  if (mhmdRowsRaw.length > 0) mhmdRows.push(mhmdRowsRaw[0]);
  for (let i = 1; i < mhmdRowsRaw.length; i++) {
    const row = mhmdRowsRaw[i];
    if (!row || row[0] === null || row[0] === undefined || row[0] === "") continue;
    const d = asExcelSerial(row[3]);
    const e = asExcelSerial(row[4]);
    if (d !== null && d > periodEnd) continue;
    if (e !== null && e < periodStart) continue;
    mhmdRows.push(row);
  }

  onProgress("Baza Problems-ə köçürülür...", 25);

  const existingHeader = sheetToRows(wb.Sheets[SHEETS.problems])[0];
  const problemsHeader: Cell[] = existingHeader ? [...existingHeader] : [...DEFAULT_HEADER];
  while (problemsHeader.length < 13) problemsHeader.push("");
  // Ensure our key columns have correct names if they were blank
  if (!problemsHeader[9]) problemsHeader[9] = DEFAULT_HEADER[9];
  if (!problemsHeader[10]) problemsHeader[10] = DEFAULT_HEADER[10];
  if (!problemsHeader[11]) problemsHeader[11] = DEFAULT_HEADER[11];
  if (!problemsHeader[12]) problemsHeader[12] = DEFAULT_HEADER[12];

  const allRows: Row[] = [];
  const bazaLastRow = findLastRow(bazaRows, 0);
  for (let i = 1; i <= bazaLastRow; i++) {
    const src = bazaRows[i];
    if (!src) continue;
    const row: Row = new Array(13).fill(null);
    for (let c = 0; c < 9; c++) row[c] = src[c] ?? null;
    allRows.push(row);
  }

  onProgress("MHMD iş günləri toplanılır...", 45);

  // From_MHMD: sum F where A=key AND C="001"
  const mhmdSum = new Map<string, number>();
  const mhmdLast = findLastRow(mhmdRows, 0);
  for (let i = 1; i <= mhmdLast; i++) {
    const row = mhmdRows[i];
    if (!row) continue;
    const key = toKey(row[0]);
    if (!key) continue;
    if (toKey(row[2]) !== "001") continue;
    mhmdSum.set(key, (mhmdSum.get(key) ?? 0) + toNumber(row[5]));
  }

  onProgress("Məzuniyyət məlumatları işlənir...", 60);

  // From_HRB_Otpusk: sum H (all vacation types) and collect details for H<>0
  const otpuskSum = new Map<string, number>();
  const otpuskDetails = new Map<string, string[]>();
  const otpuskLast = findLastRow(otpuskRows, 0);
  for (let i = 1; i <= otpuskLast; i++) {
    const row = otpuskRows[i];
    if (!row) continue;
    const key = toKey(row[0]);
    if (!key) continue;
    const days = toNumber(row[7]);
    otpuskSum.set(key, (otpuskSum.get(key) ?? 0) + days);
    if (days !== 0) {
      const type = row[4] === null || row[4] === undefined ? "" : String(row[4]).trim();
      const piece = `(${formatExcelDate(row[5])}-${formatExcelDate(row[6])}-${type});`;
      const arr = otpuskDetails.get(key) ?? [];
      arr.push(piece);
      otpuskDetails.set(key, arr);
    }
  }

  onProgress("Fərqlər hesablanır...", 75);

  for (const row of allRows) {
    const key = toKey(row[0]);
    if (!key) continue;
    const mhmdDays = mhmdSum.get(key) ?? 0;
    const otpuskTotal = otpuskSum.get(key) ?? 0;
    const hireDate = asExcelSerial(row[7]);

    // Column I in Baza = İş qrafiki. Employees whose schedule starts with
    // "6 günlük iş" use the L column of otpusk for non-work days; everyone
    // else keeps using the I column.
    const schedule = String(row[8] ?? "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();
    const holidays = schedule.startsWith("6 günlük iş") ? holidays6Day : holidaysDefault;

    let computed: number;
    if (hireDate !== null && hireDate < periodStart) {
      computed = totalWorkdays - otpuskTotal;
    } else if (hireDate !== null) {
      computed = networkDaysIntl(hireDate, periodEnd, holidays) - otpuskTotal;
    } else {
      computed = totalWorkdays - otpuskTotal;
    }

    row[9] = mhmdDays;
    row[10] = computed;
    row[11] = computed - mhmdDays;
  }

  onProgress("Fərqin səbəbi əlavə olunur...", 92);

  // Populate M for every row whose computed days are below the full month
  // (regardless of whether the diff is zero — the file keeps all employees).
  let diffCount = 0;
  for (const row of allRows) {
    const key = toKey(row[0]);
    const computed = toNumber(row[10]);
    if (computed < totalWorkdays) {
      const details = otpuskDetails.get(key);
      row[12] = details && details.length ? details.join("") : "";
    } else {
      row[12] = "";
    }
    if (toNumber(row[11]) !== 0) diffCount++;
  }

  // Replace MHMD sheet with cleaned version; Problems will be rebuilt on export
  wb.Sheets[SHEETS.mhmd] = XLSX.utils.aoa_to_sheet(mhmdRows);

  const rows: ProblemsRow[] = allRows.map((cells, idx) => ({
    id: `${toKey(cells[0]) || "row"}-${idx}`,
    cells,
  }));

  onProgress("Tamamlandı", 100);

  return {
    workbook: wb,
    header: problemsHeader,
    rows,
    totalEmployees: allRows.length,
    diffCount,
    durationMs: performance.now() - t0,
  };
}

export function exportResult(
  result: ProcessResult,
  rows: ProblemsRow[],
  filename: string,
) {
  const aoa: Row[] = [result.header, ...rows.map((r) => r.cells)];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [
    { wch: 14 },
    { wch: 20 },
    { wch: 16 },
    { wch: 16 },
    { wch: 28 },
    { wch: 24 },
    { wch: 20 },
    { wch: 14 },
    { wch: 12 },
    { wch: 12 },
    { wch: 16 },
    { wch: 10 },
    { wch: 60 },
  ];

  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1");
  for (let r = 1; r <= range.e.r; r++) {
    const addr = XLSX.utils.encode_cell({ r, c: 7 });
    const cell = ws[addr];
    if (cell && typeof cell.v === "number" && cell.v > 30000 && cell.v < 80000) {
      cell.t = "n";
      cell.z = "dd.mm.yyyy";
    }
  }

  result.workbook.Sheets[SHEETS.problems] = ws;
  XLSX.writeFile(result.workbook, filename, { compression: true });
}
