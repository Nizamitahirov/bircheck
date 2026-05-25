import * as XLSX from "xlsx";

export type ProgressFn = (step: string, pct: number) => void;

export interface ProcessResult {
  workbook: XLSX.WorkBook;
  rowsKept: number;
  rowsRemoved: number;
  totalEmployees: number;
  durationMs: number;
}

const SHEETS = {
  problems: "Problems",
  baza: "From_Excel_Baza",
  mhmd: "From_MHMD",
  otpusk: "From_HRB_Otpusk",
} as const;

type Cell = string | number | boolean | Date | null | undefined;
type Row = Cell[];

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

// Excel serial date helpers.
// Excel epoch is 1899-12-30 (because of the 1900 leap-year bug compensation).
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
  // Try parse string as date
  const s = String(v).trim();
  if (!s) return null;
  const d = new Date(s);
  if (!isNaN(d.getTime())) return excelSerialFromDate(d);
  return null;
}

/**
 * Excel NETWORKDAYS.INTL(start, end, weekend=11, holidays)
 * weekend = 11 means "Sunday only" is weekend (Mon-Sat are workdays).
 */
function networkDaysIntl(
  startSerial: number,
  endSerial: number,
  holidays: Set<number>,
  weekendType: 11 = 11,
): number {
  if (endSerial < startSerial) {
    return -networkDaysIntl(endSerial, startSerial, holidays, weekendType);
  }
  // weekend = 11 → Sunday only
  // Excel weekday for serial: 0 = Saturday (since 1900-01-00 was Saturday in Excel),
  // We'll just use the JS getUTCDay where 0 = Sunday, 6 = Saturday.
  let count = 0;
  for (let s = startSerial; s <= endSerial; s++) {
    const d = new Date(EXCEL_EPOCH_MS + s * MS_PER_DAY);
    const dow = d.getUTCDay(); // 0 = Sunday
    if (dow === 0) continue; // weekend type 11 = Sunday only
    if (holidays.has(s)) continue;
    count++;
  }
  return count;
}

function getCell(rows: Row[], r: number, c: number): Cell {
  const row = rows[r];
  if (!row) return null;
  return row[c] ?? null;
}

function findLastRow(rows: Row[], col: number): number {
  for (let i = rows.length - 1; i >= 0; i--) {
    const v = rows[i]?.[col];
    if (v !== null && v !== undefined && v !== "") return i;
  }
  return -1;
}

export async function processWorkbook(
  file: File,
  onProgress: ProgressFn = () => {},
): Promise<ProcessResult> {
  const t0 = performance.now();

  onProgress("Excel faylı oxunur...", 5);
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: false });

  // Validate required sheets
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

  // ---- Read control values from From_HRB_Otpusk ----
  // J2 = start of period, J3 = end of period (excel serials)
  // K2 = total workdays in the month
  // I2:I29 = holiday list
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

  const holidays = new Set<number>();
  for (let r = 1; r <= 28; r++) {
    const v = asExcelSerial(getCell(otpuskRows, r, 8)); // I2:I29
    if (v !== null) holidays.add(v);
  }

  onProgress("MHMD məlumatları təmizlənir...", 15);

  // ---- Step 0: Clean From_MHMD ----
  // Drop rows where D > J3 OR E < J2  (D=col 3, E=col 4 in 0-index)
  const mhmdRows: Row[] = [];
  // keep header
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

  // ---- Step 1: Copy A:I from From_Excel_Baza to Problems ----
  // Header row stays; we copy data rows.
  const problemsHeader: Row =
    sheetToRows(wb.Sheets[SHEETS.problems])[0] ??
    [
      "Personal kod",
      "Soyad",
      "Ad",
      "Ata adı",
      "Vəzifə",
      "Departament",
      "Filial",
      "İşə qəbul tarixi",
      "Status",
      "MHMD iş günü",
      "Hesablanmış iş günü",
      "Fərq",
      "Mezuniyyət detalları",
    ];

  // Ensure header has at least 13 columns
  while (problemsHeader.length < 13) problemsHeader.push("");

  const problemsRows: Row[] = [problemsHeader];

  const bazaLastRow = findLastRow(bazaRows, 0);
  for (let i = 1; i <= bazaLastRow; i++) {
    const src = bazaRows[i];
    if (!src) continue;
    const row: Row = new Array(13).fill(null);
    for (let c = 0; c < 9; c++) row[c] = src[c] ?? null;
    problemsRows.push(row);
  }

  onProgress("MHMD iş günləri toplanılır...", 45);

  // ---- Step 2: Build indices for O(N) joins ----
  // From_MHMD: sum F where A=key AND C="001"
  const mhmdSum = new Map<string, number>();
  const mhmdLast = findLastRow(mhmdRows, 0);
  for (let i = 1; i <= mhmdLast; i++) {
    const row = mhmdRows[i];
    if (!row) continue;
    const key = toKey(row[0]);
    if (!key) continue;
    const type = toKey(row[2]);
    if (type !== "001") continue;
    const val = toNumber(row[5]);
    mhmdSum.set(key, (mhmdSum.get(key) ?? 0) + val);
  }

  onProgress("Mezuniyyət məlumatları işlənir...", 60);

  // From_HRB_Otpusk: sum H where A=key; also collect entries with H<>0 for details
  const otpuskSum = new Map<string, number>();
  const otpuskDetails = new Map<string, string[]>();
  const otpuskLast = findLastRow(otpuskRows, 0);
  for (let i = 1; i <= otpuskLast; i++) {
    const row = otpuskRows[i];
    if (!row) continue;
    const key = toKey(row[0]);
    if (!key) continue;
    const days = toNumber(row[7]); // H
    otpuskSum.set(key, (otpuskSum.get(key) ?? 0) + days);
    if (days !== 0) {
      const e = row[4];
      const f = row[5];
      const g = row[6];
      const fStr = formatMaybeDate(f);
      const gStr = formatMaybeDate(g);
      const eStr = e === null || e === undefined ? "" : String(e);
      const piece = `(${fStr}-${gStr}-${eStr});`;
      const arr = otpuskDetails.get(key) ?? [];
      arr.push(piece);
      otpuskDetails.set(key, arr);
    }
  }

  onProgress("Fərqlər hesablanır...", 75);

  // ---- Step 3: For each Problems row compute J, K, L ----
  for (let i = 1; i < problemsRows.length; i++) {
    const row = problemsRows[i];
    const key = toKey(row[0]);
    if (!key) continue;

    const mhmdDays = mhmdSum.get(key) ?? 0;
    const otpuskTotal = otpuskSum.get(key) ?? 0;

    const hireDate = asExcelSerial(row[7]); // column H = index 7
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

  onProgress("Sıfır fərqli sətirlər silinir...", 85);

  // ---- Step 4: Remove rows where L == 0 ----
  const beforeCount = problemsRows.length - 1;
  const filtered: Row[] = [problemsRows[0]];
  for (let i = 1; i < problemsRows.length; i++) {
    const row = problemsRows[i];
    const diff = toNumber(row[11]);
    if (diff === 0) continue;
    filtered.push(row);
  }
  const afterCount = filtered.length - 1;

  onProgress("Mezuniyyət detalları əlavə olunur...", 92);

  // ---- Step 5: For rows with K < K2, append vacation details to M ----
  for (let i = 1; i < filtered.length; i++) {
    const row = filtered[i];
    const key = toKey(row[0]);
    const computed = toNumber(row[10]);
    if (computed < totalWorkdays) {
      const details = otpuskDetails.get(key);
      if (details && details.length) row[12] = details.join("");
      else row[12] = "";
    } else {
      row[12] = "";
    }
  }

  onProgress("Excel fayl yaradılır...", 96);

  // ---- Build output workbook ----
  const newProblemsWs = XLSX.utils.aoa_to_sheet(filtered);
  // Preserve column widths roughly
  newProblemsWs["!cols"] = [
    { wch: 14 }, // A
    { wch: 20 }, // B
    { wch: 16 }, // C
    { wch: 16 }, // D
    { wch: 28 }, // E
    { wch: 24 }, // F
    { wch: 20 }, // G
    { wch: 14 }, // H
    { wch: 12 }, // I
    { wch: 12 }, // J
    { wch: 14 }, // K
    { wch: 10 }, // L
    { wch: 60 }, // M
  ];

  // Format date column H if numeric
  const range = XLSX.utils.decode_range(newProblemsWs["!ref"] ?? "A1");
  for (let r = 1; r <= range.e.r; r++) {
    const addr = XLSX.utils.encode_cell({ r, c: 7 });
    const cell = newProblemsWs[addr];
    if (cell && typeof cell.v === "number" && cell.v > 30000 && cell.v < 80000) {
      cell.t = "n";
      cell.z = "dd.mm.yyyy";
    }
  }

  // Replace Problems sheet in workbook
  wb.Sheets[SHEETS.problems] = newProblemsWs;
  // Also replace MHMD with cleaned version
  wb.Sheets[SHEETS.mhmd] = XLSX.utils.aoa_to_sheet(mhmdRows);

  onProgress("Tamamlandı", 100);

  return {
    workbook: wb,
    rowsKept: afterCount,
    rowsRemoved: beforeCount - afterCount,
    totalEmployees: beforeCount,
    durationMs: performance.now() - t0,
  };
}

function formatMaybeDate(v: Cell): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") {
    // Heuristic: looks like an Excel date serial
    if (v > 30000 && v < 80000) {
      const d = dateFromExcelSerial(v);
      const dd = String(d.getUTCDate()).padStart(2, "0");
      const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
      const yy = d.getUTCFullYear();
      return `${dd}.${mm}.${yy}`;
    }
    return String(v);
  }
  if (v instanceof Date) {
    const dd = String(v.getUTCDate()).padStart(2, "0");
    const mm = String(v.getUTCMonth() + 1).padStart(2, "0");
    const yy = v.getUTCFullYear();
    return `${dd}.${mm}.${yy}`;
  }
  return String(v);
}

export function downloadWorkbook(wb: XLSX.WorkBook, filename: string) {
  XLSX.writeFile(wb, filename, { compression: true });
}
