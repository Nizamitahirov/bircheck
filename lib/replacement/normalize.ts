import * as XLSX from "xlsx";
import {
  BO_KEYWORDS,
  INZIBATCI_COST_NOV,
  MUVEQQETI_COST_TEXTS,
  NO_COST_POSITIONS,
  type ReplacementParams,
  type SourceName,
} from "./config";
import type { BackupSpec, BOF, CancelInfo, Cell, ErrorLogEntry, NormRow, Row } from "./types";
import {
  azLower,
  cellStr,
  fmtDate,
  findSheet,
  getCol,
  getColStr,
  headerMap,
  normCode,
  parseDateValue,
  parseYYYYMMDD,
  rowDump,
  serialMonth,
  serialYear,
  sheetToRows,
} from "./xl";

export interface SourceWorkbooks {
  inzibatci: XLSX.WorkBook;
  vacation: XLSX.WorkBook;
  filial: XLSX.WorkBook;
  resmi: XLSX.WorkBook;
  vezifeMaas: XLSX.WorkBook;
}

export interface NormalizeResult {
  rows: NormRow[];
  errors: ErrorLogEntry[];
  cancels: CancelInfo[];
  redRows: number[];
  backups: BackupSpec[];
  /** normallaşdırılan sətir sayı mənbə üzrə */
  sourceCounts: Record<SourceName, number>;
  /** SPEC 12.7 — Müvəqqəti keçiddə xərc qaydasına uyğun normallaşdırılan sətir sayı */
  muveqqetiCostCount: number;
  muveqqetiCostUnique: number;
}

interface Ctx {
  params: ReplacementParams;
  rows: NormRow[];
  errors: ErrorLogEntry[];
}

function pushErr(ctx: Ctx, menbe: string, rowNum: number | string, novu: string, teferruat: string, header: Row, row: Row) {
  ctx.errors.push({
    menbe,
    setr: String(rowNum),
    novu,
    teferruat,
    setirMelumati: rowDump(header, row),
  });
}

function classifyBO(text: string): BOF {
  const t = azLower(text);
  return BO_KEYWORDS.some((w) => t.includes(w)) ? "Baş ofis" : "Filial";
}

function boFromColumn(v: string, fallbackText: string): BOF {
  const t = azLower(v);
  if (!t) return classifyBO(fallbackText);
  if (t === "bo" || t.includes("baş ofis") || t.includes("bas ofis")) return "Baş ofis";
  if (t.includes("filial")) return "Filial";
  return classifyBO(fallbackText);
}

function isRowEmpty(row: Row): boolean {
  return row.every((c) => c === null || c === undefined || String(c).trim() === "");
}

const CAP_NOTE_DATE = "30.06.2026"; // qeyd mətnlərində istifadə olunan sərhəd

function capText(params: ReplacementParams): string {
  return fmtDate(params.capDate) || CAP_NOTE_DATE;
}

// ────────────────────────────────────────────────────────────────────────────
// Vəzifə-Maaş lookup (SPEC 3.6, 5.3)
// ────────────────────────────────────────────────────────────────────────────

interface VMEntry {
  vezife: string;
  departament: string;
}

export interface VMLookup {
  sheets: { name: string; map: Map<string, VMEntry> }[]; // index 0..5 = Yanvar..İyun
  find(kod: string, startSerial: number): { vezife: string; departament: string } | null;
  backupAoa: Row[];
  backupHeader: Row;
}

const MONTH_NAMES = ["yanvar", "fevral", "mart", "aprel", "may", "iyun"];

export function buildVMLookup(wb: XLSX.WorkBook): VMLookup {
  const sheets: { name: string; map: Map<string, VMEntry> }[] = new Array(6)
    .fill(null)
    .map(() => ({ name: "", map: new Map<string, VMEntry>() }));
  const backupAoa: Row[] = [];
  let backupHeader: Row = [];

  for (const sn of wb.SheetNames) {
    const low = azLower(sn);
    const mIdx = MONTH_NAMES.findIndex((m) => low.includes(m));
    if (mIdx < 0) continue;
    const rows = sheetToRows(wb.Sheets[sn]);
    if (rows.length < 2) continue;
    const hm = headerMap(rows[0]);
    if (!backupHeader.length) backupHeader = rows[0];
    const kodIdx = hm.get("Personal kod");
    const vezIdx = hm.get("Vəzifə");
    const depIdx = hm.get("Departament");
    const map = new Map<string, VMEntry>();
    const monthTag = `2026-${String(mIdx + 1).padStart(2, "0")}`;
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || isRowEmpty(r)) continue;
      const kod = normCode(kodIdx === undefined ? null : r[kodIdx]);
      if (!kod) continue;
      if (!map.has(kod)) {
        map.set(kod, {
          vezife: cellStr(vezIdx === undefined ? null : r[vezIdx]),
          departament: cellStr(depIdx === undefined ? null : r[depIdx]),
        });
      }
      backupAoa.push([sn, monthTag, ...r]);
    }
    sheets[mIdx] = { name: sn, map };
  }

  return {
    sheets,
    backupAoa,
    backupHeader,
    find(kod: string, startSerial: number) {
      if (!kod) return null;
      let mIdx = 5;
      const y = serialYear(startSerial);
      if (y < 2026) mIdx = 0;
      else if (y === 2026) mIdx = Math.min(Math.max(serialMonth(startSerial) - 1, 0), 5);
      const order = [mIdx, ...[5, 4, 3, 2, 1, 0].filter((i) => i !== mIdx)];
      for (const i of order) {
        const sh = this.sheets[i];
        if (!sh || !sh.name) continue;
        const hit = sh.map.get(kod);
        if (hit) {
          const suffix = ` (Vəzifə-Maaş, ${sh.name})`;
          return {
            vezife: hit.vezife ? hit.vezife + suffix : "",
            departament: hit.departament ? hit.departament + suffix : "",
          };
        }
      }
      return null;
    },
  };
}

// ────────────────────────────────────────────────────────────────────────────
// 3.1 İnzibatçı prosesi
// ────────────────────────────────────────────────────────────────────────────

function normInzibatci(ctx: Ctx, wb: XLSX.WorkBook, vm: VMLookup, backups: BackupSpec[]) {
  const MENBE: SourceName = "İnzibatçı prosesi";
  const sn = findSheet(wb, ["Sheet1"]) ?? wb.SheetNames[0];
  const rows = sheetToRows(wb.Sheets[sn]);
  if (rows.length < 1) throw new Error("İnzibatçılıq prosesi faylı boşdur.");
  const header = rows[0];
  const hm = headerMap(header);

  backups.push({
    sheetName: "BK_İnzibatçı prosesi",
    note: "Mənbə: İnzibatçılıq prosesi faylı — dəyişdirilməmiş surət.",
    aoa: rows,
  });

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const excelRow = i + 1; // massiv indeksi i (başlıq i=0) → Excel sətri i+1
    if (!r || isRowEmpty(r)) continue;

    const kod = normCode(getCol(hm, r, "Personal Kod"));
    const bas = parseDateValue(getCol(hm, r, "Baslama Tarixi"));
    if (bas === null) {
      pushErr(ctx, MENBE, excelRow, "Başlama tarixi boşdur", "Başlama tarixi yoxdur — sətir təmiz dataya salınmadı.", header, r);
      continue;
    }
    if (bas === "bad") {
      pushErr(ctx, MENBE, excelRow, "Başlama tarixi oxunmadı", `Başlama tarixi oxunmadı: "${getColStr(hm, r, "Baslama Tarixi")}"`, header, r);
      continue;
    }
    if (bas < ctx.params.filterStart) {
      pushErr(ctx, MENBE, excelRow, "Filtrdən kənar (2026-dan əvvəl)", `Başlama tarixi ${fmtDate(bas)} < ${fmtDate(ctx.params.filterStart)}.`, header, r);
      continue;
    }

    let bit = parseDateValue(getCol(hm, r, "Bitme Tarixi"));
    if (bit === "bad") bit = null; // oxunmayan bitmə → boş kimi
    const period = getColStr(hm, r, "Period");
    const isMuveqqeti = azLower(period).includes("müvəqqəti");
    const nov = getColStr(hm, r, "Növ");
    const teyinYeri = getColStr(hm, r, "Teyin Olundugu Yer");

    const qeydler: string[] = [];
    if (bit === null) {
      if (isMuveqqeti) {
        qeydler.push(`Bitmə tarixi boşdur (müvəqqəti) – gün sayı ${capText(ctx.params)}-a qədər hesablanıb`);
        pushErr(ctx, MENBE, excelRow, "Bitmə tarixi boşdur", "Müvəqqəti təyinatda bitmə tarixi boşdur — hesablama sərhədinə qədər hesablanır (sətir təmiz datada saxlanılıb).", header, r);
      } else {
        qeydler.push(`Daimi təyinat – bitmə tarixi yoxdur; gün sayı ${capText(ctx.params)}-a qədər hesablanıb`);
      }
    }

    const lk = vm.find(kod, bas as number);
    const novLow = azLower(nov);
    const xerc = INZIBATCI_COST_NOV.some((c) => novLow === azLower(c));

    ctx.rows.push({
      menbe: MENBE,
      menbeSetri: String(excelRow),
      rowNums: [excelRow],
      evezEdenKod: kod,
      evezEdenAd: getColStr(hm, r, "Evezeden Emekdas"),
      evezEdenStruktur: lk?.departament ?? "",
      evezEdenVezife: lk?.vezife ?? "",
      evezEdilenKod: "",
      evezEdilenAd: getColStr(hm, r, "Evezedilen Emekdas"),
      evezEdilenStruktur: "",
      evezEdilenVezife: "",
      teyinYeri,
      basOfis: classifyBO(teyinYeri || lk?.departament || ""),
      evezetmeNovu: `${nov} (${isMuveqqeti ? "Müvəqqəti" : "Daimi"})`,
      baslama: bas as number,
      bitme: bit,
      xerc,
      qeydler,
      dayBased: true,
    });
  }
}

// ────────────────────────────────────────────────────────────────────────────
// 3.2 İşburaxma hallarının əvəzedilməsi
// ────────────────────────────────────────────────────────────────────────────

function isburaxmaCost(basOfis: BOF, vez1: string, vez2: string): boolean {
  if (basOfis !== "Filial") return true;
  const v1 = azLower(vez1.trim());
  const v2 = azLower(vez2.trim());
  const noCost = (v: string) => NO_COST_POSITIONS.some((p) => v === azLower(p));
  return !(noCost(v1) && noCost(v2));
}

function normIsburaxma(ctx: Ctx, wb: XLSX.WorkBook, backups: BackupSpec[]) {
  const MENBE: SourceName = "İşburaxma hallarının əvəzedilməsi";
  const sn = findSheet(wb, ["Vacation replacement"]);
  if (!sn) throw new Error('İşburaxma faylında "Vacation replacement" sheet-i tapılmadı.');
  const rows = sheetToRows(wb.Sheets[sn]);
  const header = rows[0] ?? [];
  const hm = headerMap(header);

  backups.push({
    sheetName: "BK_İşburaxma əvəzedilməsi",
    note: 'Mənbə: "Vacation replacement" sheet-i — dəyişdirilməmiş surət.',
    aoa: rows,
  });

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const excelRow = i + 1;
    if (!r || isRowEmpty(r)) continue;

    const bas = parseYYYYMMDD(getCol(hm, r, "MEZUNIYYETIN_BASHLAMA_TARIXI"));
    if (bas === null) {
      pushErr(ctx, MENBE, excelRow, "Başlama tarixi boşdur", "Məzuniyyətin başlama tarixi yoxdur — sətir təmiz dataya salınmadı.", header, r);
      continue;
    }
    if (bas === "bad") {
      pushErr(ctx, MENBE, excelRow, "Başlama tarixi oxunmadı", `Başlama tarixi oxunmadı: "${getColStr(hm, r, "MEZUNIYYETIN_BASHLAMA_TARIXI")}"`, header, r);
      continue;
    }
    if (bas < ctx.params.filterStart) {
      pushErr(ctx, MENBE, excelRow, "Filtrdən kənar (2026-dan əvvəl)", `Başlama tarixi ${fmtDate(bas)} < ${fmtDate(ctx.params.filterStart)}.`, header, r);
      continue;
    }

    const qeydler: string[] = [];
    let bit = parseYYYYMMDD(getCol(hm, r, "MEZUNIYYETIN_BITME_TARIXI"));
    if (bit === "bad") bit = null;
    if (bit !== null && serialYear(bit) > 2027) {
      // SPEC 3.2 — şübhəli daxiletmə xətası
      qeydler.push(
        `Şübhəli bitmə tarixi (${fmtDate(bit)}) – bitmə boş kimi qəbul edilib, gün sayı ${capText(ctx.params)}-a qədər hesablanıb`,
      );
      pushErr(ctx, MENBE, excelRow, "Şübhəli bitmə tarixi", `Bitmə tarixi ${fmtDate(bit)} 2027-dən sonradır — daxiletmə xətası sayılıb, bitmə boş kimi qəbul edilib (sətir təmiz datada saxlanılıb).`, header, r);
      bit = null;
    } else if (bit === null) {
      qeydler.push(`Bitmə tarixi boşdur (müvəqqəti) – gün sayı ${capText(ctx.params)}-a qədər hesablanıb`);
      pushErr(ctx, MENBE, excelRow, "Bitmə tarixi boşdur", "Bitmə tarixi boşdur — hesablama sərhədinə qədər hesablanır (sətir təmiz datada saxlanılıb).", header, r);
    }

    const evezEdenVez = getColStr(hm, r, "VEZIFE");
    const evezEdilenVez = getColStr(hm, r, "VEZIFE.1");
    const teyinYeri = getColStr(hm, r, "STRUKTUR.1");
    const basOfis = boFromColumn(getColStr(hm, r, "BO/Filial"), teyinYeri);

    ctx.rows.push({
      menbe: MENBE,
      menbeSetri: String(excelRow),
      rowNums: [excelRow],
      evezEdenKod: normCode(getCol(hm, r, "EVEZ_EDEN_PERSONAL_KOD")),
      evezEdenAd: getColStr(hm, r, "EVEZ_EDEN_FULL_NAME"),
      evezEdenStruktur: getColStr(hm, r, "STRUKTUR"),
      evezEdenVezife: evezEdenVez,
      evezEdilenKod: normCode(getCol(hm, r, "EVEZ_ETDIYI_SEXS_PERSONAL_KOD")),
      evezEdilenAd: getColStr(hm, r, "EVEZ_ETDIYI_FULL_NAME"),
      evezEdilenStruktur: getColStr(hm, r, "STRUKTUR.1"),
      evezEdilenVezife: evezEdilenVez,
      teyinYeri,
      basOfis,
      evezetmeNovu: "Məzuniyyət/xəstəlik əvəzetməsi",
      baslama: bas,
      bitme: bit,
      xerc: isburaxmaCost(basOfis, evezEdenVez, evezEdilenVez),
      qeydler,
      dayBased: true,
    });
  }
}

// ────────────────────────────────────────────────────────────────────────────
// 3.3 Saatlıq icazə
// ────────────────────────────────────────────────────────────────────────────

const TIME_RE = /^\s*(\d{1,2}):(\d{2})\s+(\d{1,2}):(\d{2})\s*$/;

function normSaatliq(ctx: Ctx, wb: XLSX.WorkBook, backups: BackupSpec[]) {
  const MENBE: SourceName = "Saatlıq icazə";
  const sn = findSheet(wb, ["time off replacement"]);
  if (!sn) throw new Error('İşburaxma faylında "time off replacement" sheet-i tapılmadı.');
  const rows = sheetToRows(wb.Sheets[sn]);
  const header = rows[0] ?? [];
  const hm = headerMap(header);

  backups.push({
    sheetName: "BK_Saatlıq icazə",
    note: 'Mənbə: "time off replacement" sheet-i — dəyişdirilməmiş surət.',
    aoa: rows,
  });

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const excelRow = i + 1;
    if (!r || isRowEmpty(r)) continue;

    const day = parseDateValue(getCol(hm, r, "TIME_OFF_DATE"));
    if (day === null) {
      pushErr(ctx, MENBE, excelRow, "Başlama tarixi boşdur", "TIME_OFF_DATE boşdur — sətir təmiz dataya salınmadı.", header, r);
      continue;
    }
    if (day === "bad") {
      pushErr(ctx, MENBE, excelRow, "Başlama tarixi oxunmadı", `TIME_OFF_DATE oxunmadı: "${getColStr(hm, r, "TIME_OFF_DATE")}"`, header, r);
      continue;
    }
    const dayInt = Math.floor(day);
    if (dayInt < ctx.params.filterStart) {
      pushErr(ctx, MENBE, excelRow, "Filtrdən kənar (2026-dan əvvəl)", `Tarix ${fmtDate(dayInt)} < ${fmtDate(ctx.params.filterStart)}.`, header, r);
      continue;
    }

    const tft = getColStr(hm, r, "TIME_FROM_TO");
    const m = tft.match(TIME_RE);
    if (!m) {
      pushErr(ctx, MENBE, excelRow, "Saat aralığı oxunmadı", `TIME_FROM_TO formatı gözləniləndən fərqlidir: "${tft}"`, header, r);
      continue;
    }
    const startFrac = (+m[1] * 60 + +m[2]) / 1440;
    const endFrac = (+m[3] * 60 + +m[4]) / 1440;
    if (endFrac <= startFrac) {
      pushErr(ctx, MENBE, excelRow, "Saat aralığı məntiqsizdir", `Bitmə saatı başlama saatından böyük deyil: "${tft}"`, header, r);
      continue;
    }

    const qeydler: string[] = [];
    const acting = getColStr(hm, r, "Acting");
    if (acting) qeydler.push(`Acting vəzifə: ${acting}`);

    const teyinYeri = getColStr(hm, r, "STRUCTURE_NAME");
    const basOfis = boFromColumn(getColStr(hm, r, "Baş ofis/filial"), teyinYeri);
    const evezEdenVez = getColStr(hm, r, "REPLACING_POSITION");
    const evezEdilenVez = getColStr(hm, r, "POSITION");

    ctx.rows.push({
      menbe: MENBE,
      menbeSetri: String(excelRow),
      rowNums: [excelRow],
      evezEdenKod: normCode(getCol(hm, r, "REPLACING_EMPLOYEE")),
      evezEdenAd: getColStr(hm, r, "Replacing_full_name"),
      evezEdenStruktur: getColStr(hm, r, "REPLACING_STRUCTURE_NAME"),
      evezEdenVezife: evezEdenVez,
      evezEdilenKod: normCode(getCol(hm, r, "HR_CODE")),
      evezEdilenAd: getColStr(hm, r, "Full_Name"),
      evezEdilenStruktur: getColStr(hm, r, "STRUCTURE_NAME"),
      evezEdilenVezife: evezEdilenVez,
      teyinYeri,
      basOfis,
      evezetmeNovu: "İşdən icazə (saatlıq)",
      baslama: dayInt + startFrac,
      bitme: dayInt + endFrac,
      xerc: isburaxmaCost(basOfis, evezEdenVez, evezEdilenVez),
      qeydler,
      dayBased: false,
    });
  }
}

// ────────────────────────────────────────────────────────────────────────────
// 3.4 Müvəqqəti keçid
// ────────────────────────────────────────────────────────────────────────────

function normMuveqqeti(
  ctx: Ctx,
  wb: XLSX.WorkBook,
  backups: BackupSpec[],
  cancels: CancelInfo[],
): { costCount: number; costUnique: number } {
  const MENBE: SourceName = "Müvəqqəti keçid";
  const sn = findSheet(wb, ["Sheet1"]) ?? wb.SheetNames[0];
  const rows = sheetToRows(wb.Sheets[sn]);
  const header = rows[0] ?? [];
  const hm = headerMap(header); // headerMap özü strip edir (SPEC: artıq boşluqlar)

  backups.push({
    sheetName: "BK_Müvəqqəti keçid",
    note: "Mənbə: Filial baza faylı — dəyişdirilməmiş surət.",
    aoa: rows,
  });

  let costCount = 0;
  const costCodes = new Set<string>();

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const excelRow = i + 1;
    if (!r || isRowEmpty(r)) continue;

    const nov = getColStr(hm, r, "Ezamın Növü");
    const novLow = azLower(nov);
    const kod = normCode(getCol(hm, r, "Personal kod"));

    if (novLow.includes("ləğv") || novLow.includes("dayandır")) {
      // SPEC 3.4 — əvəzetmə deyil: Xəta Loquna + saxlanılan keçidlərə DİQQƏT qeydi (checks mərhələsində)
      const errType = novLow.includes("ləğv")
        ? "Əmr növü: Müvəqqəti keçidin ləğvi"
        : "Əmr növü: Müvəqqəti keçidin dayandırılması";
      pushErr(ctx, MENBE, excelRow, errType, `"${nov}" əmri əvəzetmə deyil — sətir təmiz dataya salınmadı.`, header, r);
      const cBas = parseDateValue(getCol(hm, r, "Ezamın Başlama tarixi"));
      const cBit = parseDateValue(getCol(hm, r, "Ezamın Bitmə tarixi"));
      cancels.push({
        kod,
        baslama: typeof cBas === "number" ? cBas : null,
        bitme: typeof cBit === "number" ? cBit : null,
        nov,
        emr: getColStr(hm, r, "Əmrin nömrəsi"),
        rowNum: excelRow,
      });
      continue;
    }

    const bas = parseDateValue(getCol(hm, r, "Ezamın Başlama tarixi"));
    if (bas === null) {
      pushErr(ctx, MENBE, excelRow, "Başlama tarixi boşdur", "Ezamın başlama tarixi yoxdur — sətir təmiz dataya salınmadı.", header, r);
      continue;
    }
    if (bas === "bad") {
      pushErr(ctx, MENBE, excelRow, "Başlama tarixi oxunmadı", `Başlama tarixi oxunmadı: "${getColStr(hm, r, "Ezamın Başlama tarixi")}"`, header, r);
      continue;
    }
    if (bas < ctx.params.filterStart) {
      pushErr(ctx, MENBE, excelRow, "Filtrdən kənar (2026-dan əvvəl)", `Başlama tarixi ${fmtDate(bas)} < ${fmtDate(ctx.params.filterStart)}.`, header, r);
      continue;
    }

    const qeydler: string[] = [];
    let bit = parseDateValue(getCol(hm, r, "Ezamın Bitmə tarixi"));
    if (bit === "bad") bit = null;
    if (bit === null) {
      qeydler.push(`Bitmə tarixi boşdur (müvəqqəti) – gün sayı ${capText(ctx.params)}-a qədər hesablanıb`);
      pushErr(ctx, MENBE, excelRow, "Bitmə tarixi boşdur", "Ezamın bitmə tarixi boşdur — hesablama sərhədinə qədər hesablanır (sətir təmiz datada saxlanılıb).", header, r);
    }

    let struktur = getColStr(hm, r, "Departament");
    if (struktur === "#" || !struktur) struktur = getColStr(hm, r, "Şöbə");
    const teyinYeri = getColStr(hm, r, "Ezam olunduğu departament/filial");
    const teyinLow = azLower(teyinYeri);
    const xerc = MUVEQQETI_COST_TEXTS.some((t) => teyinLow.includes(azLower(t)));
    if (xerc) {
      costCount++;
      if (kod) costCodes.add(kod);
    }

    ctx.rows.push({
      menbe: MENBE,
      menbeSetri: String(excelRow),
      rowNums: [excelRow],
      evezEdenKod: kod,
      evezEdenAd: getColStr(hm, r, "Soyad Ad Ata adı"),
      evezEdenStruktur: struktur,
      evezEdenVezife: getColStr(hm, r, "Vəzifə"),
      evezEdilenKod: "",
      evezEdilenAd: "",
      evezEdilenStruktur: "",
      evezEdilenVezife: "",
      teyinYeri,
      basOfis: classifyBO(teyinYeri || struktur),
      evezetmeNovu: nov || "Müvəqqəti keçid",
      baslama: bas,
      bitme: bit,
      xerc,
      qeydler,
      dayBased: true,
    });
  }

  return { costCount, costUnique: costCodes.size };
}

// ────────────────────────────────────────────────────────────────────────────
// 3.5 Rəsmi əvəzetmə (həvalə)
// ────────────────────────────────────────────────────────────────────────────

/** Sətirdə qırmızı fon (FF0000) xanası varmı — openpyxl analoqu (SPEC 3.5). */
function rowIsRed(ws: XLSX.WorkSheet, rowIdx0: number, colCount: number): boolean {
  for (let c = 0; c < colCount; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: rowIdx0, c })] as
      | { s?: { fgColor?: { rgb?: string } } }
      | undefined;
    const rgb = cell?.s?.fgColor?.rgb;
    if (rgb && String(rgb).toUpperCase().includes("FF0000")) return true;
  }
  return false;
}

function normResmi(ctx: Ctx, wb: XLSX.WorkBook, backups: BackupSpec[], redRows: number[]) {
  const MENBE: SourceName = "Rəsmi əvəzetmə (həvalə)";
  const sn = findSheet(wb, ["Select d_position_info"]) ?? wb.SheetNames[0];
  const ws = wb.Sheets[sn];
  const rows = sheetToRows(ws);
  const header = rows[0] ?? [];
  const hm = headerMap(header);

  // rows indeksi i (0-əsaslı, başlıq i=0) → Excel sətri i+1
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || isRowEmpty(r)) continue;
    if (rowIsRed(ws, i, Math.max(header.length, r.length))) redRows.push(i + 1);
  }

  backups.push({
    sheetName: "BK_Rəsmi əvəzetmə (həvalə)",
    note:
      "Mənbə: Rəsmi əvəzetmə (position info) faylı — dəyişdirilməmiş surət." +
      (redRows.length
        ? ` Orijinalda qırmızı işarələnmiş sətirlər (Excel): ${redRows.join(", ")}.`
        : ""),
    aoa: rows,
  });

  const redSet = new Set(redRows);

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const excelRow = i + 1;
    if (!r || isRowEmpty(r)) continue;

    if (redSet.has(excelRow)) {
      pushErr(ctx, MENBE, excelRow, "Qırmızı işarələnmiş sətir", "Sətir mənbə faylda qırmızı fonla xətalı kimi işarələnib — təmiz dataya salınmadı.", header, r);
      continue;
    }

    const bas = parseDateValue(getCol(hm, r, "ACTING_BEGIN_DATE"));
    if (bas === null) {
      pushErr(ctx, MENBE, excelRow, "Acting başlama tarixi boşdur", "ACTING_BEGIN_DATE boşdur — sətir təmiz dataya salınmadı.", header, r);
      continue;
    }
    if (bas === "bad") {
      pushErr(ctx, MENBE, excelRow, "Başlama tarixi oxunmadı", `ACTING_BEGIN_DATE oxunmadı: "${getColStr(hm, r, "ACTING_BEGIN_DATE")}"`, header, r);
      continue;
    }
    if (bas < ctx.params.filterStart) {
      pushErr(ctx, MENBE, excelRow, "Filtrdən kənar (2026-dan əvvəl)", `ACTING_BEGIN_DATE ${fmtDate(bas)} < ${fmtDate(ctx.params.filterStart)}.`, header, r);
      continue;
    }

    const qeydler: string[] = [];
    let bit = parseDateValue(getCol(hm, r, "ACTING_END_DATE"));
    if (bit === "bad") bit = null;
    if (bit !== null && serialYear(bit) >= 2999) {
      qeydler.push(
        `Bitmə tarixi 31.12.2999 (açıq müddət/davam edir) – gün sayı ${capText(ctx.params)}-a qədər hesablanıb`,
      );
      bit = null;
    } else if (bit === null) {
      qeydler.push(`Bitmə tarixi boşdur (müvəqqəti) – gün sayı ${capText(ctx.params)}-a qədər hesablanıb`);
      pushErr(ctx, MENBE, excelRow, "Bitmə tarixi boşdur", "ACTING_END_DATE boşdur — hesablama sərhədinə qədər hesablanır (sətir təmiz datada saxlanılıb).", header, r);
    }

    const acting = getColStr(hm, r, "ACTING");
    const teyinYeri = getColStr(hm, r, "ACTING_STRUCTURE");

    ctx.rows.push({
      menbe: MENBE,
      menbeSetri: String(excelRow),
      rowNums: [excelRow],
      evezEdenKod: normCode(getCol(hm, r, "HR_CODE")),
      evezEdenAd: getColStr(hm, r, "FULL_NAME"),
      evezEdenStruktur: getColStr(hm, r, "STRUCTURE_NAME"),
      evezEdenVezife: getColStr(hm, r, "POSITION"),
      evezEdilenKod: "",
      evezEdilenAd: "",
      evezEdilenStruktur: getColStr(hm, r, "ACTING_STRUCTURE"),
      evezEdilenVezife: acting ? `${acting} (həvalə vəzifəsi)` : "",
      teyinYeri,
      basOfis: boFromColumn(getColStr(hm, r, "Baş ofis/Filial"), teyinYeri),
      evezetmeNovu: `Rəsmi əvəzetmə / acting: ${acting}`,
      baslama: bas,
      bitme: bit,
      xerc: true, // SPEC 5.1 — qırmızılar çıxarıldıqdan sonra hamısı Bəli
      qeydler,
      dayBased: true,
    });
  }
}

// ────────────────────────────────────────────────────────────────────────────

export function normalizeAll(src: SourceWorkbooks, params: ReplacementParams): NormalizeResult {
  const ctx: Ctx = { params, rows: [], errors: [] };
  const backups: BackupSpec[] = [];
  const cancels: CancelInfo[] = [];
  const redRows: number[] = [];

  const vm = buildVMLookup(src.vezifeMaas);

  normInzibatci(ctx, src.inzibatci, vm, backups);
  normIsburaxma(ctx, src.vacation, backups);
  normSaatliq(ctx, src.vacation, backups);
  const muv = normMuveqqeti(ctx, src.filial, backups, cancels);
  normResmi(ctx, src.resmi, backups, redRows);

  // İşburaxma faylının texniki vərəqləri (SPEC 10)
  const sqlSheet = findSheet(src.vacation, ["SQL Statement"]);
  if (sqlSheet) {
    backups.push({
      sheetName: "BK_SQL Statement",
      note: "Mənbə: İşburaxma faylındakı texniki vərəq — dəyişdirilməmiş surət.",
      aoa: sheetToRows(src.vacation.Sheets[sqlSheet]),
    });
  }
  const sheet2 = findSheet(src.vacation, ["Sheet2"]);
  if (sheet2) {
    const aoa = sheetToRows(src.vacation.Sheets[sheet2]).filter((r) => r && !isRowEmpty(r));
    if (aoa.length) {
      backups.push({
        sheetName: "BK_Sheet2 (işburaxma faylı)",
        note: "Mənbə: İşburaxma faylındakı Sheet2 — dəyişdirilməmiş surət.",
        aoa,
      });
    }
  }
  if (vm.backupAoa.length) {
    backups.push({
      sheetName: "BK_Vəzifə-Maaş",
      note: "Mənbə: Vəzifə-Maaş faylı — 6 aylıq sheet bir vərəqdə alt-alta (Ay və Ay tarixi sütunları əlavə edilib).",
      aoa: [["Ay", "Ay tarixi", ...vm.backupHeader], ...vm.backupAoa],
    });
  }

  const sourceCounts = {
    "İnzibatçı prosesi": 0,
    "İşburaxma hallarının əvəzedilməsi": 0,
    "Saatlıq icazə": 0,
    "Müvəqqəti keçid": 0,
    "Rəsmi əvəzetmə (həvalə)": 0,
  } as Record<SourceName, number>;
  for (const r of ctx.rows) sourceCounts[r.menbe]++;

  return {
    rows: ctx.rows,
    errors: ctx.errors,
    cancels,
    redRows,
    backups,
    sourceCounts,
    muveqqetiCostCount: muv.costCount,
    muveqqetiCostUnique: muv.costUnique,
  };
}
