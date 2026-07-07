// "Vacation Days by Period" hesabatı — vacation_days_by_period_SPEC.md əsasında.
// Hər əməkdaş × hər iş ili (work-year) üçün bir sətir: tam illik hüquq (proporsional
// accrual deyil), istifadə FIFO ilə ən köhnə dövrdən silinir.
import * as XLSX from "xlsx";
import type { Cell, Row } from "./replacement/types";
import {
  cellStr,
  dateFromSerial,
  fmtDate,
  headerMap,
  parseDateValue,
  sheetToRows,
} from "./replacement/xl";
import { EXCEL_EPOCH_MS, MS_PER_DAY } from "./replacement/config";

export type ProgressFn = (step: string, pct: number) => void;

export interface WorkyearParams {
  /** Dövr generasiyasının son həddi (Excel serial). */
  asOfDate: number;
  /** Son (natamam) dövrün hüququnu AS_OF_DATE-ə qədər proporsional hesabla. */
  prorateLastPeriod: boolean;
  /** "Компенсация" əmrlərini istifadəyə say. */
  includeCompensation: boolean;
  /** Əməkdaş reyestri səhifəsi. */
  masterSheet: string;
}

export function todaySerial(): number {
  const d = new Date();
  return Math.round((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - EXCEL_EPOCH_MS) / MS_PER_DAY);
}

export const DEFAULT_WORKYEAR_PARAMS: WorkyearParams = {
  asOfDate: 0, // 0 → işə salınanda "bugün" götürülür
  prorateLastPeriod: false,
  includeCompensation: true,
  masterSheet: "2015-2026",
};

export interface PeriodRow {
  empKey: string;
  empCode: string;
  fullName: string;
  structure: string;
  section: string;
  position: string;
  hireDate: number;
  periodStart: number;
  periodEnd: number;
  mainDays: number; // I
  addDays: number; // J
  usedMain: number; // L
  usedAdd: number; // N
}

export interface WyLogEntry {
  tip: string;
  emp: string;
  detail: string;
}

export interface WyCheck {
  name: string;
  ok: boolean;
  detail: string;
  hard: boolean;
}

export interface WorkyearResult {
  workbook: XLSX.WorkBook;
  rows: PeriodRow[];
  logs: WyLogEntry[];
  checks: WyCheck[];
  allPassed: boolean;
  stats: {
    employees: number;
    periods: number;
    ordersIncluded: number;
    ordersExcluded: number;
    ordersDedup: number;
    totalUsed: number;
    yearSheets: string[];
    matchPct: number;
    anomalies: number;
  };
  params: WorkyearParams;
  durationMs: number;
}

// ── köməkçilər ───────────────────────────────────────────────────────────────

function nkeyOf(name: string): string {
  return name.toUpperCase().replace(/\s+/g, "");
}

function normBadge(b: string): string {
  return b.toUpperCase().replace(/\s+/g, "");
}

function toNum(v: Cell): number {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const n = Number(String(v).replace(",", "."));
  return Number.isFinite(n) ? n : 0;
}

function toNumOrNull(v: Cell): number | null {
  if (v === null || v === undefined || v === "") return null;
  const n = typeof v === "number" ? v : Number(String(v).replace(",", "."));
  return Number.isFinite(n) ? n : null;
}

function parseD(v: Cell): number | null {
  const p = parseDateValue(v);
  return typeof p === "number" ? Math.floor(p) : null;
}

/** relativedelta(years=n) analoqu: 29 fevral → 28 fevral (qeyri-uzun illərdə). */
export function addYears(serial: number, n: number): number {
  const d = dateFromSerial(serial);
  const y = d.getUTCFullYear() + n;
  const m = d.getUTCMonth();
  let day = d.getUTCDate();
  if (m === 1 && day === 29) {
    const leap = (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
    if (!leap) day = 28;
  }
  return Math.round((Date.UTC(y, m, day) - EXCEL_EPOCH_MS) / MS_PER_DAY);
}

interface MasterEmp {
  key: string;
  nkey: string;
  badge: string;
  badgeNorm: string;
  name: string;
  hire: number | null;
  term: number | null;
  structure: string;
  section: string;
  position: string;
  masterMain: number;
  masterAdd: number;
  masterBalance: number | null;
}

interface YearNorm {
  main: number;
  add: number;
  used: number | null;
}

// ── əsas proses ──────────────────────────────────────────────────────────────

export async function processWorkyear(
  file: File,
  paramsIn: WorkyearParams = DEFAULT_WORKYEAR_PARAMS,
  onProgress: ProgressFn = () => {},
): Promise<WorkyearResult> {
  const t0 = typeof performance !== "undefined" ? performance.now() : Date.now();
  const now = () => (typeof performance !== "undefined" ? performance.now() : Date.now());
  const params: WorkyearParams = {
    ...paramsIn,
    asOfDate: paramsIn.asOfDate > 0 ? paramsIn.asOfDate : todaySerial(),
  };
  const logs: WyLogEntry[] = [];
  const log = (tip: string, emp: string, detail: string) => logs.push({ tip, emp, detail });

  onProgress("Excel faylı oxunur...", 5);
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: false });

  // ── Addım 1: master siyahı ──────────────────────────────────────────────────
  onProgress("Addım 1: Master siyahı və açar normalizasiyası...", 15);
  const masterName =
    wb.SheetNames.find((n) => n.trim() === params.masterSheet.trim()) ??
    wb.SheetNames.find((n) => n.trim().toLowerCase() === params.masterSheet.trim().toLowerCase());
  if (!masterName) {
    throw new Error(
      `Master səhifə "${params.masterSheet}" tapılmadı. Mövcud səhifələr: ${wb.SheetNames.join(", ")}`,
    );
  }
  const masterRows = sheetToRows(wb.Sheets[masterName]);

  const employees: MasterEmp[] = [];
  for (let i = 2; i < masterRows.length; i++) {
    const r = masterRows[i];
    if (!r) continue;
    const badge = cellStr(r[0]);
    const name = cellStr(r[1]);
    if (!name && !badge) continue;
    if (!name) {
      log("Ad boşdur", badge, `Master sətir ${i + 1}: ad olmadan sətir ötürüldü.`);
      continue;
    }
    const hire = parseD(r[3]);
    if (hire === null) {
      log("Parse olunmayan tarix", name, `Master sətir ${i + 1}: işə qəbul tarixi oxunmadı ("${cellStr(r[3])}").`);
    }
    employees.push({
      key: nkeyOf(name),
      nkey: nkeyOf(name),
      badge,
      badgeNorm: normBadge(badge),
      name,
      hire,
      term: parseD(r[4]),
      structure: cellStr(r[6]),
      section: cellStr(r[7]),
      position: cellStr(r[10]),
      masterMain: toNum(r[14]),
      masterAdd: toNum(r[15]) + toNum(r[16]) + toNum(r[17]),
      masterBalance: toNumOrNull(r[22]),
    });
  }
  if (!employees.length) throw new Error(`"${masterName}" səhifəsində əməkdaş tapılmadı (data 3-cü sətirdən gözlənilir).`);

  // nkey unikallığı: unikal deyilsə nkey+badge kombinasiyasına keç
  const byNkey = new Map<string, MasterEmp[]>();
  for (const e of employees) {
    const l = byNkey.get(e.nkey) ?? [];
    l.push(e);
    byNkey.set(e.nkey, l);
  }
  for (const [nk, list] of byNkey) {
    if (list.length > 1) {
      for (const e of list) e.key = `${e.nkey}|${e.badgeNorm}`;
      log("Təkrar ad (nkey)", nk, `${list.length} əməkdaş eyni ada malikdir — uyğunlaşdırma "ad + badge" ilə aparılır (badge-lər: ${list.map((e) => e.badge).join(", ")}).`);
    }
  }
  const resolveEmp = (name: string, badge: string): MasterEmp | null => {
    const list = byNkey.get(nkeyOf(name));
    if (!list) return null;
    if (list.length === 1) return list[0];
    const b = normBadge(badge);
    return list.find((e) => e.badgeNorm === b) ?? null;
  };

  // ── Addım 2: illik normalar lüğəti ─────────────────────────────────────────
  onProgress("Addım 2: İllik normalar lüğəti (rehire filtri ilə)...", 30);
  const yearSheets = wb.SheetNames.filter((n) => /^\d{4}$/.test(n.trim())).sort();
  const norms = new Map<string, Map<number, YearNorm>>(); // empKey → il → norma

  for (const sn of yearSheets) {
    const year = parseInt(sn.trim(), 10);
    const rows = sheetToRows(wb.Sheets[sn]);
    // empKey → namizəd sətirlər
    const candidates = new Map<string, { start: number | null; norm: YearNorm }[]>();
    for (let i = 2; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const name = cellStr(r[1]);
      if (!name) continue;
      const emp = resolveEmp(name, cellStr(r[0]));
      if (!emp) continue;
      const l = candidates.get(emp.key) ?? [];
      l.push({
        start: parseD(r[3]),
        norm: {
          main: toNum(r[14]),
          add: toNum(r[15]) + toNum(r[16]) + toNum(r[17]),
          used: toNumOrNull(r[21]),
        },
      });
      candidates.set(emp.key, l);
    }
    for (const [key, list] of candidates) {
      const emp = employees.find((e) => e.key === key)!;
      // Rehire filtri: yalnız Start == master Start; tapılmasa istənilən sətir fallback
      const matched = emp.hire !== null ? list.filter((c) => c.start !== null && c.start === emp.hire) : [];
      const chosen = matched.length ? matched[0] : list[0];
      if (!matched.length && emp.hire !== null && list.some((c) => c.start !== null)) {
        log("Rehire fallback", emp.name, `${sn} səhifəsində master işə qəbul tarixinə (${fmtDate(emp.hire)}) uyğun sətir yoxdur — mövcud sətir fallback kimi götürüldü.`);
      }
      const m = norms.get(key) ?? new Map<number, YearNorm>();
      m.set(year, chosen.norm);
      norms.set(key, m);
    }
  }

  // ── Addım 3: əmrlərin filtrasiyası ─────────────────────────────────────────
  onProgress("Addım 3: Əmrlərin filtrasiyası (FIFO üçün istifadə)...", 45);
  const ordersName = wb.SheetNames.find((n) => n.trim().toLowerCase() === "orders");
  if (!ordersName) throw new Error('"orders" səhifəsi tapılmadı.');
  const oRows = sheetToRows(wb.Sheets[ordersName]);
  const ohm = headerMap(oRows[0] ?? []);
  const col = (r: Row, name: string): Cell => {
    const idx = ohm.get(name);
    return idx === undefined ? null : r[idx];
  };
  const INCLUDE_TYPES = ["отпуск очередной", "компенсация"];

  const usedByEmp = new Map<string, number>();
  const seenOrder = new Set<string>();
  let ordersIncluded = 0;
  let ordersExcluded = 0;
  let ordersDedup = 0;
  let totalUsedOrders = 0;

  for (let i = 1; i < oRows.length; i++) {
    const r = oRows[i];
    if (!r) continue;
    const name = cellStr(col(r, "Работник"));
    if (!name) continue;
    const tip = cellStr(col(r, "Поле1")).toLowerCase().trim();
    const approved = cellStr(col(r, "Готов к утверждению")).toLowerCase().trim() === "да";
    const isComp = tip === "компенсация";
    const included = INCLUDE_TYPES.includes(tip) && approved && (params.includeCompensation || !isComp);
    if (!included) {
      ordersExcluded++;
      continue;
    }
    const emp = resolveEmp(name, cellStr(col(r, "Badge")));
    if (!emp) {
      log("Master siyahıda yoxdur", name, `orders sətri ${i + 1}: əməkdaş master siyahıda tapılmadı — əmr nəzərə alınmır.`);
      ordersExcluded++;
      continue;
    }
    const from = parseD(col(r, "from"));
    const to = parseD(col(r, "to"));
    const dn = toNum(col(r, "(дн)"));
    const dedupKey = `${emp.key}|${from ?? ""}|${to ?? ""}|${dn}`;
    if (seenOrder.has(dedupKey)) {
      ordersDedup++;
      log("Dublikat əmr", emp.name, `orders sətri ${i + 1}: (${from !== null ? fmtDate(from) : "?"}–${to !== null ? fmtDate(to) : "?"}, ${dn} gün) təkrar əmr — bir dəfə sayıldı.`);
      continue;
    }
    seenOrder.add(dedupKey);
    const effD = from ?? parseD(col(r, "Дата приказа"));
    if (effD === null) {
      log("Parse olunmayan tarix", emp.name, `orders sətri ${i + 1}: nə "from", nə "Дата приказа" oxundu — əmr nəzərə alınmır.`);
      ordersExcluded++;
      continue;
    }
    // Cari iş dövrü filtri: rehire-dən əvvəlki istifadələr sayılmır
    if (emp.hire !== null && effD < emp.hire) {
      ordersExcluded++;
      log("Əvvəlki iş dövrü", emp.name, `orders sətri ${i + 1}: effektiv tarix ${fmtDate(effD)} master işə qəbuldan (${fmtDate(emp.hire)}) əvvəldir — cari balansa sayılmır.`);
      continue;
    }
    ordersIncluded++;
    totalUsedOrders += dn;
    usedByEmp.set(emp.key, (usedByEmp.get(emp.key) ?? 0) + dn);
  }

  // ── Addım 4–5: iş illəri + FIFO ────────────────────────────────────────────
  onProgress("Addım 4–5: İş illəri və FIFO bölüşdürməsi...", 60);
  const rows: PeriodRow[] = [];
  const asOf = params.asOfDate;

  for (const emp of employees) {
    if (emp.hire === null) {
      log("Dövr generasiya edilmədi", emp.name, "İşə qəbul tarixi oxunmadığı üçün dövrlər yaradıla bilmədi.");
      continue;
    }
    const periods: { start: number; end: number; main: number; add: number }[] = [];
    const yearMap = norms.get(emp.key);
    const availYears = yearMap ? [...yearMap.keys()].sort((a, b) => a - b) : [];
    for (let i = 0; addYears(emp.hire, i) <= asOf && i < 120; i++) {
      const start = addYears(emp.hire, i);
      if (emp.term !== null && start > emp.term) break; // işdən çıxmadan sonra dövr yaranmır
      const end = addYears(emp.hire, i + 1) - 1;
      const targetYear = dateFromSerial(start).getUTCFullYear();
      let norm: YearNorm | null = null;
      if (yearMap && availYears.length) {
        // o il yoxdursa → ən yaxın mövcud il
        const y = yearMap.has(targetYear)
          ? targetYear
          : availYears.reduce((best, cur) =>
              Math.abs(cur - targetYear) < Math.abs(best - targetYear) ||
              (Math.abs(cur - targetYear) === Math.abs(best - targetYear) && cur < best)
                ? cur
                : best,
            );
        norm = yearMap.get(y) ?? null;
      }
      let main: number;
      let add: number;
      if (norm) {
        main = norm.main;
        add = norm.add;
      } else {
        main = emp.masterMain;
        add = emp.masterAdd;
        if (main === 0 && add === 0) {
          log("Norma tapılmadı", emp.name, `${fmtDate(start)}–${fmtDate(end)} dövrü üçün nə illik səhifədə, nə masterdə norma tapıldı (0 götürüldü).`);
        }
      }
      periods.push({ start, end, main, add });
    }
    if (!periods.length) {
      log("Dövr generasiya edilmədi", emp.name, `İşə qəbul (${fmtDate(emp.hire)}) AS_OF (${fmtDate(asOf)}) tarixindən sonradır.`);
      continue;
    }

    // Cari natamam dövr üçün proporsional rejim (opsiya)
    if (params.prorateLastPeriod) {
      const last = periods[periods.length - 1];
      if (last.end > asOf) {
        const frac = (asOf - last.start + 1) / (last.end - last.start + 1);
        last.main = Math.round(last.main * frac * 100) / 100;
        last.add = Math.round(last.add * frac * 100) / 100;
      }
    }

    // Addım 5 — FIFO
    let rem = usedByEmp.get(emp.key) ?? 0;
    const fifo = periods.map((p) => {
      const usedMain = Math.min(rem, p.main);
      rem = Math.round((rem - usedMain) * 100) / 100;
      const usedAdd = Math.min(rem, p.add);
      rem = Math.round((rem - usedAdd) * 100) / 100;
      return { ...p, usedMain, usedAdd };
    });
    if (rem > 0) {
      fifo[fifo.length - 1].usedMain += rem; // avans məzuniyyət — mənfi unused normaldır
      log("Avans məzuniyyət", emp.name, `İstifadə hüquqdan ${rem} gün çoxdur — son dövrün istifadəsinə əlavə edildi (unused mənfi olacaq).`);
      rem = 0;
    }

    for (const p of fifo) {
      rows.push({
        empKey: emp.key,
        empCode: emp.badge,
        fullName: emp.name,
        structure: emp.structure,
        section: emp.section,
        position: emp.position,
        hireDate: emp.hire,
        periodStart: p.start,
        periodEnd: p.end,
        mainDays: p.main,
        addDays: p.add,
        usedMain: Math.round(p.usedMain * 100) / 100,
        usedAdd: Math.round(p.usedAdd * 100) / 100,
      });
    }
  }

  // ── Addım 6: nəticə faylı ──────────────────────────────────────────────────
  onProgress("Addım 6: Nəticə faylı yazılır...", 78);
  const out = buildOutput(rows);

  // ── Addım 7: validasiya ────────────────────────────────────────────────────
  onProgress("Addım 7: Validasiya...", 90);
  const checks: WyCheck[] = [];

  // 1. İstifadə cəmi — dəqiq bərabərlik
  const sumUsedRows = Math.round(rows.reduce((s, r) => s + r.usedMain + r.usedAdd, 0) * 100) / 100;
  const sumOrders = Math.round(totalUsedOrders * 100) / 100;
  checks.push({
    name: "İstifadə cəmi (FIFO invariantı)",
    ok: Math.abs(sumUsedRows - sumOrders) < 0.005,
    detail: `Nəticədə Σ(used) = ${sumUsedRows}; filtrlənmiş əmrlərdə Σ(дн) = ${sumOrders} — dəqiq bərabər olmalıdır.`,
    hard: true,
  });

  // 2. İllik səhifə uzlaşması (≥95%)
  let cmp = 0;
  let match = 0;
  for (const emp of employees) {
    const ym = norms.get(emp.key);
    if (!ym) continue;
    let yearUsed = 0;
    let any = false;
    for (const n of ym.values()) {
      if (n.used !== null) {
        yearUsed += n.used;
        any = true;
      }
    }
    if (!any) continue;
    cmp++;
    const used = usedByEmp.get(emp.key) ?? 0;
    if (Math.abs(used - yearUsed) < 0.01) match++;
    else log("Used uzlaşmır", emp.name, `Əmrlərdən istifadə ${used}, illik səhifələrin "Used" cəmi ${yearUsed} (fərq ${Math.round((used - yearUsed) * 100) / 100}).`);
  }
  const matchPct = cmp ? Math.round((match / cmp) * 10000) / 100 : 100;
  checks.push({
    name: "İllik səhifə uzlaşması",
    ok: matchPct >= 95,
    detail: `${cmp} əməkdaşdan ${match} uyğundur (${matchPct}%; tələb ≥95%). Fərqlər anomaliya loqundadır (əsas səbəb: rehire kompensasiya qeydləri).`,
    hard: true,
  });

  // 3. Balans uzlaşması — məlumat xarakterli
  {
    let cmpB = 0;
    let exact = 0;
    for (const emp of employees) {
      if (emp.masterBalance === null) continue;
      cmpB++;
      const unused = rows
        .filter((r) => r.empKey === emp.key)
        .reduce((s, r) => s + (r.mainDays + r.addDays - r.usedMain - r.usedAdd), 0);
      if (Math.abs(unused - emp.masterBalance) < 0.01) exact++;
    }
    checks.push({
      name: "Balans uzlaşması (informativ)",
      ok: true,
      detail: cmpB
        ? `${cmpB} əməkdaşdan ${exact} tam üst-üstə düşür. Fərq gözləniləndir: mənbə proporsional accrual, bu hesabat tam illik hüquq metodu istifadə edir — fərq cari natamam dövrün payı qədərdir.`
        : "Master səhifədə Balance Ending dəyəri tapılmadı — müqayisə aparılmadı.",
      hard: false,
    });
  }

  // 4. Əhatə
  const covered = new Set(rows.map((r) => r.empKey));
  const missing = employees.filter((e) => !covered.has(e.key));
  checks.push({
    name: "Əhatə",
    ok: missing.length === 0,
    detail:
      missing.length === 0
        ? `${employees.length} master əməkdaşın hamısının nəticədə ən azı 1 dövrü var (cəmi ${rows.length} dövr).`
        : `${missing.length} əməkdaşın dövrü yoxdur: ${missing.slice(0, 5).map((e) => e.name).join("; ")}${missing.length > 5 ? " ..." : ""} (səbəblər anomaliya loqundadır).`,
    hard: true,
  });

  // Düstur xanaları (recalc əvəzi: cache edilmiş dəyərlər yazılır)
  {
    const ws = out.Sheets["Vacation days by period"];
    let fCells = 0;
    let fBad = 0;
    const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1");
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (!cell || !cell.f) continue;
        fCells++;
        if (typeof cell.v !== "number" || !Number.isFinite(cell.v)) fBad++;
      }
    }
    checks.push({
      name: "Düstur xanaları (recalc əvəzi)",
      ok: fBad === 0,
      detail: `${fCells} düstur xanası (K, M, O, P, Q); ${fBad} xətalı dəyər (tələb: 0). Düsturlar cache edilmiş dəyərlə yazılıb — Excel açılışda yenidən hesablayır.`,
      hard: true,
    });
  }

  // 5. Anomaliya loqu — məlumat
  const anomalyTips = ["Avans məzuniyyət", "Norma tapılmadı", "Parse olunmayan tarix", "Dublikat əmr", "Təkrar ad (nkey)"];
  const anomalies = logs.filter((l) => anomalyTips.includes(l.tip)).length;
  checks.push({
    name: "Anomaliya loqu",
    ok: true,
    detail: `${logs.length} log qeydi (${anomalyTips.map((t) => `${t}: ${logs.filter((l) => l.tip === t).length}`).join(", ")}).`,
    hard: false,
  });

  onProgress("Tamamlandı", 100);

  return {
    workbook: out,
    rows,
    logs,
    checks,
    allPassed: checks.filter((c) => c.hard).every((c) => c.ok),
    stats: {
      employees: employees.length,
      periods: rows.length,
      ordersIncluded,
      ordersExcluded,
      ordersDedup,
      totalUsed: sumOrders,
      yearSheets,
      matchPct,
      anomalies,
    },
    params,
    durationMs: now() - t0,
  };
}

// ── Addım 6: A–Q sütunlu tək səhifəli nəticə faylı ──────────────────────────

const HEADER = [
  "Emp_code",
  "Full name",
  "Structure name",
  "Section name",
  "Position",
  "Hire date",
  "Period start date",
  "Period end date",
  "Main vacation days",
  "Total additional vacation days",
  "Total annual entitlement",
  "Used days for main vacation",
  "Unused days for main vacation",
  "Used days for additional vacation",
  "Unused days for additional vacation",
  "Total used vacation days",
  "Total unused vacation days",
];

function buildOutput(rows: PeriodRow[]): XLSX.WorkBook {
  const aoa: Row[] = [HEADER];
  for (const r of rows) {
    aoa.push([
      r.empCode,
      r.fullName,
      r.structure,
      r.section,
      r.position,
      r.hireDate,
      r.periodStart,
      r.periodEnd,
      r.mainDays,
      r.addDays,
      r.mainDays + r.addDays, // K =I+J
      r.usedMain,
      Math.round((r.mainDays - r.usedMain) * 100) / 100, // M =I-L
      r.usedAdd,
      Math.round((r.addDays - r.usedAdd) * 100) / 100, // O =J-N
      Math.round((r.usedMain + r.usedAdd) * 100) / 100, // P =L+N
      Math.round((r.mainDays + r.addDays - r.usedMain - r.usedAdd) * 100) / 100, // Q =K-P
    ]);
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  for (let i = 0; i < rows.length; i++) {
    const r1 = i + 2;
    for (const c of [5, 6, 7]) {
      const cell = ws[XLSX.utils.encode_cell({ r: r1 - 1, c })];
      if (cell && typeof cell.v === "number") cell.z = "dd.mm.yyyy";
    }
    const setF = (c0: number, f: string) => {
      const cell = ws[XLSX.utils.encode_cell({ r: r1 - 1, c: c0 })];
      if (cell) cell.f = f;
    };
    setF(10, `I${r1}+J${r1}`);
    setF(12, `I${r1}-L${r1}`);
    setF(14, `J${r1}-N${r1}`);
    setF(15, `L${r1}+N${r1}`);
    setF(16, `K${r1}-P${r1}`);
  }
  ws["!cols"] = [
    { wch: 12 }, { wch: 32 }, { wch: 26 }, { wch: 26 }, { wch: 26 }, { wch: 12 }, { wch: 15 }, { wch: 15 },
    { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 14 },
  ];
  ws["!autofilter"] = { ref: `A1:Q${Math.max(rows.length + 1, 1)}` };
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Vacation days by period");
  return wb;
}

export function exportWorkyear(result: WorkyearResult, filename = "Vacation_days_by_period.xlsx") {
  XLSX.writeFile(result.workbook, filename, { compression: true });
}
