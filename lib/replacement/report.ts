import * as XLSX from "xlsx";
import { SOURCE_ORDER, SOURCE_PRIORITY, type ReplacementParams, type SourceName } from "./config";
import type { ChecksCounters } from "./checks";
import { effectiveEnd } from "./checks";
import type { BackupSpec, CemRow, Cell, DupEntry, ErrorLogEntry, NormRow, Row } from "./types";
import { fmtDate, isWeekend } from "./xl";

export interface ReportInput {
  clean: NormRow[];
  errors: ErrorLogEntry[];
  dups: DupEntry[];
  backups: BackupSpec[];
  counters: ChecksCounters;
  sourceCounts: Record<SourceName, number>;
  redRows: number[];
  params: ReplacementParams;
}

export interface StatCard {
  label: string;
  value: number;
}

export interface ReportOutput {
  wb: XLSX.WorkBook;
  yekun: NormRow[];
  cem: CemRow[];
  statCards: StatCard[];
  yekunExcludedXeta: number;
  yekunExcludedEpizod: number;
  yekunExcludedLegv: number;
  mergedSourceRows: number;
}

const YEKUN_SHEET = "Xərc yaradan (yekun)";
const TEMIZ_SHEET = "Təmiz Data";
const CEM_SHEET = "Əməkdaş üzrə cəm";
const XETA_SHEET = "Xəta Loqu";
const DUP_SHEET = "Duplikatlar";

function hasCancelNote(r: NormRow): boolean {
  return r.qeydler.some((q) => q.startsWith("DİQQƏT: bu keçid üzrə"));
}

// ────────────────────────────────────────────────────────────────────────────
// 9.2 — Yekun: Bəli sətirləri + İnzibatçı epizod birləşdirilməsi
// ────────────────────────────────────────────────────────────────────────────

interface YekunBuild {
  yekun: NormRow[];
  excludedXeta: number;
  excludedEpizod: number;
  excludedLegv: number;
  mergedSourceRows: number;
}

export function buildYekun(clean: NormRow[], errors: ErrorLogEntry[], params: ReplacementParams): YekunBuild {
  const excludedXeta = clean.filter((r) => r.xerc && r.hasXeta).length;
  const excludedLegv = clean.filter((r) => r.xerc && !r.hasXeta && hasCancelNote(r)).length;

  const eligible = clean.filter((r) => r.xerc && !r.hasXeta && !hasCancelNote(r));
  const inz = eligible.filter((r) => r.menbe === "İnzibatçı prosesi");
  const rest = eligible.filter((r) => r.menbe !== "İnzibatçı prosesi");

  // Qruplaşdırma açarı: (kod, Təyin yeri, Əvəzetmə növü)
  const groups = new Map<string, NormRow[]>();
  for (const r of inz) {
    const key = `${r.evezEdenKod}|${r.teyinYeri}|${r.evezetmeNovu}`;
    const g = groups.get(key);
    if (g) g.push(r);
    else groups.set(key, [r]);
  }

  interface Episode {
    members: NormRow[];
    open: boolean;
    endReal: number | null;
    effEnd: number;
  }

  let mergedSourceRows = 0;
  const episodesByKod = new Map<string, { ep: NormRow; s: number; e: number }[]>();
  const allEpisodes: NormRow[] = [];

  for (const g of groups.values()) {
    const sorted = [...g].sort((a, b) => a.baslama - b.baslama || a.rowNums[0] - b.rowNums[0]);
    const eps: Episode[] = [];
    let cur: Episode | null = null;
    for (const r of sorted) {
      const rEff = effectiveEnd(r, params);
      if (cur && Math.floor(r.baslama) <= Math.floor(cur.effEnd) + 1) {
        cur.members.push(r);
        cur.open = cur.open || r.bitme === null;
        if (r.bitme !== null) cur.endReal = cur.endReal === null ? r.bitme : Math.max(cur.endReal, r.bitme);
        cur.effEnd = Math.max(cur.effEnd, rEff);
      } else {
        cur = { members: [r], open: r.bitme === null, endReal: r.bitme, effEnd: rEff };
        eps.push(cur);
      }
    }

    for (const ep of eps) {
      const first = ep.members[0];
      const merged = ep.members.length > 1;
      if (merged) mergedSourceRows += ep.members.length;

      const bitme = ep.open ? null : ep.endReal;
      const rowNums = ep.members.flatMap((m) => m.rowNums).sort((a, b) => a - b);
      const qeydler: string[] = [];
      for (const m of ep.members) for (const q of m.qeydler) if (!qeydler.includes(q)) qeydler.push(q);
      if (merged) {
        qeydler.push(
          `Birləşdirilib: davamlı gedən ${ep.members.length} mənbə sətri (sətirlər: ${rowNums.join(", ")}); birinci başlama və son bitmə tarixi götürülüb`,
        );
        if (ep.open) qeydler.push(`Bitmə tarixi açıqdır – ${fmtDate(params.capDate)} kimi qəbul edilib`);
      }

      const row: NormRow = {
        ...first,
        menbeSetri: rowNums.join(", "),
        rowNums,
        baslama: first.baslama,
        bitme,
        qeydler,
        hasXeta: false,
      };
      const e = effectiveEnd(row, params);
      row.effEnd = e;
      row.dayCount = Math.floor(e) - Math.floor(row.baslama) + 1;
      row.hourCount = row.dayCount * 8;

      allEpisodes.push(row);
      const list = episodesByKod.get(row.evezEdenKod) ?? [];
      list.push({ ep: row, s: Math.floor(row.baslama), e: Math.floor(e) });
      episodesByKod.set(row.evezEdenKod, list);
    }
  }

  // Birləşmədən sonra eyni şəxsin epizodları hələ də kəsişirsə → hər ikisi yekundan çıxarılır
  const excluded = new Set<NormRow>();
  for (const list of episodesByKod.values()) {
    if (list.length < 2) continue;
    for (let i = 0; i < list.length; i++) {
      for (let j = i + 1; j < list.length; j++) {
        const a = list[i];
        const b = list[j];
        if (a.s <= b.e && b.s <= a.e) {
          excluded.add(a.ep);
          excluded.add(b.ep);
        }
      }
    }
  }
  for (const ep of excluded) {
    errors.push({
      menbe: ep.menbe,
      setr: ep.menbeSetri,
      novu: "Yekun vərəqdən istisna: kəsişən epizodlar",
      teferruat: `Epizod (${fmtDate(ep.baslama)}–${ep.bitme === null ? "açıq" : fmtDate(ep.bitme)}) eyni şəxsin başqa epizodu ilə kəsişir (fərqli yer/növ üzrə) — hər iki epizod yekun vərəqdən çıxarılıb.`,
      setirMelumati: `Əvəz edən: ${ep.evezEdenAd} (${ep.evezEdenKod}) | Növ: ${ep.evezetmeNovu} | Təyin yeri: ${ep.teyinYeri}`,
    });
  }

  const yekun = [...rest, ...allEpisodes.filter((e) => !excluded.has(e))].sort(
    (a, b) => SOURCE_PRIORITY[a.menbe] - SOURCE_PRIORITY[b.menbe] || a.rowNums[0] - b.rowNums[0],
  );

  return {
    yekun,
    excludedXeta,
    excludedEpizod: excluded.size,
    excludedLegv: excludedLegv,
    mergedSourceRows,
  };
}

// ────────────────────────────────────────────────────────────────────────────
// 9.3 — Əməkdaş üzrə cəm
// ────────────────────────────────────────────────────────────────────────────

export function buildCem(clean: NormRow[], params: ReplacementParams): CemRow[] {
  const byKod = new Map<string, NormRow[]>();
  for (const r of clean) {
    if (!r.evezEdenKod) continue;
    const g = byKod.get(r.evezEdenKod);
    if (g) g.push(r);
    else byKod.set(r.evezEdenKod, [r]);
  }

  const out: CemRow[] = [];
  for (const [kod, rows] of byKod) {
    const ad = rows.find((r) => r.evezEdenAd)?.evezEdenAd ?? "";
    let say = 0;
    let cemSaat = 0;
    let cemGun = 0;
    for (const r of rows) {
      // COUNTIFS/SUMIFS kriteriyası: P >= B2 və P < D2+1
      if (r.baslama >= params.winStart && r.baslama < params.winEnd + 1) {
        say++;
        cemSaat += r.hourCount ?? 0;
        cemGun += r.dayCount ?? 0;
      }
    }
    // Unikal təqvim günü: aralıqların birləşməsi, pəncərəyə kəsilir
    const days = new Set<number>();
    for (const r of rows) {
      const s = Math.max(Math.floor(r.baslama), params.winStart);
      const e = Math.min(Math.floor(r.effEnd ?? effectiveEnd(r, params)), params.winEnd);
      for (let d = s; d <= e; d++) days.add(d);
    }
    let isGunu = 0;
    let hefteSonu = 0;
    for (const d of days) {
      if (isWeekend(d)) hefteSonu++;
      else isGunu++;
    }
    out.push({
      kod,
      ad,
      say,
      cemSaat: Math.round(cemSaat * 100) / 100,
      cemGun: Math.round(cemGun * 100) / 100,
      unikalGun: days.size,
      isGunu,
      hefteSonu,
    });
  }
  out.sort((a, b) => a.ad.localeCompare(b.ad, "az") || a.kod.localeCompare(b.kod, "az"));
  return out;
}

// ────────────────────────────────────────────────────────────────────────────
// Vərəq yazma köməkçiləri
// ────────────────────────────────────────────────────────────────────────────

function cellRef(r1: number, c0: number): string {
  return XLSX.utils.encode_cell({ r: r1 - 1, c: c0 });
}

function setZ(ws: XLSX.WorkSheet, r1: number, c0: number, z: string) {
  const cell = ws[cellRef(r1, c0)];
  if (cell && typeof cell.v === "number") cell.z = z;
}

function setF(ws: XLSX.WorkSheet, r1: number, c0: number, f: string) {
  const cell = ws[cellRef(r1, c0)];
  if (cell) cell.f = f;
}

const DATA_HEADER_BASE = [
  "№",
  "Mənbə",
  "Mənbə sətri",
  "Əvəz edən kod",
  "Əvəz edən (S.A.A.)",
  "Əvəz edən struktur",
  "Əvəz edən vəzifə",
  "Əvəz edilən kod",
  "Əvəz edilən (S.A.A.)",
  "Əvəz edilən struktur",
  "Əvəz edilən vəzifə",
  "Təyin yeri",
  "Baş ofis/Filial",
  "Əvəzetmə növü",
  "Tip",
  "Başlama",
  "Bitmə",
  "Hesablama bitməsi",
  "Gün sayı",
  "Saat sayı",
];

/** 9.1 — ortaq data sxemi ilə vərəq (Təmiz Data / Yekun). */
function writeDataSheet(rows: NormRow[], includeXerc: boolean, preface: string[], params: ReplacementParams): XLSX.WorkSheet {
  const header = includeXerc
    ? [...DATA_HEADER_BASE, "Xərc yaradır?", "Qeydlər"]
    : [...DATA_HEADER_BASE, "Qeydlər"];

  const aoa: Row[] = [...preface.map((p) => [p] as Row), header];
  const headerRow1 = preface.length + 1; // 1-based
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const base: Row = [
      i + 1,
      r.menbe,
      r.menbeSetri,
      r.evezEdenKod,
      r.evezEdenAd,
      r.evezEdenStruktur,
      r.evezEdenVezife,
      r.evezEdilenKod,
      r.evezEdilenAd,
      r.evezEdilenStruktur,
      r.evezEdilenVezife,
      r.teyinYeri,
      r.basOfis,
      r.evezetmeNovu,
      r.dayBased ? "Günlük" : "Saatlıq",
      r.baslama,
      r.bitme,
      r.effEnd ?? effectiveEnd(r, params),
      r.dayCount ?? 0,
      r.hourCount ?? 0,
    ];
    if (includeXerc) base.push(r.xerc ? "Bəli" : "Xeyr");
    base.push(r.qeydler.join(" || "));
    aoa.push(base);
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  for (let i = 0; i < rows.length; i++) {
    const r1 = headerRow1 + 1 + i; // excel row
    const r = rows[i];
    const dz = r.dayBased ? "dd.mm.yyyy" : "dd.mm.yyyy hh:mm";
    setZ(ws, r1, 15, dz); // P Başlama
    setZ(ws, r1, 16, dz); // Q Bitmə
    setZ(ws, r1, 17, dz); // R Hesablama bitməsi
    // SPEC 8 — düsturlar (cache edilmiş dəyərlə birlikdə yazılır)
    setF(ws, r1, 18, `IF(O${r1}="Saatlıq",ROUND((R${r1}-P${r1})*24/8,2),INT(R${r1})-INT(P${r1})+1)`);
    setF(ws, r1, 19, `IF(O${r1}="Saatlıq",ROUND((R${r1}-P${r1})*24,2),(INT(R${r1})-INT(P${r1})+1)*8)`);
  }

  ws["!cols"] = [
    { wch: 6 }, { wch: 26 }, { wch: 12 }, { wch: 12 }, { wch: 28 }, { wch: 28 }, { wch: 28 },
    { wch: 12 }, { wch: 28 }, { wch: 28 }, { wch: 28 }, { wch: 28 }, { wch: 12 }, { wch: 34 },
    { wch: 9 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 9 }, { wch: 9 },
    ...(includeXerc ? [{ wch: 12 }] : []),
    { wch: 60 },
  ];
  return ws;
}

// ────────────────────────────────────────────────────────────────────────────

export function buildWorkbook(input: ReportInput): ReportOutput {
  const { clean, errors, dups, backups, params } = input;

  const yb = buildYekun(clean, errors, params);
  const cem = buildCem(clean, params);

  // Xəta Loqu sıralaması: mənbə + sətir (SPEC 9.4)
  const srcIdx = (m: string) => {
    const i = SOURCE_ORDER.indexOf(m as SourceName);
    return i < 0 ? 99 : i;
  };
  errors.sort((a, b) => {
    const d = srcIdx(a.menbe) - srcIdx(b.menbe);
    if (d) return d;
    const na = parseInt(a.setr, 10) || 0;
    const nb = parseInt(b.setr, 10) || 0;
    return na - nb;
  });

  const wb = XLSX.utils.book_new();

  // ── 9.7 Ümumi Baxış ─────────────────────────────────────────────────────────
  const today = new Date();
  const todayStr = `${String(today.getDate()).padStart(2, "0")}.${String(today.getMonth() + 1).padStart(2, "0")}.${today.getFullYear()}`;
  const umumi: Row[] = [
    ["ƏVƏZETMƏ HESABATI — KONSOLİDASİYA"],
    [`Hazırlanma tarixi: ${todayStr}`],
    [`Filtr: başlama tarixi ≥ ${fmtDate(params.filterStart)}; açıq bitmələr ${fmtDate(params.capDate)}-a qədər hesablanır`],
    [`"Əməkdaş üzrə cəm" pəncərəsi: ${fmtDate(params.winStart)} – ${fmtDate(params.winEnd)} (${params.winEnd - params.winStart + 1} gün)`],
    [],
    ["MƏNBƏ FAYL → HESABATDAKI AD"],
    ["İnzibatciliq prosesi 2.xlsx (Sheet1)", "İnzibatçı prosesi"],
    ["Vacation, Time off replacement.xls (Vacation replacement)", "İşburaxma hallarının əvəzedilməsi"],
    ["Vacation, Time off replacement.xls (time off replacement)", "Saatlıq icazə"],
    ["Filial baza .xlsx (Sheet1)", "Müvəqqəti keçid"],
    ["Rəsmi əvəzetmə(position info).xlsx (Select d_position_info)", "Rəsmi əvəzetmə (həvalə)"],
    ["Vəzifə-Maaş.xlsx", "yalnız lookup — hesabata mənbə kimi girmir"],
    [],
    ["VƏRƏQLƏRİN TƏYİNATI"],
    [YEKUN_SHEET, "Yalnız xərc yaradan əvəzetmələr — əsas tələb olunan vərəq"],
    [TEMIZ_SHEET, "Normalizasiya + duplikat təmizləməsindən sonra bütün sətirlər (Bəli və Xeyr)"],
    [CEM_SHEET, "Hər əvəz edən əməkdaş üzrə cəm göstəricilər (dəyişdirilə bilən pəncərə)"],
    [XETA_SHEET, "Bütün istisna/flaq halları — mənbə sətri ilə audit izlənəbilirliyi"],
    [DUP_SHEET, "Fayl daxili və fayllar arası duplikatlar"],
    ["Statistika", "Yekun göstəricilər, bölgülər və metodologiya"],
    ["BK_* vərəqləri", "Mənbə dataların dəyişdirilməmiş surətləri"],
    [],
    ["XƏRC QAYDALARININ XÜLASƏSİ (istifadəçi tərəfindən təsdiqlənib)"],
    ["İnzibatçı prosesi", "2-ci xəzinə inzibatçısı və '3-cü xəzinə inzibatçısı və Xidmət bölməsinin rəhbəri səlahiyyətlərini icra edən şəxs' → Bəli; 1-ci xəzinə inzibatçısı və tək 3-cü xəzinə inzibatçısı → Xeyr"],
    ["İşburaxma və Saatlıq icazə", "Filial + hər iki tərəfin vəzifəsi tam olaraq 'müdir' və ya 'bölmə rəhbəri' → Xeyr; qalan bütün hallar (Baş ofis daxil) → Bəli"],
    ["Müvəqqəti keçid", "Təyinat mətnində '3-cü xəzinə inzibatçısı və Xidmət bölməsinin rəhbəri' və ya 'filial müdiri vəzifəsini icra edən' keçirsə → Bəli; qalanları → Xeyr"],
    ["Rəsmi əvəzetmə (həvalə)", "Qırmızı sətirlər istisna edildikdən sonra bütün sətirlər → Bəli"],
    ["Duplikat konflikti", "Qrupda hər hansı mənbə 'Bəli' deyirsə, yekun 'Bəli' (konservativ prinsip)"],
  ];
  const umumiWs = XLSX.utils.aoa_to_sheet(umumi);
  umumiWs["!cols"] = [{ wch: 58 }, { wch: 100 }];
  XLSX.utils.book_append_sheet(wb, umumiWs, "Ümumi Baxış");

  // ── 9.2 Xərc yaradan (yekun) ────────────────────────────────────────────────
  const yekunWs = writeDataSheet(
    yb.yekun,
    false,
    [
      "XƏRC YARADAN ƏVƏZETMƏLƏR (YEKUN)",
      `Bu vərəqdə yalnız xərc yaradan ("Bəli") əvəzetmələr göstərilir; duplikatlar, XƏTA (üst-üstə düşmə) qeydli sətirlər və ləğv/dayandırılma əmri olan keçidlər çıxarılıb. İnzibatçı prosesində davamlı sətirlər epizodlara birləşdirilib. Açıq bitmələr ${fmtDate(params.capDate)}-a qədər hesablanıb.`,
      "",
    ],
    params,
  );
  XLSX.utils.book_append_sheet(wb, yekunWs, YEKUN_SHEET);

  // ── 9.1 Təmiz Data ──────────────────────────────────────────────────────────
  const temizWs = writeDataSheet(clean, true, [], params);
  XLSX.utils.book_append_sheet(wb, temizWs, TEMIZ_SHEET);

  // ── 9.3 Əməkdaş üzrə cəm ────────────────────────────────────────────────────
  const N = 1 + clean.length; // Təmiz Data son sətri
  const cemAoa: Row[] = [
    ["ƏMƏKDAŞ ÜZRƏ CƏM"],
    ["Dövr:", params.winStart, "—", params.winEnd],
    [
      `B2 və D2 tarixlərini dəyişməklə C–E sütunları (düstur) yenidən hesablanır. F–H sütunları hazırlanma zamanı ${fmtDate(params.winStart)}–${fmtDate(params.winEnd)} pəncərəsi üçün hesablanıb: unikal təqvim günü sətir aralıqlarının birləşməsidir (maksimum ${params.winEnd - params.winStart + 1} gün).`,
    ],
    [],
    [
      "Əvəz edən kod",
      "Ad",
      "Əvəzetmə sayı",
      "Cəmi saat",
      "Sətirlərin cəm günü",
      "Unikal təqvim günü",
      "— iş günü",
      "— şənbə/bazar",
      "Kəsişmə var?",
    ],
    ...cem.map((c) => [
      c.kod,
      c.ad,
      c.say,
      c.cemSaat,
      c.cemGun,
      c.unikalGun,
      c.isGunu,
      c.hefteSonu,
      c.cemGun > c.unikalGun ? "Bəli" : "",
    ] as Row),
  ];
  const cemWs = XLSX.utils.aoa_to_sheet(cemAoa);
  setZ(cemWs, 2, 1, "dd.mm.yyyy");
  setZ(cemWs, 2, 3, "dd.mm.yyyy");
  for (let i = 0; i < cem.length; i++) {
    const r1 = 6 + i;
    const td = `'${TEMIZ_SHEET}'`;
    setF(cemWs, r1, 2, `COUNTIFS(${td}!$D$2:$D$${N},$A${r1},${td}!$P$2:$P$${N},">="&$B$2,${td}!$P$2:$P$${N},"<"&$D$2+1)`);
    setF(cemWs, r1, 3, `SUMIFS(${td}!$T$2:$T$${N},${td}!$D$2:$D$${N},$A${r1},${td}!$P$2:$P$${N},">="&$B$2,${td}!$P$2:$P$${N},"<"&$D$2+1)`);
    setF(cemWs, r1, 4, `SUMIFS(${td}!$S$2:$S$${N},${td}!$D$2:$D$${N},$A${r1},${td}!$P$2:$P$${N},">="&$B$2,${td}!$P$2:$P$${N},"<"&$D$2+1)`);
    setF(cemWs, r1, 8, `IF(E${r1}>F${r1},"Bəli","")`);
  }
  cemWs["!cols"] = [
    { wch: 14 }, { wch: 32 }, { wch: 13 }, { wch: 11 }, { wch: 16 }, { wch: 16 }, { wch: 11 }, { wch: 14 }, { wch: 12 },
  ];
  XLSX.utils.book_append_sheet(wb, cemWs, CEM_SHEET);

  // ── 9.4 Xəta Loqu ───────────────────────────────────────────────────────────
  const xetaAoa: Row[] = [
    ["№", "Mənbə", "Mənbə sətri (Excel)", "Xəta növü", "Təfərrüat", "Sətir məlumatı"],
    ...errors.map((e, i) => [i + 1, e.menbe, e.setr, e.novu, e.teferruat, e.setirMelumati] as Row),
  ];
  const xetaWs = XLSX.utils.aoa_to_sheet(xetaAoa);
  xetaWs["!cols"] = [{ wch: 6 }, { wch: 26 }, { wch: 14 }, { wch: 34 }, { wch: 70 }, { wch: 100 }];
  XLSX.utils.book_append_sheet(wb, xetaWs, XETA_SHEET);

  // ── 9.5 Duplikatlar ─────────────────────────────────────────────────────────
  const intraCount = dups.filter((d) => d.novu === "Fayl daxili tam duplikat").length;
  const crossCount = dups.length - intraCount;
  const dupEnd = 7 + dups.length;
  const dupAoa: Row[] = [
    ["DUPLİKATLAR"],
    ["Fayl daxili tam duplikat:", intraCount],
    ["Fayllar arası duplikat:", crossCount],
    ["Cəmi:", dups.length],
    [],
    [],
    ["Duplikat növü", "Əvəz edən kod", "Əvəz edən", "Başlama", "Bitmə", "Saxlanılan mənbə (təmiz datada)", "Çıxarılan mənbə", "Qeyd"],
    ...dups.map((d) => [d.novu, d.kod, d.ad, d.baslama, d.bitme, d.saxlanilanMenbe, d.cixarilanMenbe, d.qeyd] as Row),
  ];
  const dupWs = XLSX.utils.aoa_to_sheet(dupAoa);
  if (dups.length) {
    setF(dupWs, 2, 1, `COUNTIF($A$8:$A$${dupEnd},"Fayl daxili tam duplikat")`);
    setF(dupWs, 3, 1, `COUNTIF($A$8:$A$${dupEnd},"Fayllar arası duplikat")`);
    setF(dupWs, 4, 1, `COUNTA($A$8:$A$${dupEnd})`);
    for (let i = 0; i < dups.length; i++) {
      setZ(dupWs, 8 + i, 3, "dd.mm.yyyy");
      setZ(dupWs, 8 + i, 4, "dd.mm.yyyy");
    }
  }
  dupWs["!cols"] = [{ wch: 24 }, { wch: 14 }, { wch: 30 }, { wch: 12 }, { wch: 12 }, { wch: 36 }, { wch: 36 }, { wch: 70 }];
  XLSX.utils.book_append_sheet(wb, dupWs, DUP_SHEET);

  // ── 9.6 Statistika ──────────────────────────────────────────────────────────
  const yekunEnd = 4 + yb.yekun.length;
  const cemEnd = 5 + cem.length;
  const xetaEnd = 1 + errors.length;
  const yk = `'${YEKUN_SHEET}'`;
  const td = `'${TEMIZ_SHEET}'`;
  const xl = `'${XETA_SHEET}'`;

  const yekunGun = yb.yekun.filter((r) => r.dayBased).reduce((s, r) => s + (r.dayCount ?? 0), 0);
  const yekunSaat =
    Math.round(yb.yekun.filter((r) => !r.dayBased).reduce((s, r) => s + (r.hourCount ?? 0), 0) * 100) / 100;
  const beliCount = clean.filter((r) => r.xerc).length;

  const statCards: StatCard[] = [
    { label: "Yekun xərc sətri", value: yb.yekun.length },
    { label: "Yekun cəmi gün", value: yekunGun },
    { label: "Yekun cəmi saat (saatlıq)", value: yekunSaat },
    { label: "Təmiz sətir (cəmi)", value: clean.length },
    { label: "Unikal əməkdaş", value: cem.length },
  ];

  interface StatRow {
    a?: Cell;
    b?: Cell;
    bf?: string;
    c?: Cell;
    cf?: string;
  }
  const stat: StatRow[] = [
    { a: "STATİSTİKA" },
    {},
    { a: "Yekun xərc sətri", b: yb.yekun.length, bf: yb.yekun.length ? `COUNTA(${yk}!$A$5:$A$${yekunEnd})` : undefined },
    { a: "Yekun cəmi gün", b: yekunGun, bf: yb.yekun.length ? `SUMIF(${yk}!$O$5:$O$${yekunEnd},"Günlük",${yk}!$S$5:$S$${yekunEnd})` : undefined },
    { a: "Yekun cəmi saat (saatlıq)", b: yekunSaat, bf: yb.yekun.length ? `SUMIF(${yk}!$O$5:$O$${yekunEnd},"Saatlıq",${yk}!$T$5:$T$${yekunEnd})` : undefined },
    { a: "Təmiz sətir (cəmi)", b: clean.length, bf: clean.length ? `COUNTA(${td}!$A$2:$A$${N})` : undefined },
    { a: "Unikal əməkdaş", b: cem.length, bf: cem.length ? `COUNTA('${CEM_SHEET}'!$A$6:$A$${cemEnd})` : undefined },
    {},
    { a: "YEKUN VƏRƏQDƏN İSTİSNALAR" },
    { a: "XƏTA (üst-üstə düşmə) qeydli Bəli sətirləri", b: yb.excludedXeta },
    { a: "Ləğv/dayandırılma əmri qeydli keçidlər (Bəli)", b: yb.excludedLegv },
    {
      a: "Kəsişən epizodlar",
      b: yb.excludedEpizod,
      bf: errors.length ? `COUNTIF(${xl}!$D$2:$D$${xetaEnd},"Yekun vərəqdən istisna: kəsişən epizodlar")` : undefined,
    },
    { a: "Birləşdirilən İnzibatçı mənbə sətri", b: yb.mergedSourceRows },
    {},
    { a: "TƏMİZ DATA XƏRC BÖLGÜSÜ" },
    { a: "Xərc yaradan (Bəli)", b: beliCount, bf: clean.length ? `COUNTIF(${td}!$U$2:$U$${N},"Bəli")` : undefined },
    { a: "Xərc yaratmayan (Xeyr)", b: clean.length - beliCount, bf: clean.length ? `COUNTIF(${td}!$U$2:$U$${N},"Xeyr")` : undefined },
    {},
    { a: "MƏNBƏ ÜZRƏ (sətir sayı | xərc yaradan)" },
    ...SOURCE_ORDER.map((s) => {
      const cnt = clean.filter((r) => r.menbe === s).length;
      const yes = clean.filter((r) => r.menbe === s && r.xerc).length;
      return {
        a: s,
        b: cnt,
        bf: clean.length ? `COUNTIF(${td}!$B$2:$B$${N},"${s}")` : undefined,
        c: yes,
        cf: clean.length ? `COUNTIFS(${td}!$B$2:$B$${N},"${s}",${td}!$U$2:$U$${N},"Bəli")` : undefined,
      } as StatRow;
    }),
    {},
    { a: "BAŞ OFİS / FİLİAL (Təmiz Data)" },
    {
      a: "Baş ofis",
      b: clean.filter((r) => r.basOfis === "Baş ofis").length,
      bf: clean.length ? `COUNTIF(${td}!$M$2:$M$${N},"Baş ofis")` : undefined,
    },
    {
      a: "Filial",
      b: clean.filter((r) => r.basOfis === "Filial").length,
      bf: clean.length ? `COUNTIF(${td}!$M$2:$M$${N},"Filial")` : undefined,
    },
    {},
    { a: "DUPLİKATLAR" },
    { a: "Fayl daxili tam duplikat", b: intraCount, bf: dups.length ? `COUNTIF('${DUP_SHEET}'!$A$8:$A$${dupEnd},"Fayl daxili tam duplikat")` : undefined },
    { a: "Fayllar arası duplikat", b: crossCount, bf: dups.length ? `COUNTIF('${DUP_SHEET}'!$A$8:$A$${dupEnd},"Fayllar arası duplikat")` : undefined },
    {},
    { a: "XƏTA LOQU KATEQORİYALARI" },
    ...Array.from(new Set(errors.map((e) => e.novu)))
      .sort((x, y) => x.localeCompare(y, "az"))
      .map((t) => ({
        a: t,
        b: errors.filter((e) => e.novu === t).length,
        bf: `COUNTIF(${xl}!$D$2:$D$${xetaEnd},"${t.replace(/"/g, '""')}")`,
      })),
    {},
    { a: "METODOLOGİYA" },
    { a: `1. Beş mənbə fayl vahid sxemə normallaşdırılır; başlama tarixi ${fmtDate(params.filterStart)}-dən əvvəl olan sətirlər Xəta Loquna düşür.` },
    { a: `2. Açıq (boş və ya 31.12.2999) bitmələr ${fmtDate(params.capDate)} sərhədinə qədər hesablanır; mənfi gün sayı yarana bilməz.` },
    { a: "3. Xərc qaydaları mənbə üzrə tətbiq olunur: İnzibatçıda Növ dəyərinə, İşburaxma/Saatlıqda Filial + hər iki tərəfin 'müdir'/'bölmə rəhbəri' vəzifəsinə, Müvəqqəti keçiddə təyinat mətninə görə; Rəsmi əvəzetmədə qırmızılar çıxarıldıqdan sonra hamısı Bəli." },
    { a: "4. Fayl daxili tam duplikatlar (mənbə+kod+başlama+bitmə) və fayllar arası duplikatlar (kod+tarix səviyyəsində başlama/bitmə) çıxarılır; fayllar arası saxlama prioriteti: İnzibatçı → İşburaxma → Saatlıq → Müvəqqəti keçid → Rəsmi əvəzetmə." },
    { a: "5. Duplikat qrupunda xərc konflikti konservativ həll olunur: hər hansı mənbə 'Bəli' deyirsə, yekun 'Bəli'." },
    { a: "6. Eyni mənbə daxilində eyni şəxsin kəsişən dövrləri XƏTA kimi flaqlanır — Təmiz Datada qalır, Yekun vərəqə düşmür. İşburaxmada iç-içə hallar məlumat xarakterli qeydlə saxlanılır." },
    { a: "7. Günlük sətirlərdə gün sayı hər iki tarix daxil olmaqla təqvim günüdür (saat = gün × 8); saatlıq sətirlərdə saat = bitmə − başlama, gün = saat / 8. Hesablamalar Excel düsturları ilə aparılır." },
    { a: "8. Yekun vərəqdə İnzibatçı prosesinin davamlı (kəsişən/bitişik) sətirləri epizodlara birləşdirilir; birləşmədən sonra hələ də kəsişən epizodlar yekundan çıxarılır." },
    { a: `9. Əməkdaş üzrə cəmdə unikal təqvim günü sətir aralıqlarının birləşməsidir (${fmtDate(params.winStart)}–${fmtDate(params.winEnd)} pəncərəsinə kəsilir); iş günü + şənbə/bazar = unikal gün.` },
  ];

  const statAoa: Row[] = stat.map((s) => [s.a ?? null, s.b ?? null, s.c ?? null]);
  const statWs = XLSX.utils.aoa_to_sheet(statAoa);
  for (let i = 0; i < stat.length; i++) {
    const s = stat[i];
    if (s.bf) setF(statWs, i + 1, 1, s.bf);
    if (s.cf) setF(statWs, i + 1, 2, s.cf);
  }
  statWs["!cols"] = [{ wch: 110 }, { wch: 14 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(wb, statWs, "Statistika");

  // ── 10. Backup vərəqləri ────────────────────────────────────────────────────
  for (const bk of backups) {
    const name = bk.sheetName.slice(0, 31);
    const aoa: Row[] = [[bk.note], ...bk.aoa];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const colCount = Math.max(1, ...bk.aoa.map((r) => (r ? r.length : 0)));
    ws["!cols"] = new Array(Math.min(colCount, 60)).fill({ wch: 18 });
    XLSX.utils.book_append_sheet(wb, ws, name);
  }

  return {
    wb,
    yekun: yb.yekun,
    cem,
    statCards,
    yekunExcludedXeta: yb.excludedXeta,
    yekunExcludedEpizod: yb.excludedEpizod,
    yekunExcludedLegv: yb.excludedLegv,
    mergedSourceRows: yb.mergedSourceRows,
  };
}
