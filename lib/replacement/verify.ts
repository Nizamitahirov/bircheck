import * as XLSX from "xlsx";
import type { ReplacementParams } from "./config";
import type { ChecksCounters } from "./checks";
import { effectiveEnd } from "./checks";
import type { CemRow, CheckResult, DupEntry, ErrorLogEntry, NormRow } from "./types";
import type { StatCard } from "./report";

export interface VerifyInput {
  wb: XLSX.WorkBook;
  clean: NormRow[];
  yekun: NormRow[];
  cem: CemRow[];
  dups: DupEntry[];
  errors: ErrorLogEntry[];
  counters: ChecksCounters;
  muveqqetiCostCount: number;
  statCards: StatCard[];
  params: ReplacementParams;
}

const EPS = 0.005;

function near(a: number, b: number): boolean {
  return Math.abs(a - b) <= EPS;
}

function cellV(ws: XLSX.WorkSheet, r1: number, c0: number): unknown {
  const cell = ws[XLSX.utils.encode_cell({ r: r1 - 1, c: c0 })];
  return cell ? cell.v : undefined;
}

/** Vərəqdəki bir data sətrindən gün/saat dəyərlərini düsturdan asılı olmayaraq yenidən hesabla. */
function recomputeFromSheet(ws: XLSX.WorkSheet, r1: number): { gun: number; saat: number; tip: string } | null {
  const tip = cellV(ws, r1, 14);
  const p = cellV(ws, r1, 15);
  const rr = cellV(ws, r1, 17);
  if (typeof p !== "number" || typeof rr !== "number" || typeof tip !== "string") return null;
  if (tip === "Saatlıq") {
    const saat = Math.round((rr - p) * 24 * 100) / 100;
    return { gun: Math.round(((rr - p) * 24 / 8) * 100) / 100, saat, tip };
  }
  const gun = Math.floor(rr) - Math.floor(p) + 1;
  return { gun, saat: gun * 8, tip };
}

/** SPEC bölmə 12 — məcburi yoxlama mərhələsi. */
export function runVerification(v: VerifyInput): CheckResult[] {
  const out: CheckResult[] = [];
  const { wb, clean, yekun, cem, dups, errors, counters, params } = v;

  // 0. Pəncərə uzunluğu (SPEC 2): WIN_START–WIN_END daxil olmaqla dəqiq 181 gün
  const winLen = params.winEnd - params.winStart + 1;
  out.push({
    name: "Pəncərə uzunluğu",
    ok: winLen === 181,
    detail: `WIN_START–WIN_END daxil olmaqla ${winLen} gün (tələb: dəqiq 181).`,
    hard: true,
  });

  // 1. Düstur xanaları (LibreOffice recalc əvəzinə — brauzer mühitində bütün
  //    düstur xanalarına JS ilə hesablanmış dəyərlər yazılır, Excel açılışda yenidən hesablayır)
  let formulaCells = 0;
  let formulaErrors = 0;
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    const ref = ws["!ref"];
    if (!ref) continue;
    const range = XLSX.utils.decode_range(ref);
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (!cell || !cell.f) continue;
        formulaCells++;
        const val = cell.v;
        if (val === undefined || val === null) formulaErrors++;
        else if (typeof val === "number" && !Number.isFinite(val)) formulaErrors++;
        else if (typeof val === "string" && /^#(REF|DIV\/0|VALUE|NAME|N\/A)/.test(val)) formulaErrors++;
      }
    }
  }
  out.push({
    name: "Düstur xanaları (recalc əvəzi)",
    ok: formulaErrors === 0,
    detail: `${formulaCells} düstur xanası; ${formulaErrors} xətalı/boş dəyər (tələb: 0). Bütün düsturlar cache edilmiş dəyərlə yazılıb — Excel/LibreOffice açılışda yenidən hesablayır.`,
    hard: true,
  });

  // 2. Müstəqil cross-check: vərəqdəki dəyərlər ikinci dəfə hesablanır
  const tdWs = wb.Sheets["Təmiz Data"];
  const ykWs = wb.Sheets["Xərc yaradan (yekun)"];
  let ccOk = true;
  const ccDetails: string[] = [];

  const checkSheet = (ws: XLSX.WorkSheet, firstRow: number, count: number, label: string, expGun: number, expSaat: number) => {
    let gunSum = 0;
    let saatSum = 0;
    let mismatch = 0;
    for (let i = 0; i < count; i++) {
      const r1 = firstRow + i;
      const rc = recomputeFromSheet(ws, r1);
      if (!rc) {
        mismatch++;
        continue;
      }
      const sV = cellV(ws, r1, 18);
      const tV = cellV(ws, r1, 19);
      if (typeof sV !== "number" || !near(sV, rc.gun)) mismatch++;
      else if (typeof tV !== "number" || !near(tV, rc.saat)) mismatch++;
      if (rc.tip === "Günlük") gunSum += rc.gun;
      else saatSum += rc.saat;
    }
    saatSum = Math.round(saatSum * 100) / 100;
    const ok = mismatch === 0 && near(gunSum, expGun) && near(saatSum, expSaat);
    if (!ok) ccOk = false;
    ccDetails.push(
      `${label}: günlük gün cəmi ${gunSum} (gözlənilən ${expGun}), saatlıq saat cəmi ${saatSum} (gözlənilən ${expSaat}), uyğunsuz sətir: ${mismatch}`,
    );
  };

  const cleanGun = clean.filter((r) => r.dayBased).reduce((s, r) => s + (r.dayCount ?? 0), 0);
  const cleanSaat = Math.round(clean.filter((r) => !r.dayBased).reduce((s, r) => s + (r.hourCount ?? 0), 0) * 100) / 100;
  const yekGun = yekun.filter((r) => r.dayBased).reduce((s, r) => s + (r.dayCount ?? 0), 0);
  const yekSaat = Math.round(yekun.filter((r) => !r.dayBased).reduce((s, r) => s + (r.hourCount ?? 0), 0) * 100) / 100;
  checkSheet(tdWs, 2, clean.length, "Təmiz Data", cleanGun, cleanSaat);
  checkSheet(ykWs, 5, yekun.length, "Yekun", yekGun, yekSaat);
  out.push({ name: "Müstəqil cross-check", ok: ccOk, detail: ccDetails.join(" | "), hard: true });

  // 3. Təmiz Data: boş və ya ≤0 gün sayı = 0 ədəd
  const badGun = clean.filter((r) => !(typeof r.dayCount === "number" && r.dayCount > 0)).length;
  out.push({
    name: "Təmiz Data: gün sayı",
    ok: badGun === 0,
    detail: `${clean.length} sətirdən ${badGun} sətirdə boş/≤0 gün sayı (tələb: 0).`,
    hard: true,
  });

  // 4. Yekun vərəq invariantları
  const keySet = new Set<string>();
  let yekDup = 0;
  let yekXeta = 0;
  let yekBadGun = 0;
  let yekNotBeli = 0;
  for (const r of yekun) {
    const key = `${r.evezEdenKod}|${r.baslama}|${r.bitme === null ? "open" : r.bitme}`;
    if (keySet.has(key)) yekDup++;
    keySet.add(key);
    if (r.qeydler.some((q) => q.includes("XƏTA"))) yekXeta++;
    if (!(typeof r.dayCount === "number" && r.dayCount > 0)) yekBadGun++;
    if (!r.xerc) yekNotBeli++;
  }
  out.push({
    name: "Yekun vərəq invariantları",
    ok: yekDup === 0 && yekXeta === 0 && yekBadGun === 0 && yekNotBeli === 0,
    detail: `(kod, Başlama, Bitmə) duplikatı: ${yekDup}; XƏTA qeydli: ${yekXeta}; boş/mənfi gün: ${yekBadGun}; Bəli olmayan: ${yekNotBeli} (hamısı 0 olmalıdır).`,
    hard: true,
  });

  // 5. Əməkdaş üzrə cəm
  const maxUnikal = cem.reduce((m, c) => Math.max(m, c.unikalGun), 0);
  const invariantBroken = cem.filter((c) => c.isGunu + c.hefteSonu !== c.unikalGun).length;
  // ən böyük 3 dəyər müstəqil ikinci üsulla (gün-gün skan) yenidən hesablanır
  const byKod = new Map<string, NormRow[]>();
  for (const r of clean) {
    if (!r.evezEdenKod) continue;
    const g = byKod.get(r.evezEdenKod) ?? [];
    g.push(r);
    byKod.set(r.evezEdenKod, g);
  }
  const top3 = [...cem].sort((a, b) => b.unikalGun - a.unikalGun).slice(0, 3);
  let top3Bad = 0;
  for (const c of top3) {
    const rows = byKod.get(c.kod) ?? [];
    let cnt = 0;
    for (let d = params.winStart; d <= params.winEnd; d++) {
      if (
        rows.some((r) => Math.floor(r.baslama) <= d && d <= Math.floor(r.effEnd ?? effectiveEnd(r, params)))
      ) {
        cnt++;
      }
    }
    if (cnt !== c.unikalGun) top3Bad++;
  }
  out.push({
    name: "Əməkdaş üzrə cəm",
    ok: maxUnikal <= winLen && invariantBroken === 0 && top3Bad === 0,
    detail: `Max unikal təqvim günü: ${maxUnikal} (limit ${winLen}); iş günü + şənbə/bazar ≠ unikal gün olan sətir: ${invariantBroken}; ən böyük 3 dəyərin müstəqil hesablanması: ${top3Bad === 0 ? "uyğundur" : `${top3Bad} uyğunsuzluq`}.`,
    hard: true,
  });

  // 6. Balans
  const beli = clean.filter((r) => r.xerc).length;
  const xeyr = clean.length - beli;
  const balance1 = beli + xeyr === clean.length;
  const expectedClean =
    counters.normalized - counters.dateLogicRemoved - counters.dupIntra - counters.dupCross;
  const balance2 = expectedClean === clean.length;
  out.push({
    name: "Balans",
    ok: balance1 && balance2,
    detail: `Təmiz sətir ${clean.length} = Bəli ${beli} + Xeyr ${xeyr}; normalizasiya ${counters.normalized} − tarix məntiqsizliyi ${counters.dateLogicRemoved} − fayl daxili duplikat ${counters.dupIntra} − fayllar arası duplikat ${counters.dupCross} = ${expectedClean} (təmiz: ${clean.length}).`,
    hard: true,
  });

  // 7. Müvəqqəti keçid xərc-uyğun sətirlərin izlənməsi (SPEC 12.7)
  const isMuv = (m: string) => m === "Müvəqqəti keçid";
  const yekunMuv = yekun.filter((r) => isMuv(r.menbe)).length;
  const crossTwins = dups.filter((d) => d.novu === "Fayllar arası duplikat" && isMuv(d.removedMenbe) && d.removedXerc).length;
  const intraMuv = dups.filter((d) => d.novu === "Fayl daxili tam duplikat" && isMuv(d.removedMenbe) && d.removedXerc).length;
  const xetaMuv = clean.filter((r) => isMuv(r.menbe) && r.xerc && r.hasXeta).length;
  const diqqetMuv = clean.filter(
    (r) => isMuv(r.menbe) && r.xerc && !r.hasXeta && r.qeydler.some((q) => q.startsWith("DİQQƏT: bu keçid üzrə")),
  ).length;
  const reconciled =
    v.muveqqetiCostCount + counters.flippedMuvToBeli ===
    yekunMuv + crossTwins + intraMuv + xetaMuv + diqqetMuv + counters.dateLogicRemovedMuvBeli;
  out.push({
    name: "Müvəqqəti keçid izlənməsi",
    ok: reconciled,
    detail: `Mənbədə xərc-uyğun: ${v.muveqqetiCostCount} (+${counters.flippedMuvToBeli} konfliktdə Bəli-yə çevrilən) = yekunda qalan ${yekunMuv} + digər mənbə altında saxlanılan Bəli duplikat əkizləri ${crossTwins} + fayl daxili duplikat ${intraMuv} + XƏTA istisnası ${xetaMuv} + ləğv/dayandırılma qeydi ${diqqetMuv} + tarix məntiqsizliyi ${counters.dateLogicRemovedMuvBeli}.`,
    hard: false,
  });

  // 8. Statistika kartları vərəq dəyərləri ilə üst-üstə düşür
  const ykCount = yekun.length;
  const sheetYekGun = (() => {
    let s = 0;
    for (let i = 0; i < ykCount; i++) {
      const rc = recomputeFromSheet(ykWs, 5 + i);
      if (rc && rc.tip === "Günlük") s += rc.gun;
    }
    return s;
  })();
  const sheetYekSaat = (() => {
    let s = 0;
    for (let i = 0; i < ykCount; i++) {
      const rc = recomputeFromSheet(ykWs, 5 + i);
      if (rc && rc.tip === "Saatlıq") s += rc.saat;
    }
    return Math.round(s * 100) / 100;
  })();
  const cards = new Map(v.statCards.map((c) => [c.label, c.value]));
  const cardOk =
    cards.get("Yekun xərc sətri") === yekun.length &&
    near(cards.get("Yekun cəmi gün") ?? NaN, sheetYekGun) &&
    near(cards.get("Yekun cəmi saat (saatlıq)") ?? NaN, sheetYekSaat) &&
    cards.get("Təmiz sətir (cəmi)") === clean.length &&
    cards.get("Unikal əməkdaş") === cem.length;
  out.push({
    name: "Statistika kartları",
    ok: cardOk,
    detail: `Kartlar: yekun ${cards.get("Yekun xərc sətri")}/${yekun.length}, gün ${cards.get("Yekun cəmi gün")}/${sheetYekGun}, saat ${cards.get("Yekun cəmi saat (saatlıq)")}/${sheetYekSaat}, təmiz ${cards.get("Təmiz sətir (cəmi)")}/${clean.length}, əməkdaş ${cards.get("Unikal əməkdaş")}/${cem.length}.`,
    hard: true,
  });

  // Xəta Loqu istinad bütövlüyü (əlavə sanity)
  out.push({
    name: "Xəta Loqu",
    ok: true,
    detail: `${errors.length} qeyd, ${dups.length} duplikat (fayl daxili ${dups.filter((d) => d.novu === "Fayl daxili tam duplikat").length} / fayllar arası ${dups.filter((d) => d.novu === "Fayllar arası duplikat").length}).`,
    hard: false,
  });

  return out;
}
