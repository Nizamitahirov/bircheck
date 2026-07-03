import { SOURCE_PRIORITY, type ReplacementParams, type SourceName } from "./config";
import type { CancelInfo, DupEntry, ErrorLogEntry, NormRow } from "./types";
import { fmtDate, fmtDateTime } from "./xl";

export interface ChecksCounters {
  normalized: number;
  dateLogicRemoved: number;
  dupIntra: number;
  dupCross: number;
  overlapRows: number;
  nestedRows: number;
  costConflicts: number;
  /** 12.7 balansı üçün: tarix məntiqsizliyi ilə çıxarılan Müvəqqəti keçid Bəli sətirləri */
  dateLogicRemovedMuvBeli: number;
  /** 12.7 balansı üçün: konflikt nəticəsində Xeyr→Bəli çevrilən Müvəqqəti keçid sətirləri */
  flippedMuvToBeli: number;
}

export interface ChecksResult {
  clean: NormRow[];
  dups: DupEntry[];
  counters: ChecksCounters;
}

function sourceTag(r: NormRow): string {
  return `${r.menbe} (sətir ${r.menbeSetri})`;
}

/** Effektiv (hesablama) bitməsi: real bitmə; boşdursa max(CAP_DATE, başlama günü). */
export function effectiveEnd(r: NormRow, params: ReplacementParams): number {
  return r.bitme !== null ? r.bitme : Math.max(params.capDate, Math.floor(r.baslama));
}

function periodText(r: NormRow, params: ReplacementParams): string {
  const e = effectiveEnd(r, params);
  return r.dayBased
    ? `${fmtDate(r.baslama)}–${r.bitme === null ? "açıq" : fmtDate(e)}`
    : `${fmtDateTime(r.baslama)}–${fmtDateTime(e)}`;
}

export function runChecks(
  input: NormRow[],
  errors: ErrorLogEntry[],
  cancels: CancelInfo[],
  params: ReplacementParams,
): ChecksResult {
  const counters: ChecksCounters = {
    normalized: input.length,
    dateLogicRemoved: 0,
    dupIntra: 0,
    dupCross: 0,
    overlapRows: 0,
    nestedRows: 0,
    costConflicts: 0,
    dateLogicRemovedMuvBeli: 0,
    flippedMuvToBeli: 0,
  };
  const dups: DupEntry[] = [];

  // ── 7.3 Tarix məntiqsizliyi (günlük sətirdə real bitmə < başlama) ──────────
  let rows: NormRow[] = [];
  for (const r of input) {
    if (r.dayBased && r.bitme !== null && r.bitme < r.baslama) {
      counters.dateLogicRemoved++;
      if (r.menbe === "Müvəqqəti keçid" && r.xerc) counters.dateLogicRemovedMuvBeli++;
      errors.push({
        menbe: r.menbe,
        setr: r.menbeSetri,
        novu: "Tarix məntiqsizliyi (çıxarıldı)",
        teferruat: `Bitmə (${fmtDate(r.bitme)}) başlamadan (${fmtDate(r.baslama)}) əvvəldir — sətir təmiz dataya salınmadı.`,
        setirMelumati: `Əvəz edən: ${r.evezEdenAd} (${r.evezEdenKod}) | Növ: ${r.evezetmeNovu}`,
      });
      continue;
    }
    rows.push(r);
  }

  // ── 6.1 Fayl daxili tam duplikat ───────────────────────────────────────────
  const seenExact = new Map<string, NormRow>();
  const afterIntra: NormRow[] = [];
  for (const r of rows) {
    const key = `${r.menbe}|${r.evezEdenKod}|${r.baslama}|${r.bitme === null ? "open" : r.bitme}`;
    const first = seenExact.get(key);
    if (!first) {
      seenExact.set(key, r);
      afterIntra.push(r);
      continue;
    }
    counters.dupIntra++;
    dups.push({
      novu: "Fayl daxili tam duplikat",
      kod: r.evezEdenKod,
      ad: r.evezEdenAd,
      baslama: r.baslama,
      bitme: r.bitme,
      saxlanilanMenbe: sourceTag(first),
      cixarilanMenbe: sourceTag(r),
      qeyd: "Eyni fayl daxilində kod + başlama + bitmə tam üst-üstə düşür; birinci sətir saxlanılıb.",
      removedMenbe: r.menbe,
      removedXerc: r.xerc,
      keptMenbe: first.menbe,
    });
  }
  rows = afterIntra;

  // ── 6.2 Fayllar arası duplikat (tarix səviyyəsində) ────────────────────────
  const groups = new Map<string, NormRow[]>();
  for (const r of rows) {
    const key = `${r.evezEdenKod}|${Math.floor(r.baslama)}|${r.bitme === null ? "open" : Math.floor(r.bitme)}`;
    const g = groups.get(key);
    if (g) g.push(r);
    else groups.set(key, [r]);
  }

  const removedCross = new Set<NormRow>();
  for (const g of groups.values()) {
    const distinctSources = new Set(g.map((r) => r.menbe));
    if (distinctSources.size < 2) continue;

    const sorted = [...g].sort(
      (a, b) => SOURCE_PRIORITY[a.menbe] - SOURCE_PRIORITY[b.menbe] || a.rowNums[0] - b.rowNums[0],
    );
    const kept = sorted[0];
    const removed = sorted.slice(1).filter((r) => r.menbe !== kept.menbe);
    if (!removed.length) continue;

    kept.qeydler.push(
      `Fayllar arası duplikat: eyni əvəzetmə ${removed
        .map((r) => sourceTag(r))
        .join(", ")} mənbələrində də mövcuddur (Request sisteminin fərqli tool-larından yaradılıb); yalnız bu sətir saxlanılıb`,
    );

    // 5.1 — xərc konflikti: qrupda hər hansı mənbə "Bəli" deyirsə, yekun "Bəli"
    const votes = [kept, ...removed];
    const anyYes = votes.some((r) => r.xerc);
    const anyNo = votes.some((r) => !r.xerc);
    if (anyYes && anyNo) {
      counters.costConflicts++;
      const detail = votes.map((r) => `${sourceTag(r)}: ${r.xerc ? "Bəli" : "Xeyr"}`).join("; ");
      if (!kept.xerc) {
        kept.xerc = true;
        if (kept.menbe === "Müvəqqəti keçid") counters.flippedMuvToBeli++;
      }
      kept.qeydler.push(
        `Xərc qiymətləndirməsi konflikti: duplikat qrupunda mənbələr fərqli dəyər verir (${detail}); konservativ prinsiplə "Bəli" götürülüb`,
      );
      errors.push({
        menbe: kept.menbe,
        setr: kept.menbeSetri,
        novu: "Xərc qiymətləndirməsi konflikti",
        teferruat: `Duplikat qrupunda mənbələr fərqli xərc dəyəri verir: ${detail}. Konservativ prinsiplə "Bəli" götürülüb.`,
        setirMelumati: `Əvəz edən: ${kept.evezEdenAd} (${kept.evezEdenKod}) | Dövr: ${periodText(kept, params)}`,
      });
    }

    for (const r of removed) {
      removedCross.add(r);
      counters.dupCross++;
      dups.push({
        novu: "Fayllar arası duplikat",
        kod: r.evezEdenKod,
        ad: r.evezEdenAd,
        baslama: r.baslama,
        bitme: r.bitme,
        saxlanilanMenbe: sourceTag(kept),
        cixarilanMenbe: sourceTag(r),
        qeyd: "Eyni sorğu Request sisteminin fərqli tool-larından yaradılıb; prioritet üzrə saxlanılıb.",
        removedMenbe: r.menbe,
        removedXerc: r.xerc,
        keptMenbe: kept.menbe,
      });
    }
  }
  rows = rows.filter((r) => !removedCross.has(r));

  // ── 3.4 Ləğv/dayandırma əmrləri üçün DİQQƏT qeydləri ───────────────────────
  for (const c of cancels) {
    if (!c.kod) continue;
    for (const r of rows) {
      if (r.menbe !== "Müvəqqəti keçid" || r.evezEdenKod !== c.kod) continue;
      if (c.baslama !== null) {
        const cs = Math.floor(c.baslama);
        const ce = Math.floor(c.bitme ?? c.baslama);
        const rs = Math.floor(r.baslama);
        const re = Math.floor(effectiveEnd(r, params));
        if (!(rs <= ce && cs <= re)) continue;
      }
      r.qeydler.push(
        `DİQQƏT: bu keçid üzrə '${c.nov}' əmri mövcuddur (mənbə sətri ${c.rowNum}${c.emr ? `, əmr ${c.emr}` : ""}) – faktiki müddət fərqli ola bilər`,
      );
    }
  }

  // ── 7.1 Üst-üstə düşən əvəzetmə (eyni mənbə, eyni kod, günlük) ─────────────
  const byMenbeKod = new Map<string, NormRow[]>();
  for (const r of rows) {
    if (!r.dayBased) continue;
    const key = `${r.menbe}|${r.evezEdenKod}`;
    const g = byMenbeKod.get(key);
    if (g) g.push(r);
    else byMenbeKod.set(key, [r]);
  }
  const overlapNotes = new Map<NormRow, string[]>();
  for (const g of byMenbeKod.values()) {
    if (g.length < 2) continue;
    for (let i = 0; i < g.length; i++) {
      for (let j = i + 1; j < g.length; j++) {
        const a = g[i];
        const b = g[j];
        const s1 = Math.floor(a.baslama);
        const e1 = Math.floor(effectiveEnd(a, params));
        const s2 = Math.floor(b.baslama);
        const e2 = Math.floor(effectiveEnd(b, params));
        if (s1 <= e2 && s2 <= e1) {
          const la = overlapNotes.get(a) ?? [];
          la.push(`sətir ${b.menbeSetri} (${periodText(b, params)})`);
          overlapNotes.set(a, la);
          const lb = overlapNotes.get(b) ?? [];
          lb.push(`sətir ${a.menbeSetri} (${periodText(a, params)})`);
          overlapNotes.set(b, lb);
        }
      }
    }
  }
  for (const [r, others] of overlapNotes) {
    counters.overlapRows++;
    r.hasXeta = true;
    r.qeydler.push(
      `XƏTA – üst-üstə düşmə: ${others.join("; ")} ilə kəsişir. Bir şəxs eyni dövrdə iki əvəzetmədə ola bilməz.`,
    );
    errors.push({
      menbe: r.menbe,
      setr: r.menbeSetri,
      novu: "Üst-üstə düşən əvəzetmə",
      teferruat: `Bir şəxs eyni dövrdə iki əvəzetmədə ola bilməz. ${periodText(r, params)} dövrü ${others.join("; ")} ilə kəsişir.`,
      setirMelumati: `Əvəz edən: ${r.evezEdenAd} (${r.evezEdenKod}) | Növ: ${r.evezetmeNovu}`,
    });
  }

  // ── 7.2 İç-içə əvəzetmə (yalnız İşburaxma) ─────────────────────────────────
  const isb = rows.filter((r) => r.menbe === "İşburaxma hallarının əvəzedilməsi");
  const byReplaced = new Map<string, NormRow[]>();
  for (const r of isb) {
    if (!r.evezEdilenKod) continue;
    const g = byReplaced.get(r.evezEdilenKod);
    if (g) g.push(r);
    else byReplaced.set(r.evezEdilenKod, [r]);
  }
  const nestedFlagged = new Set<NormRow>();
  for (const a of isb) {
    if (!a.evezEdenKod) continue;
    const matches = byReplaced.get(a.evezEdenKod);
    if (!matches) continue;
    const s1 = Math.floor(a.baslama);
    const e1 = Math.floor(effectiveEnd(a, params));
    for (const b of matches) {
      if (b === a) continue;
      const s2 = Math.floor(b.baslama);
      const e2 = Math.floor(effectiveEnd(b, params));
      if (!(s1 <= e2 && s2 <= e1)) continue;

      a.qeydler.push(
        `İÇ-İÇƏ HAL: əvəz edən şəxs özü ${fmtDate(s2)}–${fmtDate(e2)} dövründə ${b.evezEdenAd} (${b.evezEdenKod}) tərəfindən əvəz edilib (sətir ${b.menbeSetri}); faktiki əvəzetmə müddəti sənəddəki tam müddətdən azdır`,
      );
      b.qeydler.push(
        `İÇ-İÇƏ HAL: bu sətirdə əvəz edilən ${a.evezEdenAd} (${a.evezEdenKod}) özü ${fmtDate(s1)}–${fmtDate(e1)} dövründə əvəz edən kimi çıxış edir (sətir ${a.menbeSetri})`,
      );
      for (const x of [a, b]) {
        if (nestedFlagged.has(x)) continue;
        nestedFlagged.add(x);
        errors.push({
          menbe: x.menbe,
          setr: x.menbeSetri,
          novu: "İç-içə əvəzetmə",
          teferruat:
            x === a
              ? `Əvəz edən şəxs (${a.evezEdenKod}) eyni dövrdə özü əvəz edilib — faktiki əvəzetmə müddəti sənəddəki tam müddətdən azdır.`
              : `Bu sətirdə əvəz edilən şəxs (${a.evezEdenKod}) kəsişən dövrdə başqa sətirdə əvəz edən kimi çıxış edir.`,
          setirMelumati: `Əvəz edən: ${x.evezEdenAd} (${x.evezEdenKod}) | Dövr: ${periodText(x, params)}`,
        });
      }
    }
  }
  counters.nestedRows = nestedFlagged.size;

  // ── 7.4 + bölmə 8: hesablama bitməsi və gün/saat dəyərləri ─────────────────
  for (const r of rows) {
    if (r.bitme === null && Math.floor(r.baslama) > params.capDate) {
      r.qeydler.push(
        `Başlama tarixi hesablama sərhədindən (${fmtDate(params.capDate)}) sonradır və bitmə açıqdır – gün sayı minimum 1 gün kimi hesablanıb`,
      );
    }
    const e = effectiveEnd(r, params);
    r.effEnd = e;
    if (r.dayBased) {
      r.dayCount = Math.floor(e) - Math.floor(r.baslama) + 1;
      r.hourCount = r.dayCount * 8;
    } else {
      r.hourCount = Math.round((e - r.baslama) * 24 * 100) / 100;
      r.dayCount = Math.round(((e - r.baslama) * 24 / 8) * 100) / 100;
    }
  }

  // Təmiz Data sırası: mənbə prioriteti, sonra sətir nömrəsi
  rows.sort(
    (a, b) => SOURCE_PRIORITY[a.menbe] - SOURCE_PRIORITY[b.menbe] || a.rowNums[0] - b.rowNums[0],
  );

  return { clean: rows, dups, counters };
}
