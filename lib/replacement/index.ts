import * as XLSX from "xlsx";
import { DEFAULT_PARAMS, type ReplacementParams, type SourceName } from "./config";
import { normalizeAll, type SourceWorkbooks } from "./normalize";
import { runChecks } from "./checks";
import { buildWorkbook, type StatCard } from "./report";
import { runVerification } from "./verify";
import type { CemRow, CheckResult, DupEntry, ErrorLogEntry, NormRow } from "./types";

export { DEFAULT_PARAMS, serialFromYMD } from "./config";
export type { ReplacementParams } from "./config";
export { fmtDate, fmtDateTime } from "./xl";
export type { CemRow, CheckResult, DupEntry, ErrorLogEntry, NormRow } from "./types";

export type ProgressFn = (step: string, pct: number) => void;

export interface ReplacementFiles {
  /** İnzibatciliq prosesi 2.xlsx */
  inzibatci: File;
  /** Vacation, Time off replacement.xls */
  vacation: File;
  /** Filial baza .xlsx */
  filial: File;
  /** Rəsmi əvəzetmə(position info).xlsx */
  resmi: File;
  /** Vəzifə-Maaş.xlsx */
  vezifeMaas: File;
}

export interface ReplacementResult {
  workbook: XLSX.WorkBook;
  checks: CheckResult[];
  allPassed: boolean;
  clean: NormRow[];
  yekun: NormRow[];
  cem: CemRow[];
  errors: ErrorLogEntry[];
  dups: DupEntry[];
  statCards: StatCard[];
  counters: {
    normalized: number;
    dateLogicRemoved: number;
    dupIntra: number;
    dupCross: number;
    overlapRows: number;
    nestedRows: number;
    costConflicts: number;
    redRows: number;
    beli: number;
    xeyr: number;
    yekunExcludedXeta: number;
    yekunExcludedEpizod: number;
    yekunExcludedLegv: number;
    mergedSourceRows: number;
    maxUnikalGun: number;
    muveqqetiCostCount: number;
    muveqqetiCostUnique: number;
  };
  sourceCounts: Record<SourceName, number>;
  params: ReplacementParams;
  durationMs: number;
}

async function readWb(file: File, opts?: XLSX.ParsingOptions): Promise<XLSX.WorkBook> {
  const buf = await file.arrayBuffer();
  return XLSX.read(buf, { type: "array", cellDates: false, ...opts });
}

export async function processReplacement(
  files: ReplacementFiles,
  params: ReplacementParams = DEFAULT_PARAMS,
  onProgress: ProgressFn = () => {},
): Promise<ReplacementResult> {
  const t0 = typeof performance !== "undefined" ? performance.now() : Date.now();
  const now = () => (typeof performance !== "undefined" ? performance.now() : Date.now());

  onProgress("Mərhələ 0: Excel faylları oxunur...", 5);
  const src: SourceWorkbooks = {
    inzibatci: await readWb(files.inzibatci),
    vacation: await readWb(files.vacation),
    filial: await readWb(files.filial),
    // qırmızı sətir aşkarlanması üçün üslub məlumatı da oxunur (SPEC 3.5)
    resmi: await readWb(files.resmi, { cellStyles: true }),
    vezifeMaas: await readWb(files.vezifeMaas),
  };

  onProgress("Mərhələ 1: Normalizasiya (vahid sxem + filtr + xərc qaydaları)...", 25);
  const norm = normalizeAll(src, params);

  onProgress("Mərhələ 2: Duplikatlar, üst-üstə düşmə və iç-içə yoxlamaları...", 45);
  const chk = runChecks(norm.rows, norm.errors, norm.cancels, params);

  onProgress("Mərhələ 3–4: Yekun workbook və backup vərəqləri qurulur...", 65);
  const rep = buildWorkbook({
    clean: chk.clean,
    errors: norm.errors,
    dups: chk.dups,
    backups: norm.backups,
    counters: chk.counters,
    sourceCounts: norm.sourceCounts,
    redRows: norm.redRows,
    params,
  });

  onProgress("Mərhələ 5: Müstəqil yoxlamalar (bölmə 12)...", 88);
  const checks = runVerification({
    wb: rep.wb,
    clean: chk.clean,
    yekun: rep.yekun,
    cem: rep.cem,
    dups: chk.dups,
    errors: norm.errors,
    counters: chk.counters,
    muveqqetiCostCount: norm.muveqqetiCostCount,
    statCards: rep.statCards,
    params,
  });
  const allPassed = checks.filter((c) => c.hard).every((c) => c.ok);

  onProgress("Tamamlandı", 100);

  const beli = chk.clean.filter((r) => r.xerc).length;
  return {
    workbook: rep.wb,
    checks,
    allPassed,
    clean: chk.clean,
    yekun: rep.yekun,
    cem: rep.cem,
    errors: norm.errors,
    dups: chk.dups,
    statCards: rep.statCards,
    counters: {
      normalized: chk.counters.normalized,
      dateLogicRemoved: chk.counters.dateLogicRemoved,
      dupIntra: chk.counters.dupIntra,
      dupCross: chk.counters.dupCross,
      overlapRows: chk.counters.overlapRows,
      nestedRows: chk.counters.nestedRows,
      costConflicts: chk.counters.costConflicts,
      redRows: norm.redRows.length,
      beli,
      xeyr: chk.clean.length - beli,
      yekunExcludedXeta: rep.yekunExcludedXeta,
      yekunExcludedEpizod: rep.yekunExcludedEpizod,
      yekunExcludedLegv: rep.yekunExcludedLegv,
      mergedSourceRows: rep.mergedSourceRows,
      maxUnikalGun: rep.cem.reduce((m, c) => Math.max(m, c.unikalGun), 0),
      muveqqetiCostCount: norm.muveqqetiCostCount,
      muveqqetiCostUnique: norm.muveqqetiCostUnique,
    },
    sourceCounts: norm.sourceCounts,
    params,
    durationMs: now() - t0,
  };
}

/** Yekun hesabatı .xlsx kimi yükləyir (yalnız bütün məcburi yoxlamalar keçəndə çağırılmalıdır). */
export function exportReplacement(result: ReplacementResult, filename = "Əvəzetmə_Hesabatı.xlsx") {
  XLSX.writeFile(result.workbook, filename, { compression: true });
}
