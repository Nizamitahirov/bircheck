// SPEC.md bölmə 2 — konfiqurasiya sabitləri.
// Tarixlər Excel serial nömrələri kimi saxlanılır (1900 sistemi).

export const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
export const MS_PER_DAY = 86_400_000;

export function serialFromYMD(y: number, m: number, d: number): number {
  return Math.round((Date.UTC(y, m - 1, d) - EXCEL_EPOCH_MS) / MS_PER_DAY);
}

export const WORK_HOURS_PER_DAY = 8;

export interface ReplacementParams {
  /** Yalnız başlama tarixi bu tarixdən (daxil) sonrakı sətirlər təmiz dataya düşür. */
  filterStart: number;
  /** Açıq (davam edən) sətirlərdə hesablama sərhədi. */
  capDate: number;
  /** "Əməkdaş üzrə cəm" pəncərəsinin başlanğıcı. */
  winStart: number;
  /** Pəncərənin sonu. WIN_START–WIN_END daxil olmaqla 181 gün olmalıdır. */
  winEnd: number;
}

export const DEFAULT_PARAMS: ReplacementParams = {
  filterStart: serialFromYMD(2026, 1, 1),
  capDate: serialFromYMD(2026, 6, 30),
  winStart: serialFromYMD(2026, 1, 3),
  winEnd: serialFromYMD(2026, 7, 2),
};

export type SourceName =
  | "İnzibatçı prosesi"
  | "İşburaxma hallarının əvəzedilməsi"
  | "Saatlıq icazə"
  | "Müvəqqəti keçid"
  | "Rəsmi əvəzetmə (həvalə)";

/** SPEC 6.2 — fayllar arası duplikatda saxlama prioriteti. */
export const SOURCE_PRIORITY: Record<SourceName, number> = {
  "İnzibatçı prosesi": 0,
  "İşburaxma hallarının əvəzedilməsi": 1,
  "Saatlıq icazə": 2,
  "Müvəqqəti keçid": 3,
  "Rəsmi əvəzetmə (həvalə)": 4,
};

export const SOURCE_ORDER: SourceName[] = [
  "İnzibatçı prosesi",
  "İşburaxma hallarının əvəzedilməsi",
  "Saatlıq icazə",
  "Müvəqqəti keçid",
  "Rəsmi əvəzetmə (həvalə)",
];

/** SPEC 5.1 — İnzibatçı prosesində xərc yaradan Növ dəyərləri. */
export const INZIBATCI_COST_NOV = [
  "2-ci xəzinə inzibatçısı",
  "3-cü xəzinə inzibatçısı və xidmət bölməsinin rəhbəri səlahiyyətlərini icra edən şəxs",
];

/** SPEC 5.1 — İşburaxma/Saatlıq: filialda hər iki tərəfin vəzifəsi tam bunlardan biridirsə xərc yaratmır. */
export const NO_COST_POSITIONS = ["müdir", "bölmə rəhbəri"];

/** SPEC 5.1 — Müvəqqəti keçiddə xərc yaradan təyinat mətnləri (case-insensitive contains). */
export const MUVEQQETI_COST_TEXTS = [
  "3-cü xəzinə inzibatçısı və xidmət bölməsinin rəhbəri",
  "filial müdiri vəzifəsini icra edən",
];

/** SPEC 5.2 — struktur adında bu sözlər keçirsə Baş ofis sayılır. */
export const BO_KEYWORDS = [
  "departament",
  "sahəsi",
  "mərkəzi",
  "mühasibatlıq",
  "ofis",
  "rəhbərlik",
  "şura",
];
