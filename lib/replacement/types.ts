import type { SourceName } from "./config";

export type Cell = string | number | boolean | Date | null | undefined;
export type Row = Cell[];

export type BOF = "Baş ofis" | "Filial";

/** SPEC bölmə 4 — vahid sxem. */
export interface NormRow {
  menbe: SourceName;
  /** Orijinal fayldakı Excel sətir nömrəsi; birləşmədə vergüllə siyahı. */
  menbeSetri: string;
  /** Numerik sətir nömrələri (prioritet/sıralama üçün). */
  rowNums: number[];
  evezEdenKod: string;
  evezEdenAd: string;
  evezEdenStruktur: string;
  evezEdenVezife: string;
  evezEdilenKod: string;
  evezEdilenAd: string;
  evezEdilenStruktur: string;
  evezEdilenVezife: string;
  teyinYeri: string;
  basOfis: BOF;
  evezetmeNovu: string;
  /** Excel serial; saatlıq sətirlərdə kəsrli. */
  baslama: number;
  /** null = açıq bitmə. */
  bitme: number | null;
  xerc: boolean;
  qeydler: string[];
  dayBased: boolean;

  // Mərhələ 2-də hesablanır:
  effEnd?: number;
  dayCount?: number;
  hourCount?: number;
  hasXeta?: boolean;
}

export interface ErrorLogEntry {
  menbe: string;
  setr: string;
  novu: string;
  teferruat: string;
  setirMelumati: string;
}

export interface DupEntry {
  novu: "Fayl daxili tam duplikat" | "Fayllar arası duplikat";
  kod: string;
  ad: string;
  baslama: number;
  bitme: number | null;
  saxlanilanMenbe: string;
  cixarilanMenbe: string;
  qeyd: string;
  /** çıxarılan sətrin mənbəyi/xərc dəyəri (yoxlama 12.7 üçün) */
  removedMenbe: SourceName;
  removedXerc: boolean;
  keptMenbe: SourceName;
}

/** Müvəqqəti keçidin ləğvi/dayandırılması əmrləri (SPEC 3.4). */
export interface CancelInfo {
  kod: string;
  baslama: number | null;
  bitme: number | null;
  nov: string;
  emr: string;
  rowNum: number;
}

export interface BackupSpec {
  sheetName: string;
  note: string;
  aoa: Row[];
}

export interface CemRow {
  kod: string;
  ad: string;
  say: number;
  cemSaat: number;
  cemGun: number;
  unikalGun: number;
  isGunu: number;
  hefteSonu: number;
}

export interface CheckResult {
  name: string;
  ok: boolean;
  detail: string;
  /** hard=true uğursuzdursa nəticə qaytarılmır (SPEC 12). */
  hard: boolean;
}
