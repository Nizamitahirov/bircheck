import * as XLSX from "xlsx";
import { EXCEL_EPOCH_MS, MS_PER_DAY } from "./config";
import type { Cell, Row } from "./types";

export function dateFromSerial(s: number): Date {
  return new Date(EXCEL_EPOCH_MS + Math.round(s * MS_PER_DAY));
}

export function serialFromDateObj(d: Date): number {
  const utc = Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes());
  return (utc - EXCEL_EPOCH_MS) / MS_PER_DAY;
}

export function fmtDate(s: number | null): string {
  if (s === null || s === undefined || !Number.isFinite(s)) return "";
  const d = dateFromSerial(Math.floor(s));
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  return `${dd}.${mm}.${d.getUTCFullYear()}`;
}

export function fmtDateTime(s: number | null): string {
  if (s === null || s === undefined || !Number.isFinite(s)) return "";
  const base = fmtDate(s);
  const frac = s - Math.floor(s);
  const totalMin = Math.round(frac * 24 * 60);
  const hh = String(Math.floor(totalMin / 60)).padStart(2, "0");
  const mi = String(totalMin % 60).padStart(2, "0");
  return frac > 0 ? `${base} ${hh}:${mi}` : base;
}

/** 0=Bazar … 6=Şənbə (UTC). Həftəsonu: Şənbə + Bazar. */
export function isWeekend(daySerial: number): boolean {
  const wd = dateFromSerial(daySerial).getUTCDay();
  return wd === 0 || wd === 6;
}

export function serialYear(s: number): number {
  return dateFromSerial(Math.floor(s)).getUTCFullYear();
}

export function serialMonth(s: number): number {
  return dateFromSerial(Math.floor(s)).getUTCMonth() + 1;
}

/** Azərbaycan əlifbası üçün etibarlı kiçik hərf (İ→i, I→ı, birləşən nöqtə silinir). */
export function azLower(s: string): string {
  return s
    .replace(/İ/g, "i")
    .replace(/I/g, "ı")
    .toLowerCase()
    .replace(/̇/g, "");
}

/** SPEC 6 — kod normallaşdırması: str(int(float(x))). */
export function normCode(v: Cell): string {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number") return Number.isFinite(v) ? String(Math.trunc(v)) : "";
  const s = String(v).trim();
  if (/^-?\d+(?:[.,]\d+)?$/.test(s)) return String(Math.trunc(parseFloat(s.replace(",", "."))));
  return s;
}

export function cellStr(v: Cell): string {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return fmtDate(serialFromDateObj(v));
  return String(v).trim();
}

export type ParsedDate = number | null | "bad";

/** Universal tarix oxunması → Excel serial. null = boş, "bad" = oxunmadı. */
export function parseDateValue(v: Cell): ParsedDate {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") {
    if (!Number.isFinite(v)) return "bad";
    if (v >= 19000101) return parseYYYYMMDD(v); // int YYYYMMDD
    if (v > 0 && v < 2958466) return v; // Excel serial
    return "bad";
  }
  if (v instanceof Date) return serialFromDateObj(v);
  const s = String(v).trim();
  if (!s) return null;
  let m = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (m) {
    const base = ymdToSerial(+m[3], +m[2], +m[1]);
    if (base === "bad") return "bad";
    if (m[4]) return base + (+m[4] * 60 + +m[5]) / 1440;
    return base;
  }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s](\d{1,2}):(\d{2}))?/);
  if (m) {
    const base = ymdToSerial(+m[1], +m[2], +m[3]);
    if (base === "bad") return "bad";
    if (m[4]) return base + (+m[4] * 60 + +m[5]) / 1440;
    return base;
  }
  m = s.match(/^(\d{8})(?:\.0+)?$/);
  if (m) return parseYYYYMMDD(+m[1]);
  return "bad";
}

/** SPEC 3.2 — int YYYYMMDD formatı (məs. 20260213). */
export function parseYYYYMMDD(v: Cell): ParsedDate {
  if (v === null || v === undefined || v === "") return null;
  const n = typeof v === "number" ? Math.trunc(v) : parseInt(String(v).trim(), 10);
  if (!Number.isFinite(n)) return "bad";
  const y = Math.trunc(n / 10000);
  const mo = Math.trunc((n % 10000) / 100);
  const d = n % 100;
  return ymdToSerial(y, mo, d);
}

function ymdToSerial(y: number, mo: number, d: number): number | "bad" {
  if (y < 1900 || y > 3000 || mo < 1 || mo > 12 || d < 1 || d > 31) return "bad";
  const utc = Date.UTC(y, mo - 1, d);
  const check = new Date(utc);
  if (check.getUTCMonth() + 1 !== mo || check.getUTCDate() !== d) return "bad";
  return Math.round((utc - EXCEL_EPOCH_MS) / MS_PER_DAY);
}

export function sheetToRows(ws: XLSX.WorkSheet): Row[] {
  return XLSX.utils.sheet_to_json<Row>(ws, {
    header: 1,
    raw: true,
    defval: null,
    blankrows: true, // Excel sətir nömrələri sürüşməsin deyə boş sətirlər saxlanılır
  });
}

/** Sheet adına görə tapır: əvvəl tam uyğunluq, sonra contains (azLower). */
export function findSheet(wb: XLSX.WorkBook, candidates: string[]): string | null {
  const names = wb.SheetNames;
  for (const c of candidates) {
    const hit = names.find((n) => azLower(n.trim()) === azLower(c));
    if (hit) return hit;
  }
  for (const c of candidates) {
    const hit = names.find((n) => azLower(n).includes(azLower(c)));
    if (hit) return hit;
  }
  return null;
}

/**
 * Başlıq xəritəsi: ad → sütun indeksi. Adlar strip edilir, \n boşluqla əvəzlənir,
 * təkrarlanan adlar pandas kimi ".1", ".2" şəkilçisi alır.
 */
export function headerMap(header: Row): Map<string, number> {
  const map = new Map<string, number>();
  const seen = new Map<string, number>();
  for (let i = 0; i < header.length; i++) {
    const raw = header[i];
    if (raw === null || raw === undefined || raw === "") continue;
    const name = String(raw).replace(/\n/g, " ").replace(/\s+/g, " ").trim();
    if (!name) continue;
    const count = seen.get(name) ?? 0;
    seen.set(name, count + 1);
    map.set(count === 0 ? name : `${name}.${count}`, i);
  }
  return map;
}

/** Xəta Loqu üçün orijinal sətrin dump-u: `sütun=dəyər | ...` (boşlar buraxılır). */
export function rowDump(header: Row, row: Row): string {
  const hm: string[] = [];
  const names: string[] = [];
  const seen = new Map<string, number>();
  for (let i = 0; i < header.length; i++) {
    const raw = header[i];
    let name = raw === null || raw === undefined ? `Sütun${i + 1}` : String(raw).replace(/\n/g, " ").replace(/\s+/g, " ").trim();
    if (!name) name = `Sütun${i + 1}`;
    const c = seen.get(name) ?? 0;
    seen.set(name, c + 1);
    names.push(c === 0 ? name : `${name}.${c}`);
  }
  for (let i = 0; i < header.length; i++) {
    const v = row[i];
    if (v === null || v === undefined || v === "") continue;
    hm.push(`${names[i]}=${cellStr(v)}`);
  }
  return hm.join(" | ");
}

export function getCol(hm: Map<string, number>, row: Row, name: string): Cell {
  const idx = hm.get(name);
  return idx === undefined ? null : row[idx];
}

export function getColStr(hm: Map<string, number>, row: Row, name: string): string {
  return cellStr(getCol(hm, row, name));
}
