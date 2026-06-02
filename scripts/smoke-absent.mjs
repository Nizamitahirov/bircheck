import * as XLSX from "xlsx";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { readFileSync, unlinkSync } from "node:fs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Build a Data sheet covering today and the past 90 days. The processor should
// clip to the last 63 days, drop types 42 / 29, dedup on (code, date) keeping
// highest type, compute D weights and apply the chain filter.

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;
function toSerial(date) {
  const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  return Math.round((utc - EXCEL_EPOCH_MS) / MS_PER_DAY);
}
function daysAgo(n) {
  const d = new Date();
  d.setDate(d.getDate() - n);
  return toSerial(d);
}

const header = ["Personal kod", "Ad", "Növ", "Status", "Başlama", "Bitmə"];
const data = [
  header,
  // emp 1: clean recent absence (3 consecutive days)
  [1001, "Ali", 15, "", daysAgo(5), daysAgo(3)],

  // emp 2: a chain with a gap >3, plus a recent recurrence — chain filter should keep only the latest
  [1002, "Veli", 15, "", daysAgo(50), daysAgo(48)],
  [1002, "Veli", 15, "", daysAgo(40), daysAgo(38)],
  [1002, "Veli", 15, "", daysAgo(2), daysAgo(2)],

  // emp 3: only type 42 → should be filtered out entirely
  [1003, "Samir", 42, "", daysAgo(10), daysAgo(8)],

  // emp 4: type 30 close to another absence — should yield D=-1
  [1004, "Rauf", 15, "", daysAgo(7), daysAgo(7)],
  [1004, "Rauf", 30, "", daysAgo(6), daysAgo(6)],

  // emp 5: out-of-window (older than 63 days) — should be dropped
  [1005, "Eldar", 15, "", daysAgo(90), daysAgo(85)],

  // emp 6: window straddle (start before window, end inside) — start clipped
  [1006, "Tural", 15, "", daysAgo(70), daysAgo(60)],
];

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), "Data");
const tmp = path.join(__dirname, "_absent.xlsx");
XLSX.writeFile(wb, tmp);
const buf = readFileSync(tmp);

const fileLike = {
  arrayBuffer: async () =>
    buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
  name: "absent.xlsx",
};
if (typeof performance === "undefined") globalThis.performance = { now: () => Date.now() };

const { processAbsenteeism } = await import("../dist-smoke/absenteeism.mjs");

const res = await processAbsenteeism(fileLike, (s, p) => console.log(`  [${p}%] ${s}`));

console.log("\n=== Netice ===");
for (const r of res.netice) console.log(`${r.code}\t${r.name}\t${r.total}`);

console.log("\n=== Muqayise ===");
for (const m of res.muqayise) {
  const d = new Date(EXCEL_EPOCH_MS + m.dateSerial * MS_PER_DAY);
  console.log(`${m.code}\t${d.toISOString().slice(0, 10)}\ttype=${m.type}\tD=${m.weight}`);
}

const e1 = res.netice.find((r) => r.code === "1001");
const e2 = res.netice.find((r) => r.code === "1002");
const e3 = res.netice.find((r) => r.code === "1003");
const e4 = res.netice.find((r) => r.code === "1004");
const e5 = res.netice.find((r) => r.code === "1005");
const e6 = res.netice.find((r) => r.code === "1006");

let pass = true;
function check(name, cond) {
  console.log(`${cond ? "PASS" : "FAIL"}  ${name}`);
  if (!cond) pass = false;
}

// 1001: 3 days, all within 63d window: dates daysAgo(5), 4, 3 → D=1,1,1 → total=3
check("emp 1001 total = 3", e1 && e1.total === 3);

// 1002: chain filter — earlier absences erased after the big gap → only the latest day (D=1) survives
check("emp 1002 total = 1", e2 && e2.total === 1);

// 1003: all rows are type 42 → filtered out, total = 0 (still appears as a unique code)
check("emp 1003 total = 0", e3 && e3.total === 0);

// 1004: day -7 (type 15, D=1), day -6 (type 30, gap=1 ≤ 3 → D=-1). Total = 0
check("emp 1004 total = 0", e4 && e4.total === 0);

// 1005: out of window → not present
check("emp 1005 absent from results", !e5);

// 1006: start clipped to today-63, end at today-60 → days -63..-60 = 4 days
check("emp 1006 total = 4", e6 && e6.total === 4);

unlinkSync(tmp);
console.log("\n", pass ? "ALL CHECKS PASS" : "SOME CHECKS FAILED");
process.exit(pass ? 0 : 1);
