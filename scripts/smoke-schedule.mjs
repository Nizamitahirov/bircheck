// Verify: a mid-period hire on a 6-day schedule reads holidays from column L,
// everyone else reads from column I. Both cases should produce different K
// values when the same date shows up in one list but not the other.

import * as XLSX from "xlsx";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { readFileSync, unlinkSync } from "node:fs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;
const dToSerial = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - EXCEL_EPOCH_MS) / MS_PER_DAY);

const J2 = dToSerial(2026, 5, 1);
const J3 = dToSerial(2026, 5, 31);
const K2 = 26;

// otpusk: I has 05-11 (in period), L is empty.
// Regular workers exclude 05-11; 6-day workers exclude nothing.
const otp = [
  ["A","B","C","D","E","F","G","H","I","J","K","L"],
  [0,"","","","","","",0, dToSerial(2026,5,11), J2, K2, ""],
  [0,"","","","","","",0, "", J3, "", ""],
];

// Two mid-period hires on 2026-05-11 (a Monday):
//   201 = regular schedule → I-column holiday on the hire day → K excludes it
//   202 = "6 günlük iş" schedule → L-column holiday on 05-14 → K excludes 05-14
const bazaHeader = ["Kod","Soyad","Ad","Ata","Vez","Dep","Filial","HireDate","Iş qrafiki"];
const baza = [
  bazaHeader,
  [201,"A","B","C","V","D","F", dToSerial(2026,5,11), "5 günlük iş"],
  [202,"X","Y","Z","V","D","F", dToSerial(2026,5,11), "6 günlük iş rejimi"],
];

// No MHMD / no otpusk absence for either employee.
const mhmd = [["A","B","C","D","E","F"]];

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["x"]]), "Problems");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(baza), "From_Excel_Baza");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mhmd), "From_MHMD");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(otp), "From_HRB_Otpusk");

const tmp = path.join(__dirname, "_sched.xlsx");
XLSX.writeFile(wb, tmp);
const buf = readFileSync(tmp);
const fileLike = {
  arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
  name: "sched.xlsx",
};
if (typeof performance === "undefined") globalThis.performance = { now: () => Date.now() };
const { processWorkbook } = await import("../dist-smoke/processor.mjs");
const res = await processWorkbook(fileLike);

// NETWORKDAYS.INTL(2026-05-11 .. 2026-05-31, weekend=Sunday, minus holidays):
//   Days 11..31 = 21. Sundays on 17, 24, 31 → 3 removed → base 18.
//   201 (regular, uses I): -1 for 05-11 → 17
//   202 (6-day, uses L empty): -0 → 18
const r201 = res.rows.find(r => Number(r.cells[0]) === 201);
const r202 = res.rows.find(r => Number(r.cells[0]) === 202);
console.log("201 (regular) K =", r201?.cells[10]);
console.log("202 (6-day)   K =", r202?.cells[10]);

let pass = true;
function check(name, cond) {
  console.log(`${cond ? "PASS" : "FAIL"}  ${name}`);
  if (!cond) pass = false;
}
check("201 (regular, I list) K = 17", r201?.cells[10] === 17);
check("202 (6 günlük iş, L empty) K = 18", r202?.cells[10] === 18);
unlinkSync(tmp);
console.log("\n", pass ? "ALL CHECKS PASS" : "SOME CHECKS FAILED");
process.exit(pass ? 0 : 1);
