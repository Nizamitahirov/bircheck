// Verify: a 6-day worker hired BEFORE the period uses the M2 workday total
// (= period length - L holidays), not K2. A regular worker uses K2.
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
const K2 = 26; // regular monthly workdays

// L column has 4 six-day non-work days. Period length = 31 days.
// M2 (leave empty so the processor computes it) = 31 - 4 = 27.
const L = [dToSerial(2026, 5, 3), dToSerial(2026, 5, 10), dToSerial(2026, 5, 17), dToSerial(2026, 5, 24)];
const otp = [
  ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"],
  [0, "", "", "", "", "", "", 0, "", J2, K2, L[0], ""],
  [0, "", "", "", "", "", "", 0, "", J3, "", L[1], ""],
  [0, "", "", "", "", "", "", 0, "", "", "", L[2], ""],
  [0, "", "", "", "", "", "", 0, "", "", "", L[3], ""],
];

// Both hired well before the period (< J2) → they hit the "monthWorkdays - otpusk" branch.
const bazaHeader = ["Kod", "Soyad", "Ad", "Ata", "Vez", "Dep", "Filial", "HireDate", "Iş qrafiki"];
const baza = [
  bazaHeader,
  [301, "Reg", "Ular", "C", "V", "D", "F", dToSerial(2020, 1, 1), "5 günlük iş"],
  [302, "Six", "Day", "Z", "V", "D", "F", dToSerial(2020, 1, 1), "6 günlük iş rejimi"],
];
const mhmd = [["A", "B", "C", "D", "E", "F"]]; // no MHMD → J=0, so diff = K

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["x"]]), "Problems");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(baza), "From_Excel_Baza");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mhmd), "From_MHMD");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(otp), "From_HRB_Otpusk");

const tmp = path.join(__dirname, "_m2.xlsx");
XLSX.writeFile(wb, tmp);
const buf = readFileSync(tmp);
const fileLike = {
  arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
  name: "m2.xlsx",
};
if (typeof performance === "undefined") globalThis.performance = { now: () => Date.now() };
const { processWorkbook } = await import("../dist-smoke/processor.mjs");
const res = await processWorkbook(fileLike);

const reg = res.rows.find((r) => Number(r.cells[0]) === 301);
const six = res.rows.find((r) => Number(r.cells[0]) === 302);
console.log("301 regular  K =", reg?.cells[10], "is6Day =", reg?.is6Day);
console.log("302 6-day    K =", six?.cells[10], "is6Day =", six?.is6Day);

let pass = true;
function check(name, cond) {
  console.log(`${cond ? "PASS" : "FAIL"}  ${name}`);
  if (!cond) pass = false;
}
check("301 uses K2 = 26", reg?.cells[10] === 26);
check("301 is6Day = false", reg?.is6Day === false);
check("302 uses M2 = 27 (31 - 4 L-days)", six?.cells[10] === 27);
check("302 is6Day = true", six?.is6Day === true);

unlinkSync(tmp);
console.log("\n", pass ? "ALL CHECKS PASS" : "SOME CHECKS FAILED");
process.exit(pass ? 0 : 1);
