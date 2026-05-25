// Quick smoke test: build a synthetic workbook, run the processor, assert output.
import * as XLSX from "xlsx";
import { fileURLToPath } from "node:url";
import path from "node:path";

// Simulate a File for processor.ts
// Easiest: write to disk and read back.
import { writeFileSync, readFileSync, unlinkSync } from "node:fs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Excel epoch helpers
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;
const dToSerial = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - EXCEL_EPOCH_MS) / MS_PER_DAY);

// Period: May 1..31, 2026
const J2 = dToSerial(2026, 5, 1);
const J3 = dToSerial(2026, 5, 31);
const K2 = 26; // total workdays in month (Mon-Sat, Sunday only weekend, no holidays)

// Holidays I2:I29 — leave empty
const otpuskHeader = ["A_kod", "B", "C", "D", "E_baslama", "F_son", "G_seb", "H_gun", "I_bayram", "J", "K_total"];
const otpuskRows = [
  otpuskHeader,
  // first data row: A=101, H=5, J2/K2 here
  [101, "", "", "", dToSerial(2026,5,5), dToSerial(2026,5,9), "Mezuniyyet", 5, "", J2, K2],
  // J3 in third row col J
  [102, "", "", "", dToSerial(2026,5,10), dToSerial(2026,5,12), "Otpuska", 3, "", J3, ""],
  [103, "", "", "", "", "", "", 0, "", "", ""], // H=0 should NOT add to details
];

// MHMD: A=kod, C=type, D=start, E=end, F=days
const mhmdHeader = ["A_kod", "B", "C_type", "D_start", "E_end", "F_days"];
const mhmdRows = [
  mhmdHeader,
  [101, "", "001", J2, J3, 20],
  [101, "", "001", J2, J3, 1],          // sums to 21 for 101
  [102, "", "001", J2, J3, 23],
  [103, "", "001", J2, J3, 26],
  [104, "", "001", J2, J3, 26],
  // out-of-range row (should be cleaned)
  [101, "", "001", dToSerial(2026,4,1), dToSerial(2026,4,30), 100],
  // wrong type, should be ignored in J sum
  [102, "", "002", J2, J3, 5],
];

// Baza: A=kod, B=Soyad, C=Ad, D=Ata, E=Vəz, F=Dep, G=Filial, H=İşə qəbul, I=Status
const bazaHeader = ["Kod","Soyad","Ad","Ata","Vez","Dep","Filial","HireDate","Status"];
const bazaRows = [
  bazaHeader,
  [101, "Aliyev","Ali","Hüseyn","Mühasib","Maliyyə","Bakı", dToSerial(2020,1,15), "Aktiv"],
  [102, "Mammadov","Vüsal","Rəşid","Operator","IT","Bakı", dToSerial(2020,3,20), "Aktiv"],
  [103, "Hasanov","Tural","Niyaz","Menecer","HR","Bakı", dToSerial(2020,6,10), "Aktiv"],
  [104, "Quliyev","Rauf","Akif","Analitik","Maliyyə","Bakı", dToSerial(2026,5,20), "Aktiv"], // mid-period hire
];

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["dummy"]]), "Problems");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(bazaRows), "From_Excel_Baza");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mhmdRows), "From_MHMD");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(otpuskRows), "From_HRB_Otpusk");

const tmp = path.join(__dirname, "_smoke.xlsx");
XLSX.writeFile(wb, tmp);

const buf = readFileSync(tmp);

// Polyfill File-like object: processor uses .arrayBuffer()
const fileLike = {
  arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
  name: "smoke.xlsx",
};

// performance.now polyfill
if (typeof performance === "undefined") {
  globalThis.performance = { now: () => Date.now() };
}

const { processWorkbook } = await import("../dist-smoke/processor.mjs");

const res = await processWorkbook(fileLike, (s, p) => console.log(`  [${p}%] ${s}`));

console.log("\nProblems rows from result:");
for (const r of res.rows) console.log(JSON.stringify(r.cells));

const ok =
  res.rows.length === 1 &&
  Number(res.rows[0].cells[0]) === 104 &&
  Number(res.rows[0].cells[11]) === -16;
console.log("Kept rows:", res.rows.length, "Expected: 1");
console.log("Smoke result:", ok ? "PASS" : "FAIL");

unlinkSync(tmp);
process.exit(ok ? 0 : 1);
