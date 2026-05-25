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

const out = XLSX.utils.sheet_to_json(res.workbook.Sheets.Problems, { header: 1, defval: null });
console.log("\nProblems sheet output:");
for (const row of out) console.log(JSON.stringify(row));

// Expected:
// 101: J=21, K=NETWORKDAYS(2020-01-15..2026-05-31, sunday only, no holidays) wait, hire < J2 so K = K2 - otpusk(5) = 21 → L = 0 → REMOVED
// 102: J=23, hire < J2, K = 26 - 3 = 23 → L = 0 → REMOVED
// 103: J=26, hire < J2, K = 26 - 0 = 26 → L = 0 → REMOVED
// 104: J=26, hire (2026-05-20) >= J2, K = NETWORKDAYS(2026-05-20..2026-05-31, Sunday weekend) - 0
//     2026-05-20 (Wed) thru 2026-05-31 (Sun): days = Wed,Thu,Fri,Sat,(Sun skip),Mon,Tue,Wed,Thu,Fri,Sat,(Sun skip) = 10
//     L = 10 - 26 = -16 → KEPT

const dataRows = out.slice(1);
console.log("\nKept rows:", dataRows.length);
console.log("Expected kept: 1 (employee 104)");
const ok = dataRows.length === 1 && Number(dataRows[0][0]) === 104 && Number(dataRows[0][11]) === -16;
console.log("Smoke result:", ok ? "PASS ✅" : "FAIL ❌");

unlinkSync(tmp);
process.exit(ok ? 0 : 1);
