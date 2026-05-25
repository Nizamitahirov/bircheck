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

const N = 5000; // employees
const M_PER_EMP = 3; // mhmd entries per employee
const O_PER_EMP = 1; // otpusk entries per employee

const baza = [["Kod","Soyad","Ad","Ata","Vez","Dep","Filial","HireDate","Status"]];
for (let i = 1; i <= N; i++) {
  baza.push([1000+i,"S"+i,"A"+i,"At"+i,"V","D","F",dToSerial(2020,1,15),"Aktiv"]);
}
const mhmd = [["A","B","C","D","E","F"]];
for (let i = 1; i <= N; i++) {
  for (let k = 0; k < M_PER_EMP; k++) {
    mhmd.push([1000+i,"","001",J2,J3, k===0 ? 25 : (k===1 ? 1 : 0)]);
  }
}
const otp = [["A","B","C","D","E","F","G","H","I","J","K"]];
otp.push([0,"","","","","","",0,"",J2,K2]);
otp.push([0,"","","","","","",0,"",J3,""]);
for (let i = 1; i <= N; i++) {
  for (let k = 0; k < O_PER_EMP; k++) {
    otp.push([1000+i,"","","",dToSerial(2026,5,5),dToSerial(2026,5,7),"Mez", k===0?0:2,"","",""]);
  }
}

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["x"]]), "Problems");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(baza), "From_Excel_Baza");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mhmd), "From_MHMD");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(otp), "From_HRB_Otpusk");

const tmp = path.join(__dirname, "_perf.xlsx");
XLSX.writeFile(wb, tmp);
const buf = readFileSync(tmp);
const fileLike = { arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength), name:"perf.xlsx" };

if (typeof performance === "undefined") globalThis.performance = { now: () => Date.now() };
const { processWorkbook } = await import("../dist-smoke/processor.mjs");

console.log(`Dataset: ${N} employees, ${mhmd.length-1} MHMD rows, ${otp.length-1} otpusk rows`);
const t0 = Date.now();
const res = await processWorkbook(fileLike);
const ms = Date.now() - t0;
console.log(`Processed in ${ms}ms (${(ms/1000).toFixed(2)}s)`);
console.log(`Rows kept: ${res.rowsKept} / ${res.totalEmployees}`);
unlinkSync(tmp);
