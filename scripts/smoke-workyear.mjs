// Smoke test: Vacation by Workyear pipeline (vacation_days_by_period_SPEC.md).
// Build first:
//   npx esbuild lib/vacation-workyear.ts --bundle --format=esm --external:xlsx --outfile=dist-smoke/vacation-workyear.mjs
// Run:
//   node scripts/smoke-workyear.mjs
import * as XLSX from "xlsx";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { readFileSync, unlinkSync } from "node:fs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;
const S = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - EXCEL_EPOCH_MS) / MS_PER_DAY);

// Master / illik səhifə sətri (23 sütun; header 2 sətir, data 3-cü sətirdən)
function mrow({ badge, name, hire, term = "", dept = "", div = "", pos = "", main = 0, child = 0, exper = 0, bad = 0, used = "", balance = "" }) {
  const r = new Array(23).fill("");
  r[0] = badge; r[1] = name; r[3] = hire; r[4] = term; r[6] = dept; r[7] = div; r[10] = pos;
  r[14] = main; r[15] = child; r[16] = exper; r[17] = bad; r[21] = used; r[22] = balance;
  return r;
}
const H2 = [new Array(23).fill(""), new Array(23).fill("")]; // 2 sətirlik header imitasiyası

const master = [
  ...H2,
  mrow({ badge: "COM 1", name: "Aliyev Ali A", hire: "15.03.2023", dept: "IT", div: "Dev", pos: "Engineer", main: 21, exper: 2 }),
  mrow({ badge: "COM 2", name: "Huseynova Leyla H", hire: "01.02.2025", dept: "HR", pos: "Specialist", main: 30 }),
  mrow({ badge: "COM 3", name: "Mammadov Elnur M", hire: "10.06.2024", dept: "Ops", pos: "Cashier", main: 21 }), // rehire
  mrow({ badge: "COM 4", name: "Qasimov Anar Q", hire: "01.01.2026", dept: "IT", pos: "Analyst", main: 21 }), // avans halı
];

const y2023 = [...H2, mrow({ badge: "COM 1", name: "Aliyev Ali A", hire: "15.03.2023", main: 21, exper: 2, used: 14 })];
const y2024 = [
  ...H2,
  // köhnə iş dövrü sətri ƏVVƏL gəlir — rehire filtri onu ötüb düzgün sətri seçməlidir
  mrow({ badge: "COM 3", name: "Mammadov Elnur M", hire: "01.01.2019", main: 99, used: 0 }),
  mrow({ badge: "COM 3", name: "Mammadov Elnur M", hire: "10.06.2024", main: 21, used: 30 }),
  mrow({ badge: "COM 1", name: "Aliyev Ali A", hire: "15.03.2023", main: 21, exper: 2, used: 5 }),
];
const y2025 = [
  ...H2,
  mrow({ badge: "COM 1", name: "Aliyev Ali A", hire: "15.03.2023", main: 30, exper: 2, used: 0 }), // 21→30 keçidi
  mrow({ badge: "COM 2", name: "Huseynova Leyla H", hire: "01.02.2025", main: 30, used: 10 }),
  mrow({ badge: "COM 3", name: "Mammadov Elnur M", hire: "10.06.2024", main: 21, used: 0 }),
];
const y2026 = [
  ...H2,
  mrow({ badge: "COM 2", name: "Huseynova Leyla H", hire: "01.02.2025", main: 30, used: 0 }),
  mrow({ badge: "COM 4", name: "Qasimov Anar Q", hire: "01.01.2026", main: 21, used: 25 }),
];

const orders = [
  ["Работник", "Badge", "Поле1", "Шаблон", "Готов к утверждению", "from", "to", "(дн)", "Дата приказа"],
  ["Aliyev Ali A", "COM 1", "Отпуск очередной", "Main Vacation", "Да", "01.07.2023", "14.07.2023", 14, "20.06.2023"],
  ["Aliyev Ali A", "COM 1", "Отпуск очередной", "Main Vacation", "Да", "01.07.2023", "14.07.2023", 14, "21.06.2023"], // dublikat (fərqli əmr tarixi)
  ["Aliyev Ali A", "COM 1", "Больничный", "", "Да", "01.09.2023", "07.09.2023", 7, "01.09.2023"], // istisna növ
  ["Aliyev Ali A", "COM 1", "Отпуск очередной", "Main Vacation", "Нет", "01.10.2023", "03.10.2023", 3, "25.09.2023"], // təsdiqsiz
  ["Aliyev Ali A", "COM 1", "Компенсация", "", "Да", "", "", 5, "05.05.2024"], // from boş → əmr tarixi
  ["Huseynova Leyla H", "COM 2", "Отпуск очередной", "Main Vacation", "Да", "10.03.2025", "19.03.2025", 10, "01.03.2025"],
  ["Mammadov Elnur M", "COM3", "Отпуск очередной", "Main Vacation", "Да", "05.05.2019", "14.05.2019", 10, "01.05.2019"], // əvvəlki iş dövrü
  ["Mammadov Elnur M", "COM 3", "Отпуск очередной", "Main Vacation", "Да", "01.08.2024", "30.08.2024", 30, "20.07.2024"],
  ["Qasimov Anar Q", "COM 4", "Отпуск очередной", "Main Vacation", "Да", "10.01.2026", "03.02.2026", 25, "05.01.2026"], // avans (25 > 21)
  ["Nobody X", "COM 9", "Отпуск очередной", "Main Vacation", "Да", "01.02.2026", "05.02.2026", 5, "25.01.2026"], // masterdə yoxdur
];

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(master), "2015-2026");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(y2023), "2023");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(y2024), "2024");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(y2025), "2025");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(y2026), "2026");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(orders), "orders");
const tmp = path.join(__dirname, "_wy.xlsx");
XLSX.writeFile(wb, tmp);

const buf = readFileSync(tmp);
const fileLike = {
  name: "_wy.xlsx",
  arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
};
if (typeof performance === "undefined") globalThis.performance = { now: () => Date.now() };

const { processWorkyear, DEFAULT_WORKYEAR_PARAMS } = await import("../dist-smoke/vacation-workyear.mjs");

const res = await processWorkyear(
  fileLike,
  { ...DEFAULT_WORKYEAR_PARAMS, asOfDate: S(2026, 7, 7) },
  (s, p) => console.log(`  [${p}%] ${s}`),
);

console.log("\n=== Yoxlamalar ===");
for (const c of res.checks) console.log(`${c.ok ? "PASS" : "FAIL"}${c.hard ? "" : " (info)"}  ${c.name}: ${c.detail}`);

console.log("\n=== Dövrlər ===");
const F = (s) => new Date(EXCEL_EPOCH_MS + s * MS_PER_DAY).toISOString().slice(0, 10);
for (const r of res.rows)
  console.log(`${r.empCode} | ${F(r.periodStart)}–${F(r.periodEnd)} | main ${r.mainDays} add ${r.addDays} | used ${r.usedMain}/${r.usedAdd}`);

console.log("\n=== Loq ===");
for (const l of res.logs) console.log(`${l.tip} | ${l.emp} | ${l.detail}`);

let pass = true;
function check(name, cond) {
  console.log(`${cond ? "PASS" : "FAIL"}  ${name}`);
  if (!cond) pass = false;
}

console.log("\n=== Assertions ===");
check("bütün məcburi yoxlamalar keçdi", res.allPassed === true);
check("əməkdaş = 4, dövr = 10", res.stats.employees === 4 && res.rows.length === 10);
check("istifadə cəmi = 84", res.stats.totalUsed === 84);
check("sayılan əmr = 5, dublikat = 1", res.stats.ordersIncluded === 5 && res.stats.ordersDedup === 1);
check("Used uzlaşması = 100%", res.stats.matchPct === 100);

const e1 = res.rows.filter((r) => r.empCode === "COM 1");
check("E1: 4 dövr", e1.length === 4);
check("E1 P1: 15.03.2023–14.03.2024, main 21, add 2, used 19", e1[0] && e1[0].periodStart === S(2023, 3, 15) && e1[0].periodEnd === S(2024, 3, 14) && e1[0].mainDays === 21 && e1[0].addDays === 2 && e1[0].usedMain === 19 && e1[0].usedAdd === 0);
check("E1 P3: norma keçidi 21→30", e1[2] && e1[2].periodStart === S(2025, 3, 15) && e1[2].mainDays === 30);
check("E1 P4: 2026 üçün ən yaxın il (2025) norması", e1[3] && e1[3].mainDays === 30);
check("E1 P2–P4 istifadəsiz", e1.slice(1).every((r) => r.usedMain === 0 && r.usedAdd === 0));

const e2 = res.rows.filter((r) => r.empCode === "COM 2");
check("E2: 2 dövr, P1 used 10", e2.length === 2 && e2[0].usedMain === 10);

const e3 = res.rows.filter((r) => r.empCode === "COM 3");
check("E3: 3 dövr", e3.length === 3);
check("E3 rehire filtri: norma 21 (99 deyil)", e3.every((r) => r.mainDays === 21));
check("E3 FIFO: P1 used 21, P2 used 9", e3[0].usedMain === 21 && e3[1].usedMain === 9 && e3[2].usedMain === 0);
check("E3 köhnə dövr əmri sayılmayıb (30, 40 deyil)", e3.reduce((s, r) => s + r.usedMain + r.usedAdd, 0) === 30);

const e4 = res.rows.filter((r) => r.empCode === "COM 4");
check("E4: 1 dövr, avans used 25 (main 21)", e4.length === 1 && e4[0].usedMain === 25 && e4[0].mainDays === 21);
check("E4 avans loqu", res.logs.some((l) => l.tip === "Avans məzuniyyət" && l.emp === "Qasimov Anar Q"));
check("əvvəlki iş dövrü loqu (E3)", res.logs.some((l) => l.tip === "Əvvəlki iş dövrü"));
check("masterdə olmayan əmr loqu", res.logs.some((l) => l.tip === "Master siyahıda yoxdur"));

const ws = res.workbook.Sheets["Vacation days by period"];
check("nəticə səhifəsi mövcuddur", !!ws);
const k2 = ws && ws["K2"];
check("K2: düstur I2+J2 + cache dəyəri", !!k2 && k2.f === "I2+J2" && k2.v === 23);
const q = ws && ws[`Q${1 + res.rows.length}`];
check("son sətirdə Q düsturu", !!q && typeof q.f === "string" && q.f.startsWith("K"));
check("autofilter A1:Q11", !!ws && ws["!autofilter"] && ws["!autofilter"].ref === "A1:Q11");
const e4row = res.rows.findIndex((r) => r.empCode === "COM 4") + 2;
const m4 = ws && ws[`M${e4row}`];
check("E4 unused main = -4 (avans gizlədilmir)", !!m4 && m4.v === -4);

unlinkSync(tmp);
console.log("\n", pass ? "ALL CHECKS PASS" : "SOME CHECKS FAILED");
process.exit(pass ? 0 : 1);
