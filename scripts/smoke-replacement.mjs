// Smoke test: Replacement check pipeline (SPEC.md).
// Build first:
//   npx esbuild lib/replacement/index.ts --bundle --format=esm --external:xlsx --outfile=dist-smoke/replacement.mjs
// Run:
//   node scripts/smoke-replacement.mjs
import * as XLSX from "xlsx";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { readFileSync, writeFileSync, unlinkSync } from "node:fs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;
const S = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - EXCEL_EPOCH_MS) / MS_PER_DAY);

// ── 1. İnzibatçılıq prosesi ──────────────────────────────────────────────────
const inzibatci = [
  ["Personal Kod", "Evezeden Emekdas", "Evezedilen Emekdas", "Baslama Tarixi", "Bitme Tarixi", "Period", "Teyin Olundugu Yer", "Növ"],
  // 101: davamlı iki sətir → epizoda birləşməlidir (22 gün)
  [101, "Elvin Əliyev", "", S(2026, 1, 10), S(2026, 1, 20), "{1}Müvəqqəti", "Gəncə filialı", "2-ci xəzinə inzibatçısı"],
  [101, "Elvin Əliyev", "", S(2026, 1, 21), S(2026, 1, 31), "{1}Müvəqqəti", "Gəncə filialı", "2-ci xəzinə inzibatçısı"],
  // 110: Müvəqqəti keçidlə fayllar arası duplikat + xərc konflikti (burada Xeyr)
  [110, "Nigar Həsənova", "", S(2026, 2, 1), S(2026, 2, 5), "{1}Müvəqqəti", "Sahil", "3-cü xəzinə inzibatçısı"],
  // 120: kassa məsul şəxsi → Xeyr
  [120, "Tural Quliyev", "", S(2026, 3, 1), S(2026, 3, 5), "{0}Daimi", "Nizami filialı", "1-ci xəzinə inzibatçısı - kassa əməliyyatlarına məsul şəxs"],
  // 121: 2026-dan əvvəl → filtrdən kənar
  [121, "Rəna İsmayılova", "", S(2025, 12, 1), S(2025, 12, 20), "{1}Müvəqqəti", "Gəncə filialı", "2-ci xəzinə inzibatçısı"],
];

// ── 2. Vacation / time off ──────────────────────────────────────────────────
const vacHeader = [
  "EVEZ_EDEN_PERSONAL_KOD", "EVEZ_EDEN_FULL_NAME", "STRUKTUR", "SHOBE", "BOLME", "VEZIFE",
  "HEVALE_STRUKTUR", "HEVALE_VEZIFE", "MEZUNIYYETIN_BASHLAMA_TARIXI", "MEZUNIYYETIN_BITME_TARIXI",
  "ORDER_NUMBER", "ORDER_DATE", "EVEZ_ETDIYI_SEXS_PERSONAL_KOD", "EVEZ_ETDIYI_FULL_NAME",
  "STRUKTUR", "SHOBE", "BOLME", "VEZIFE", "HEVALE_STRUKTUR", "HEVALE_VEZIFE", "BO/Filial",
];
const vacRow = (kod, ad, vez, bas, bit, ekod, ead, evez, bof) => [
  kod, ad, "Struktur A", "", "", vez, "", "", bas, bit, "OR-1", 20260101, ekod, ead, "Struktur B", "", "", evez, "", "", bof,
];
const vacation = [
  vacHeader,
  vacRow(201, "Aynur Məmmədova", "Mütəxəssis", 20260213, 20260220, 601, "Lalə Rzayeva", "Baş mütəxəssis", "BO"), // Bəli, 8 gün
  vacRow(201, "Aynur Məmmədova", "Mütəxəssis", 20260213, 20260220, 601, "Lalə Rzayeva", "Baş mütəxəssis", "BO"), // fayl daxili tam duplikat
  vacRow(205, "Samir Bağırov", "Mütəxəssis", 20260301, 20260310, 602, "Cavid Nəbiyev", "Mütəxəssis", "BO"), // üst-üstə düşmə (a)
  vacRow(205, "Samir Bağırov", "Mütəxəssis", 20260305, 20260312, 603, "Fidan Kərimli", "Mütəxəssis", "BO"), // üst-üstə düşmə (b)
  vacRow(210, "Günel Sadıqova", "Mütəxəssis", 20260601, 20290305, 604, "Emin Talıbov", "Mütəxəssis", "BO"), // şübhəli bitmə → açıq
  vacRow(220, "Kənan Abbasov", "Müdir", 20260501, 20260505, 605, "Vüsal Orucov", "Bölmə rəhbəri", "Filial"), // Xeyr
  vacRow(301, "Xəyal Musayev", "Mütəxəssis", 20260401, 20260410, 302, "Ramil Cəfərov", "Mütəxəssis", "BO"), // iç-içə (A)
  vacRow(303, "Sevinc Hüseynova", "Mütəxəssis", 20260405, 20260408, 301, "Xəyal Musayev", "Mütəxəssis", "BO"), // iç-içə (B)
];

const toHeader = [
  "Source_type", "REPLACING_EMPLOYEE", "Replacing_full_name", "REPLACING_STRUCTURE_NAME",
  "REPLACING_SECTION_NAME", "REPLACING_SUB_SECTION_NAME", "REPLACING_POSITION", "HR_CODE",
  "STRUCTURE_NAME", "SECTION_NAME", "SUB_SECTION_NAME", "POSITION", "Full_Name", "TIME_OFF_DATE",
  "TIME_FROM_TO", "Acting", "Baş ofis/filial",
];
const timeOff = [
  toHeader,
  ["TO", 401, "Zaur Məlikov", "Əməliyyat departamenti", "", "", "Mütəxəssis", 606, "Əməliyyat departamenti", "", "", "Baş mütəxəssis", "Aysel Qasımova", 20260415, "10:00 14:00", "Şöbə müdiri", "BO"], // 4 saat
  ["TO", 402, "Orxan Salamov", "Əməliyyat departamenti", "", "", "Mütəxəssis", 607, "Əməliyyat departamenti", "", "", "Mütəxəssis", "Nərmin Axundova", 20260416, "14:00 10:00", "", "BO"], // məntiqsiz saat aralığı
];

// ── 3. Filial baza (Müvəqqəti keçid) ─────────────────────────────────────────
const filial = [
  ["Ezamın Növü", "Personal kod", "Soyad Ad Ata adı", "Departament", "Şöbə", "Bölmə", "Vəzifə", "Həva olunmuş vəzifələ", "Ezam olunduğu departament/filial", "Ezamın Başlama tarixi", "Ezamın Bitmə tarixi", "Əvəzedici şəxs", "", "Overtime", "Əmrin nömrəsi", "Əmrin tarixi"],
  ["Müvəqqəti keçid", 110, "Nigar Həsənova", "Sahil", "", "", "Mütəxəssis", "", "Filial müdiri vəzifəsini icra edən şəxs (Gəncə)", S(2026, 2, 1), S(2026, 2, 5), "", "", "", "ƏM-10", S(2026, 1, 30)], // Bəli; İnzibatçı ilə duplikat
  ["Müvəqqəti keçid", 111, "Elnur Şirinov", "Sahil", "", "", "Mütəxəssis", "", "Sahil filialı", S(2026, 4, 1), S(2026, 4, 10), "", "", "", "ƏM-20", S(2026, 3, 28)], // Xeyr; ləğv qeydi alacaq
  ["Müvəqqəti keçidin ləğvi", 111, "Elnur Şirinov", "Sahil", "", "", "Mütəxəssis", "", "Sahil filialı", S(2026, 4, 1), S(2026, 4, 10), "", "", "", "ƏM-77", S(2026, 4, 2)],
  ["Müvəqqəti keçid", 112, "Vaqif Nağıyev", "Sahil", "", "", "Mütəxəssis", "", "Sahil filialı", S(2026, 5, 10), S(2026, 5, 1), "", "", "", "ƏM-30", S(2026, 5, 8)], // tarix məntiqsizliyi
];

// ── 4. Vəzifə-Maaş (lookup) ──────────────────────────────────────────────────
const vmSheet = [
  ["Personal kod", "Soyadı, adı və ata adı", "Departament", "Şöbə", "Bölmə", "Vəzifə"],
  [101, "Elvin Əliyev", "Gəncə filialı", "", "", "2-ci xəzinə inzibatçısı"],
];

// ── 5. Rəsmi əvəzetmə: qırmızı sətirli hazır fixture (SheetJS CE fill yaza bilmir) ──
const RESMI_B64 =
  "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";

function wbFile(name, sheets) {
  const wb = XLSX.utils.book_new();
  for (const [sn, aoa] of sheets) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), sn);
  const tmp = path.join(__dirname, name);
  XLSX.writeFile(wb, tmp);
  return tmp;
}

const paths = [
  wbFile("_inz.xlsx", [["Sheet1", inzibatci]]),
  wbFile("_vac.xlsx", [["Vacation replacement", vacation], ["time off replacement", timeOff]]),
  wbFile("_fil.xlsx", [["Sheet1", filial]]),
  path.join(__dirname, "_res.xlsx"),
  wbFile("_vm.xlsx", [["Yanvar 2026", vmSheet]]),
];
writeFileSync(paths[3], Buffer.from(RESMI_B64, "base64"));

const fileLike = (p) => {
  const buf = readFileSync(p);
  return {
    name: path.basename(p),
    arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
  };
};
if (typeof performance === "undefined") globalThis.performance = { now: () => Date.now() };

const { processReplacement, DEFAULT_PARAMS } = await import("../dist-smoke/replacement.mjs");

const res = await processReplacement(
  {
    inzibatci: fileLike(paths[0]),
    vacation: fileLike(paths[1]),
    filial: fileLike(paths[2]),
    resmi: fileLike(paths[3]),
    vezifeMaas: fileLike(paths[4]),
  },
  DEFAULT_PARAMS,
  (s, p) => console.log(`  [${p}%] ${s}`),
);

console.log("\n=== Yoxlamalar (bölmə 12) ===");
for (const c of res.checks) console.log(`${c.ok ? "PASS" : "FAIL"}${c.hard ? "" : " (info)"}  ${c.name}: ${c.detail}`);

console.log("\n=== Yekun ===");
for (const r of res.yekun) console.log(`${r.menbe} | sətir ${r.menbeSetri} | ${r.evezEdenKod} | ${r.dayCount} gün / ${r.hourCount} saat`);

let pass = true;
function check(name, cond) {
  console.log(`${cond ? "PASS" : "FAIL"}  ${name}`);
  if (!cond) pass = false;
}

console.log("\n=== Assertions ===");
check("bütün məcburi yoxlamalar keçdi", res.allPassed === true);
check("normallaşdırılan = 18", res.counters.normalized === 18);
check("təmiz data = 15", res.clean.length === 15);
check("Bəli = 12, Xeyr = 3", res.counters.beli === 12 && res.counters.xeyr === 3);
check("yekun = 9 sətir", res.yekun.length === 9);

const yekGun = res.yekun.filter((r) => r.dayBased).reduce((s, r) => s + r.dayCount, 0);
const yekSaat = res.yekun.filter((r) => !r.dayBased).reduce((s, r) => s + r.hourCount, 0);
check("yekun cəmi gün = 212", yekGun === 212);
check("yekun cəmi saat (saatlıq) = 4", yekSaat === 4);

check("fayl daxili duplikat = 1", res.counters.dupIntra === 1);
check("fayllar arası duplikat = 1", res.counters.dupCross === 1);
check("tarix məntiqsizliyi = 1", res.counters.dateLogicRemoved === 1);
check("üst-üstə düşmə = 2 sətir", res.counters.overlapRows === 2);
check("iç-içə = 2 sətir", res.counters.nestedRows === 2);
check("xərc konflikti = 1", res.counters.costConflicts === 1);
check("qırmızı sətir = 1", res.counters.redRows === 1);
check("birləşdirilən İnzibatçı sətri = 2", res.counters.mergedSourceRows === 2);

const ep = res.yekun.find((r) => r.menbe === "İnzibatçı prosesi" && r.evezEdenKod === "101");
check("101 epizodu birləşib (sətirlər 2, 3; 22 gün)", !!ep && ep.menbeSetri === "2, 3" && ep.dayCount === 22);
check("101 epizodunda 'Birləşdirilib' qeydi", !!ep && ep.qeydler.some((q) => q.startsWith("Birləşdirilib: davamlı gedən 2 mənbə sətri")));
check("101 lookup: Vəzifə-Maaş şəkilçisi", !!ep && ep.evezEdenVezife.includes("(Vəzifə-Maaş, Yanvar 2026)"));

const dupKept = res.yekun.find((r) => r.menbe === "İnzibatçı prosesi" && r.evezEdenKod === "110");
check("110 İnzibatçıda saxlanılıb və konservativ Bəli", !!dupKept && dupKept.xerc === true);
check("110 qeydində fayllar arası duplikat izahı", !!dupKept && dupKept.qeydler.some((q) => q.startsWith("Fayllar arası duplikat")));
check("110 Müvəqqəti keçid yekunda yoxdur", !res.yekun.some((r) => r.menbe === "Müvəqqəti keçid" && r.evezEdenKod === "110"));

const supheli = res.clean.find((r) => r.evezEdenKod === "210");
check("210 şübhəli bitmə → açıq, 30 gün", !!supheli && supheli.bitme === null && supheli.dayCount === 30);

const overlap = res.clean.filter((r) => r.evezEdenKod === "205");
check("205 hər iki sətir XƏTA flaqlı", overlap.length === 2 && overlap.every((r) => r.hasXeta));
check("205 yekunda yoxdur", !res.yekun.some((r) => r.evezEdenKod === "205"));

const nestedA = res.clean.find((r) => r.evezEdenKod === "301");
const nestedB = res.clean.find((r) => r.evezEdenKod === "303");
check("iç-içə qeydləri (301 və 303)",
  !!nestedA && nestedA.qeydler.some((q) => q.startsWith("İÇ-İÇƏ HAL: əvəz edən şəxs özü")) &&
  !!nestedB && nestedB.qeydler.some((q) => q.includes("İÇ-İÇƏ HAL")));
check("iç-içə sətirlər yekunda qalır", res.yekun.some((r) => r.evezEdenKod === "301") && res.yekun.some((r) => r.evezEdenKod === "303"));

const legv = res.clean.find((r) => r.menbe === "Müvəqqəti keçid" && r.evezEdenKod === "111");
check("111 DİQQƏT (ləğv əmri) qeydi", !!legv && legv.qeydler.some((q) => q.startsWith("DİQQƏT: bu keçid üzrə 'Müvəqqəti keçidin ləğvi'") && q.includes("əmr ƏM-77")));

const saatliq = res.yekun.find((r) => r.evezEdenKod === "401");
check("401 saatlıq: 4 saat / 0.5 gün + Acting qeydi", !!saatliq && saatliq.hourCount === 4 && saatliq.dayCount === 0.5 && saatliq.qeydler.some((q) => q === "Acting vəzifə: Şöbə müdiri"));

const acdiq = res.yekun.find((r) => r.evezEdenKod === "503");
check("503 açıq müddət (2999) → CAP-a qədər 122 gün", !!acdiq && acdiq.bitme === null && acdiq.dayCount === 122 && acdiq.qeydler.some((q) => q.startsWith("Bitmə tarixi 31.12.2999")));

const errTypes = new Set(res.errors.map((e) => e.novu));
for (const t of [
  "Filtrdən kənar (2026-dan əvvəl)",
  "Şübhəli bitmə tarixi",
  "Saat aralığı məntiqsizdir",
  "Tarix məntiqsizliyi (çıxarıldı)",
  "Üst-üstə düşən əvəzetmə",
  "İç-içə əvəzetmə",
  "Xərc qiymətləndirməsi konflikti",
  "Qırmızı işarələnmiş sətir",
  "Əmr növü: Müvəqqəti keçidin ləğvi",
]) {
  check(`Xəta Loqu: "${t}"`, errTypes.has(t));
}

const sheetNames = res.workbook.SheetNames;
for (const sn of ["Ümumi Baxış", "Xərc yaradan (yekun)", "Təmiz Data", "Əməkdaş üzrə cəm", "Xəta Loqu", "Duplikatlar", "Statistika", "BK_İnzibatçı prosesi", "BK_Rəsmi əvəzetmə (həvalə)", "BK_Vəzifə-Maaş"]) {
  check(`vərəq: ${sn}`, sheetNames.includes(sn));
}

// Yazılan faylın yenidən oxunması (round-trip)
const outPath = path.join(__dirname, "_out.xlsx");
XLSX.writeFile(res.workbook, outPath, { compression: true });
const reread = XLSX.read(readFileSync(outPath), { type: "buffer" });
const yk = reread.Sheets["Xərc yaradan (yekun)"];
const s5 = yk[XLSX.utils.encode_cell({ r: 4, c: 18 })];
check("yekun S5 xanasında düstur + dəyər var", !!s5 && !!s5.f && typeof s5.v === "number");

for (const p of [...paths, outPath]) unlinkSync(p);
console.log("\n", pass ? "ALL CHECKS PASS" : "SOME CHECKS FAILED");
process.exit(pass ? 0 : 1);
