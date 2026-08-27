import * as XLSX from "xlsx";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { readFileSync, unlinkSync } from "node:fs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function build(rows, name = "Dates") {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name);
  const tmp = path.join(__dirname, "_cross.xlsx");
  XLSX.writeFile(wb, tmp);
  const buf = readFileSync(tmp);
  return {
    tmp,
    fileLike: {
      arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
      name: "cross.xlsx",
    },
  };
}

if (typeof performance === "undefined") globalThis.performance = { now: () => Date.now() };
const { processCrossCheck } = await import("../dist-smoke/crosscheck.mjs");

let pass = true;
function check(name, cond) {
  console.log(`${cond ? "PASS" : "FAIL"}  ${name}`);
  if (!cond) pass = false;
}

// --- Test 1: the user's exact sample → expected 0 crossings ---
const header = ["Personal kod", "Növ", "Tarixdən", "Tarixə"];
const sample = [
  header,
  [10002, "~m'k m'zuniyy't", "05.06.2026", "14.06.2026"],
  [10010, "~m'k m'zuniyy't", "03.07.2026", "05.07.2026"],
  [10010, "~m'k m'zuniyy't", "10.07.2026", "12.07.2026"],
  [10015, "~m'k m'zuniyy't", "10.07.2026", "12.07.2026"],
  [10015, "~m'k m'zuniyy't", "17.07.2026", "19.07.2026"],
  [10015, "~m'k m'zuniyy't", "24.07.2026", "26.07.2026"],
  [10015, "~m'k m'zuniyy't", "31.07.2026", "02.08.2026"],
  [10020, "~m'k m'zuniyy't", "05.06.2026", "06.06.2026"],
  [10020, "~m'k m'zuniyy't", "07.06.2026", "14.06.2026"],
  [10037, "~m'k m'zuniyy't", "27.07.2026", "16.08.2026"],
  [10037, "~m'k m'zuniyy't", "17.08.2026", "17.08.2026"],
  [10037, "~m'k m'zuniyy't", "18.08.2026", "23.08.2026"],
  [10059, "~m'k m'zuniyy't", "17.06.2026", "30.06.2026"],
  [10063, "~m'k m'zuniyy't", "12.06.2026", "14.06.2026"],
  [10063, "overtime BO", "28.06.2026", "28.06.2026"],
  [10063, "~m'k m'zuniyy't", "21.07.2026", "21.07.2026"],
  [10077, "~m'k m'zuniyy't", "18.06.2026", "21.06.2026"],
  [10095, "ezamiyyə", "02.06.2026", "05.06.2026"],
  [10095, "ezamiyyə", "23.06.2026", "25.06.2026"],
];
{
  const { tmp, fileLike } = build(sample);
  const res = await processCrossCheck(fileLike);
  console.log(`\nTest 1 (sample): records=${res.totalRecords}, crossings=${res.pairs.length}`);
  check("sample: 19 records read", res.totalRecords === 19);
  check("sample: 0 crossings", res.pairs.length === 0);
  unlinkSync(tmp);
}

// --- Test 2: synthetic overlaps ---
const t2 = [
  header,
  // 200: two overlapping (05-10 and 08-12), plus a third contained (06-07) → pairs: (a,b),(a,c),(b?)
  [200, "Əmək", "05.06.2026", "10.06.2026"],
  [200, "Xəstəlik", "08.06.2026", "12.06.2026"],
  [200, "Ezamiyyə", "06.06.2026", "07.06.2026"],
  // 201: touching boundary (ends 10th, next starts 10th) → 1 crossing
  [201, "Əmək", "01.06.2026", "10.06.2026"],
  [201, "Xəstəlik", "10.06.2026", "15.06.2026"],
  // 202: non-overlapping → 0
  [202, "Əmək", "01.06.2026", "05.06.2026"],
  [202, "Xəstəlik", "10.06.2026", "15.06.2026"],
];
{
  const { tmp, fileLike } = build(t2);
  const res = await processCrossCheck(fileLike);
  console.log(`\nTest 2 (synthetic): crossings=${res.pairs.length}, employees=${res.employeesWithCross}`);
  for (const g of res.groups) console.log(`  ${g.code}: ${g.pairs.length} pairs`);

  // 200: orders sorted by start: (05-10),(06-07),(08-12)
  //   (05-10) vs (06-07): overlap → yes
  //   (05-10) vs (08-12): 08<=10 → yes
  //   (06-07) vs (08-12): 08<=07? no
  //   → 2 pairs
  const g200 = res.groups.find((g) => g.code === "200");
  const g201 = res.groups.find((g) => g.code === "201");
  const g202 = res.groups.find((g) => g.code === "202");
  check("200 has 2 crossings", g200?.pairs.length === 2);
  check("201 has 1 crossing (touching boundary)", g201?.pairs.length === 1);
  check("202 not in results (no crossing)", !g202);
  check("employeesWithCross = 2", res.employeesWithCross === 2);
  check("total pairs = 3", res.pairs.length === 3);
  // 200 has 3 distinct orders all involved in crossings
  check("200 distinct orders = 3", g200?.orders.length === 3);
  check("201 distinct orders = 2", g201?.orders.length === 2);
  // orders sorted by start
  check("200 orders sorted by start", g200 && g200.orders[0].start <= g200.orders[1].start);
  unlinkSync(tmp);
}

console.log("\n", pass ? "ALL CHECKS PASS" : "SOME CHECKS FAILED");
process.exit(pass ? 0 : 1);
