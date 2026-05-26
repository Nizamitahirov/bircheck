"use client";

import { useCallback, useMemo, useRef, useState } from "react";
import {
  processWorkbook,
  exportResult,
  formatExcelDate,
  type ProcessResult,
  type ProblemsRow,
} from "@/lib/processor";

type Status = "idle" | "processing" | "done" | "error";
type FilterMode = "range" | "all";

const RANGE_MIN = 1;
const RANGE_MAX = 31;

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [progress, setProgress] = useState(0);
  const [stepText, setStepText] = useState("");
  const [result, setResult] = useState<ProcessResult | null>(null);
  const [rows, setRows] = useState<ProblemsRow[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [dragging, setDragging] = useState(false);
  const [filterMode, setFilterMode] = useState<FilterMode>("range");
  const [search, setSearch] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  const reset = () => {
    setFile(null);
    setStatus("idle");
    setProgress(0);
    setStepText("");
    setResult(null);
    setRows([]);
    setError(null);
    setSearch("");
    if (inputRef.current) inputRef.current.value = "";
  };

  const onPick = (f: File | null) => {
    if (!f) return;
    if (!/\.(xlsx|xlsm|xls)$/i.test(f.name)) {
      setError("Yalnız .xlsx, .xlsm və ya .xls faylları yüklənə bilər.");
      return;
    }
    setError(null);
    setResult(null);
    setRows([]);
    setStatus("idle");
    setProgress(0);
    setFile(f);
  };

  const run = useCallback(async () => {
    if (!file) return;
    setStatus("processing");
    setProgress(0);
    setError(null);
    setResult(null);
    setRows([]);
    try {
      const res = await processWorkbook(file, (step, pct) => {
        setStepText(step);
        setProgress(pct);
      });
      setResult(res);
      setRows(res.rows);
      setStatus("done");
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : "Naməlum xəta baş verdi.";
      setError(msg);
      setStatus("error");
    }
  }, [file]);

  const download = () => {
    if (!result) return;
    const base = file?.name?.replace(/\.(xlsx|xlsm|xls)$/i, "") ?? "BirCheck";
    exportResult(result, rows, `${base}_yoxlanilmis.xlsx`);
  };

  const updateReason = useCallback((id: string, value: string) => {
    setRows((prev) =>
      prev.map((r) =>
        r.id === id ? { ...r, cells: r.cells.map((c, i) => (i === 12 ? value : c)) } : r,
      ),
    );
  }, []);

  const visibleRows = useMemo(() => {
    let list = rows;
    if (filterMode === "range") {
      list = list.filter((r) => {
        const d = Math.abs(Number(r.cells[11]) || 0);
        return d >= RANGE_MIN && d <= RANGE_MAX;
      });
    }
    if (search.trim()) {
      const q = search.toLowerCase().trim();
      list = list.filter((r) => {
        const a = String(r.cells[0] ?? "").toLowerCase();
        const b = String(r.cells[1] ?? "").toLowerCase();
        const c = String(r.cells[2] ?? "").toLowerCase();
        return a.includes(q) || b.includes(q) || c.includes(q);
      });
    }
    return list;
  }, [rows, filterMode, search]);

  const stats = useMemo(() => {
    let pos = 0;
    let neg = 0;
    for (const r of rows) {
      const d = Number(r.cells[11]) || 0;
      if (d > 0) pos++;
      else if (d < 0) neg++;
    }
    return { pos, neg };
  }, [rows]);

  const fileSize = useMemo(() => (file ? formatBytes(file.size) : ""), [file]);

  return (
    <main className="min-h-screen w-full px-4 py-10 md:py-16">
      <div className="mx-auto w-full max-w-6xl">
        <header className="mb-10 text-center">
          <h1 className="text-4xl font-extrabold tracking-tight text-slate-900 md:text-5xl">
            Bir<span className="text-brand-600">Check</span>
          </h1>
        </header>

        <section className="glass mx-auto max-w-3xl rounded-3xl p-6 shadow-soft md:p-8">
          <label
            onDragOver={(e) => {
              e.preventDefault();
              setDragging(true);
            }}
            onDragLeave={() => setDragging(false)}
            onDrop={(e) => {
              e.preventDefault();
              setDragging(false);
              onPick(e.dataTransfer.files?.[0] ?? null);
            }}
            className={`group relative flex cursor-pointer flex-col items-center justify-center rounded-2xl border-2 border-dashed px-6 py-10 text-center transition ${
              dragging
                ? "border-brand-500 bg-brand-50/60"
                : "border-slate-300 bg-white/40 hover:border-brand-400 hover:bg-white/70"
            }`}
          >
            <input
              ref={inputRef}
              type="file"
              accept=".xlsx,.xlsm,.xls"
              className="hidden"
              onChange={(e) => onPick(e.target.files?.[0] ?? null)}
            />
            <div className="mb-3 inline-flex h-14 w-14 items-center justify-center rounded-2xl bg-brand-100 text-brand-600 transition group-hover:scale-105">
              <UploadIcon />
            </div>
            {file ? (
              <div className="space-y-1">
                <p className="font-medium text-slate-900">{file.name}</p>
                <p className="text-sm text-slate-500">{fileSize}</p>
                <p className="mt-2 text-xs text-brand-600">Başqa fayl seçmək üçün klikləyin</p>
              </div>
            ) : (
              <div className="space-y-1">
                <p className="font-semibold text-slate-800">
                  Excel faylını bura sürüşdürün və ya seçin
                </p>
                <p className="text-sm text-slate-500">
                  Tələb olunan sheet-lər:{" "}
                  <code className="rounded bg-slate-100 px-1.5">From_Excel_Baza</code>,{" "}
                  <code className="rounded bg-slate-100 px-1.5">From_MHMD</code>,{" "}
                  <code className="rounded bg-slate-100 px-1.5">From_HRB_Otpusk</code>,{" "}
                  <code className="rounded bg-slate-100 px-1.5">Problems</code>
                </p>
              </div>
            )}
          </label>

          <div className="mt-6 flex flex-col gap-3 sm:flex-row">
            <button
              onClick={run}
              disabled={!file || status === "processing"}
              className="inline-flex flex-1 items-center justify-center gap-2 rounded-xl bg-brand-600 px-5 py-3 font-semibold text-white shadow-soft transition hover:bg-brand-700 disabled:cursor-not-allowed disabled:bg-slate-300 disabled:shadow-none"
            >
              {status === "processing" ? (
                <>
                  <Spinner /> İşlənir...
                </>
              ) : (
                <>
                  <PlayIcon /> Yoxlamanı başlat
                </>
              )}
            </button>
            <button
              onClick={reset}
              disabled={status === "processing"}
              className="inline-flex items-center justify-center gap-2 rounded-xl border border-slate-200 bg-white px-5 py-3 font-medium text-slate-700 transition hover:bg-slate-50 disabled:opacity-50"
            >
              Təmizlə
            </button>
          </div>

          {(status === "processing" || status === "done") && (
            <div className="mt-6">
              <div className="mb-2 flex items-center justify-between text-sm">
                <span className="font-medium text-slate-700">{stepText || "Hazır"}</span>
                <span className="tabular-nums text-slate-500">{progress}%</span>
              </div>
              <div className="h-2 w-full overflow-hidden rounded-full bg-slate-200">
                <div
                  className="h-full rounded-full bg-gradient-to-r from-brand-400 to-brand-600 transition-[width] duration-300"
                  style={{ width: `${progress}%` }}
                />
              </div>
            </div>
          )}

          {error && (
            <div className="mt-6 rounded-xl border border-red-200 bg-red-50 p-4 text-sm text-red-800">
              <p className="font-semibold">Xəta</p>
              <p className="mt-1">{error}</p>
            </div>
          )}
        </section>

        {result && status === "done" && (
          <>
            <section className="mt-8 grid grid-cols-2 gap-3 md:grid-cols-4">
              <Stat label="İşçi sayı" value={result.totalEmployees} />
              <Stat label="Fərq tapıldı" value={result.diffCount} accent="brand" />
              <Stat label="Müsbət fərq" value={stats.pos} tone="amber" />
              <Stat label="Mənfi fərq" value={stats.neg} tone="rose" />
            </section>

            <section className="glass mt-6 rounded-3xl p-4 shadow-soft md:p-6">
              <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                <div>
                  <h2 className="text-lg font-bold text-slate-900">Fərqlər</h2>
                  <p className="text-sm text-slate-500">
                    {filterMode === "range" ? (
                      <>
                        |Fərq| ∈ [{RANGE_MIN}; {RANGE_MAX}] aralığında ·{" "}
                        <span className="font-semibold text-slate-700">
                          {visibleRows.length}
                        </span>{" "}
                        sətir göstərilir
                      </>
                    ) : (
                      <>
                        Bütün işçilər ·{" "}
                        <span className="font-semibold text-slate-700">
                          {visibleRows.length}
                        </span>{" "}
                        sətir göstərilir
                      </>
                    )}
                    <span className="ml-2 inline-flex items-center gap-1 rounded-full bg-slate-100 px-2 py-0.5 text-xs text-slate-600">
                      Endirilən fayl: bütün {result.totalEmployees} işçi
                    </span>
                  </p>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <input
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                    placeholder="Axtar (kod, ad, soyad)"
                    className="w-56 rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm focus:border-brand-400 focus:outline-none focus:ring-2 focus:ring-brand-200"
                  />
                  <div className="inline-flex rounded-xl border border-slate-200 bg-white p-1 text-sm shadow-sm">
                    <SegmentBtn
                      active={filterMode === "range"}
                      onClick={() => setFilterMode("range")}
                    >
                      1–31 aralığı
                    </SegmentBtn>
                    <SegmentBtn
                      active={filterMode === "all"}
                      onClick={() => setFilterMode("all")}
                    >
                      Hamısı
                    </SegmentBtn>
                  </div>
                  <button
                    onClick={download}
                    className="inline-flex items-center gap-2 rounded-xl bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-soft transition hover:bg-emerald-700"
                  >
                    <DownloadIcon /> Hamısını endir (.xlsx)
                  </button>
                </div>
              </div>

              <div className="overflow-x-auto rounded-2xl border border-slate-200 bg-white">
                <table className="min-w-full divide-y divide-slate-200 text-sm">
                  <thead className="bg-slate-50">
                    <tr className="text-left text-xs font-semibold uppercase tracking-wider text-slate-500">
                      <th className="px-3 py-3">#</th>
                      <th className="px-3 py-3">Kod</th>
                      <th className="px-3 py-3">Ad Soyad</th>
                      <th className="px-3 py-3">Vəzifə</th>
                      <th className="px-3 py-3">Filial</th>
                      <th className="px-3 py-3 whitespace-nowrap">İşə qəbul</th>
                      <th className="px-3 py-3 text-right">MHMD</th>
                      <th className="px-3 py-3 text-right">Bizdə</th>
                      <th className="px-3 py-3 text-right">Fərq</th>
                      <th className="px-3 py-3 min-w-[260px]">Fərqin səbəbi</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {visibleRows.length === 0 ? (
                      <tr>
                        <td colSpan={10} className="px-3 py-10 text-center text-sm text-slate-400">
                          Bu filtrə uyğun sətir yoxdur.
                        </td>
                      </tr>
                    ) : (
                      visibleRows.map((r, i) => (
                        <RowItem
                          key={r.id}
                          index={i + 1}
                          row={r}
                          onReason={(v) => updateReason(r.id, v)}
                        />
                      ))
                    )}
                  </tbody>
                </table>
              </div>

              <p className="mt-3 text-xs text-slate-500">
                İpucu: <span className="font-medium">Fərqin səbəbi</span> sütununa istənilən qeyd
                əlavə edə bilərsiniz. Düzəlişlər endirilən fayla daxil olur.
              </p>
            </section>
          </>
        )}

        <footer className="mt-12 text-center text-xs text-slate-400">
          Copyright created by Nizami Tahir
        </footer>
      </div>
    </main>
  );
}

function RowItem({
  index,
  row,
  onReason,
}: {
  index: number;
  row: ProblemsRow;
  onReason: (v: string) => void;
}) {
  const c = row.cells;
  const diff = Number(c[11]) || 0;
  const diffClass =
    diff > 0
      ? "bg-amber-100 text-amber-800 ring-amber-200"
      : diff < 0
        ? "bg-rose-100 text-rose-800 ring-rose-200"
        : "bg-slate-100 text-slate-700 ring-slate-200";
  const surname = String(c[1] ?? "").trim();
  const name = String(c[2] ?? "").trim();

  return (
    <tr className="transition hover:bg-slate-50/60">
      <td className="px-3 py-2 tabular-nums text-slate-400">{index}</td>
      <td className="px-3 py-2 font-medium tabular-nums text-slate-700">{String(c[0] ?? "")}</td>
      <td className="px-3 py-2 text-slate-900">
        <div className="font-medium">
          {surname} {name}
        </div>
      </td>
      <td className="px-3 py-2 text-slate-600">{String(c[4] ?? "")}</td>
      <td className="px-3 py-2 text-slate-600">{String(c[6] ?? "")}</td>
      <td className="px-3 py-2 whitespace-nowrap tabular-nums text-slate-600">
        {formatExcelDate(c[7])}
      </td>
      <td className="px-3 py-2 text-right tabular-nums text-slate-700">{formatNum(c[9])}</td>
      <td className="px-3 py-2 text-right tabular-nums text-slate-700">{formatNum(c[10])}</td>
      <td className="px-3 py-2 text-right">
        <span
          className={`inline-flex min-w-[2.5rem] justify-center rounded-full px-2 py-0.5 text-xs font-semibold ring-1 ring-inset tabular-nums ${diffClass}`}
        >
          {diff > 0 ? `+${diff}` : diff}
        </span>
      </td>
      <td className="px-3 py-2">
        <textarea
          value={String(c[12] ?? "")}
          onChange={(e) => onReason(e.target.value)}
          rows={1}
          className="w-full resize-y rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-xs leading-relaxed shadow-sm transition focus:border-brand-400 focus:outline-none focus:ring-2 focus:ring-brand-200"
          placeholder="Səbəbi daxil edin..."
        />
      </td>
    </tr>
  );
}

function Stat({
  label,
  value,
  accent,
  tone,
}: {
  label: string;
  value: number;
  accent?: "brand";
  tone?: "amber" | "rose";
}) {
  let color = "text-slate-900";
  if (accent === "brand") color = "text-brand-700";
  else if (tone === "amber") color = "text-amber-700";
  else if (tone === "rose") color = "text-rose-700";
  return (
    <div className="glass rounded-2xl p-4 text-center shadow-sm">
      <div className={`text-3xl font-bold tabular-nums ${color}`}>{value}</div>
      <div className="mt-1 text-xs uppercase tracking-wide text-slate-500">{label}</div>
    </div>
  );
}

function SegmentBtn({
  active,
  onClick,
  children,
}: {
  active: boolean;
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`rounded-lg px-3 py-1.5 text-sm font-medium transition ${
        active ? "bg-brand-600 text-white shadow-sm" : "text-slate-600 hover:bg-slate-50"
      }`}
    >
      {children}
    </button>
  );
}

function Spinner() {
  return (
    <svg className="h-4 w-4 animate-spin" viewBox="0 0 24 24" fill="none">
      <circle cx="12" cy="12" r="10" stroke="currentColor" strokeOpacity="0.25" strokeWidth="4" />
      <path
        d="M22 12a10 10 0 0 0-10-10"
        stroke="currentColor"
        strokeWidth="4"
        strokeLinecap="round"
      />
    </svg>
  );
}

function PlayIcon() {
  return (
    <svg viewBox="0 0 24 24" className="h-4 w-4" fill="currentColor">
      <path d="M8 5v14l11-7z" />
    </svg>
  );
}

function UploadIcon() {
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
      className="h-7 w-7"
    >
      <path d="M12 16V4" />
      <path d="m6 10 6-6 6 6" />
      <path d="M4 20h16" />
    </svg>
  );
}

function DownloadIcon() {
  return (
    <svg
      viewBox="0 0 24 24"
      className="h-4 w-4"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      <path d="M12 4v12" />
      <path d="m6 10 6 6 6-6" />
      <path d="M4 20h16" />
    </svg>
  );
}

function formatBytes(n: number) {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(2)} MB`;
}

function formatNum(v: unknown): string {
  if (v === null || v === undefined || v === "") return "";
  const n = Number(v);
  if (!Number.isFinite(n)) return String(v);
  return Number.isInteger(n) ? String(n) : n.toFixed(2);
}
