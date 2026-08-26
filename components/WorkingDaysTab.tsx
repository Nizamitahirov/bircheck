"use client";

import { useCallback, useMemo, useState } from "react";
import {
  processWorkbook,
  exportResult,
  formatExcelDate,
  type ProcessResult,
  type ProblemsRow,
} from "@/lib/processor";
import {
  DownloadIcon,
  ErrorBox,
  FileDrop,
  PlayIcon,
  ProgressBar,
  SegmentBtn,
  Spinner,
  Stat,
  formatNum,
} from "./ui";

type Status = "idle" | "processing" | "done" | "error";
type DiffMode = "range" | "all";
type SchedMode = "all" | "six" | "regular";

const RANGE_MIN = 1;
const RANGE_MAX = 31;

export function WorkingDaysTab() {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [progress, setProgress] = useState(0);
  const [stepText, setStepText] = useState("");
  const [result, setResult] = useState<ProcessResult | null>(null);
  const [rows, setRows] = useState<ProblemsRow[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [diffMode, setDiffMode] = useState<DiffMode>("range");
  const [schedMode, setSchedMode] = useState<SchedMode>("all");
  const [search, setSearch] = useState("");

  const reset = () => {
    setFile(null);
    setStatus("idle");
    setProgress(0);
    setStepText("");
    setResult(null);
    setRows([]);
    setError(null);
    setSearch("");
    setDiffMode("range");
    setSchedMode("all");
  };

  const onPick = (f: File | null) => {
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

  const fileBase = () => file?.name?.replace(/\.(xlsx|xlsm|xls)$/i, "") ?? "BirCheck";

  const exportAll = () => {
    if (!result) return;
    exportResult(result, rows, `${fileBase()}_yoxlanilmis.xlsx`);
  };

  const exportSelected = () => {
    if (!result) return;
    exportResult(result, visibleRows, `${fileBase()}_secilmis.xlsx`);
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
    if (diffMode === "range") {
      list = list.filter((r) => {
        const d = Math.abs(Number(r.cells[11]) || 0);
        return d >= RANGE_MIN && d <= RANGE_MAX;
      });
    }
    if (schedMode === "six") list = list.filter((r) => r.is6Day);
    else if (schedMode === "regular") list = list.filter((r) => !r.is6Day);
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
  }, [rows, diffMode, schedMode, search]);

  const stats = useMemo(() => {
    let pos = 0;
    let neg = 0;
    let six = 0;
    for (const r of rows) {
      const d = Number(r.cells[11]) || 0;
      if (d > 0) pos++;
      else if (d < 0) neg++;
      if (r.is6Day) six++;
    }
    return { pos, neg, six };
  }, [rows]);

  return (
    <>
      <section className="glass mx-auto max-w-3xl rounded-3xl p-6 shadow-soft md:p-8">
        <FileDrop
          file={file}
          onPick={onPick}
          hint={
            <>
              Tələb olunan sheet-lər:{" "}
              <code className="rounded bg-slate-100 px-1.5">From_Excel_Baza</code>,{" "}
              <code className="rounded bg-slate-100 px-1.5">From_MHMD</code>,{" "}
              <code className="rounded bg-slate-100 px-1.5">From_HRB_Otpusk</code>,{" "}
              <code className="rounded bg-slate-100 px-1.5">Problems</code>
            </>
          }
        />

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
            <ProgressBar pct={progress} label={stepText} />
          </div>
        )}

        {error && (
          <div className="mt-6">
            <ErrorBox message={error} />
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
            {/* Header: title + export actions */}
            <div className="flex flex-col gap-4 border-b border-slate-200/70 pb-4 md:flex-row md:items-start md:justify-between">
              <div>
                <h2 className="text-lg font-bold text-slate-900">Fərqlər</h2>
                <p className="mt-0.5 text-sm text-slate-500">
                  <span className="font-semibold text-slate-700">{visibleRows.length}</span> /{" "}
                  {result.totalEmployees} işçi göstərilir
                  {stats.six > 0 && (
                    <span className="ml-2 inline-flex items-center gap-1 rounded-full bg-indigo-50 px-2 py-0.5 text-xs font-medium text-indigo-700 ring-1 ring-inset ring-indigo-200">
                      6 günlük: {stats.six}
                    </span>
                  )}
                </p>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <button
                  onClick={exportSelected}
                  disabled={visibleRows.length === 0}
                  className="inline-flex items-center gap-2 rounded-xl border border-emerald-600 bg-white px-4 py-2 text-sm font-semibold text-emerald-700 transition hover:bg-emerald-50 disabled:cursor-not-allowed disabled:border-slate-200 disabled:text-slate-400"
                >
                  <DownloadIcon /> Seçilmişi endir ({visibleRows.length})
                </button>
                <button
                  onClick={exportAll}
                  className="inline-flex items-center gap-2 rounded-xl bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-soft transition hover:bg-emerald-700"
                >
                  <DownloadIcon /> Hamısını endir ({result.totalEmployees})
                </button>
              </div>
            </div>

            {/* Filter toolbar */}
            <div className="mt-4 flex flex-wrap items-end gap-x-5 gap-y-3">
              <label className="flex flex-col gap-1 text-xs font-medium text-slate-500">
                Axtar
                <div className="relative">
                  <span className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400">
                    <SearchIcon />
                  </span>
                  <input
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                    placeholder="Kod, ad, soyad"
                    className="w-60 rounded-lg border border-slate-200 bg-white py-1.5 pl-9 pr-3 text-sm shadow-sm focus:border-brand-400 focus:outline-none focus:ring-2 focus:ring-brand-200"
                  />
                </div>
              </label>

              <label className="flex flex-col gap-1 text-xs font-medium text-slate-500">
                Fərq
                <div className="inline-flex rounded-lg border border-slate-200 bg-white p-1 shadow-sm">
                  <SegmentBtn active={diffMode === "range"} onClick={() => setDiffMode("range")}>
                    1–31 aralığı
                  </SegmentBtn>
                  <SegmentBtn active={diffMode === "all"} onClick={() => setDiffMode("all")}>
                    Hamısı
                  </SegmentBtn>
                </div>
              </label>

              <label className="flex flex-col gap-1 text-xs font-medium text-slate-500">
                İş rejimi
                <div className="inline-flex rounded-lg border border-slate-200 bg-white p-1 shadow-sm">
                  <SegmentBtn active={schedMode === "all"} onClick={() => setSchedMode("all")}>
                    Hamısı
                  </SegmentBtn>
                  <SegmentBtn active={schedMode === "six"} onClick={() => setSchedMode("six")}>
                    6 günlük
                  </SegmentBtn>
                  <SegmentBtn active={schedMode === "regular"} onClick={() => setSchedMode("regular")}>
                    Digər
                  </SegmentBtn>
                </div>
              </label>
            </div>

            <div className="mt-4 overflow-x-auto rounded-2xl border border-slate-200 bg-white">
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
    </>
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
        <div className="flex items-center gap-2">
          <span className="font-medium">
            {surname} {name}
          </span>
          {row.is6Day && (
            <span className="inline-flex flex-shrink-0 items-center rounded-full bg-indigo-100 px-1.5 py-0.5 text-[10px] font-semibold uppercase tracking-wide text-indigo-700 ring-1 ring-inset ring-indigo-200">
              6 günlük
            </span>
          )}
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

function SearchIcon() {
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
      <circle cx="11" cy="11" r="7" />
      <path d="m21 21-4.3-4.3" />
    </svg>
  );
}
