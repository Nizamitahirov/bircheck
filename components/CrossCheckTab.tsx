"use client";

import { useCallback, useMemo, useState } from "react";
import {
  processCrossCheck,
  exportCrossCheck,
  formatSerial,
  type CrossResult,
  type CrossGroup,
  type Order,
} from "@/lib/crosscheck";
import { DownloadIcon, ErrorBox, FileDrop, PlayIcon, ProgressBar, Spinner, Stat } from "./ui";

type Status = "idle" | "processing" | "done" | "error";

export function CrossCheckTab() {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [progress, setProgress] = useState(0);
  const [stepText, setStepText] = useState("");
  const [result, setResult] = useState<CrossResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState("");

  const reset = () => {
    setFile(null);
    setStatus("idle");
    setProgress(0);
    setStepText("");
    setResult(null);
    setError(null);
    setSearch("");
  };

  const onPick = (f: File | null) => {
    setError(null);
    setResult(null);
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
    try {
      const res = await processCrossCheck(file, (step, pct) => {
        setStepText(step);
        setProgress(pct);
      });
      setResult(res);
      setStatus("done");
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : "Naməlum xəta baş verdi.";
      setError(msg);
      setStatus("error");
    }
  }, [file]);

  const fileBase = () => file?.name?.replace(/\.(xlsx|xlsm|xls)$/i, "") ?? "BirCheck";

  const visibleGroups = useMemo(() => {
    if (!result) return [] as CrossGroup[];
    if (!search.trim()) return result.groups;
    const q = search.toLowerCase().trim();
    return result.groups.filter((g) => g.code.toLowerCase().includes(q));
  }, [result, search]);

  const visiblePairCount = useMemo(
    () => visibleGroups.reduce((sum, g) => sum + g.pairs.length, 0),
    [visibleGroups],
  );

  const exportAll = () => {
    if (!result) return;
    exportCrossCheck(result, `${fileBase()}_crosscheck.xlsx`);
  };

  const exportSelected = () => {
    if (!result) return;
    const pairs = visibleGroups.flatMap((g) => g.pairs);
    exportCrossCheck(result, `${fileBase()}_crosscheck_secilmis.xlsx`, pairs);
  };

  return (
    <>
      <section className="glass mx-auto max-w-3xl rounded-3xl p-6 shadow-soft md:p-8">
        <FileDrop
          file={file}
          onPick={onPick}
          hint={
            <>
              Sheet:{" "}
              <code className="rounded bg-slate-100 px-1.5">Dates</code> (və ya ilk sheet) ·
              Sütunlar: A=Personal kod, B=Növ, C=Tarixdən, D=Tarixə
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
          <section className="mt-8 grid grid-cols-2 gap-3 md:grid-cols-3">
            <Stat label="Ümumi qeyd" value={result.totalRecords} />
            <Stat label="Kəsişən işçi" value={result.employeesWithCross} accent="brand" />
            <Stat label="Kəsişən cütlər" value={result.pairs.length} tone="rose" />
          </section>

          <section className="glass mt-6 rounded-3xl p-4 shadow-soft md:p-6">
            <div className="flex flex-col gap-4 border-b border-slate-200/70 pb-4 md:flex-row md:items-start md:justify-between">
              <div>
                <h2 className="text-lg font-bold text-slate-900">Kəsişən əmrlər</h2>
                <p className="mt-0.5 text-sm text-slate-500">
                  <span className="font-semibold text-slate-700">{visibleGroups.length}</span> unikal
                  işçi · <span className="font-semibold text-slate-700">{visiblePairCount}</span>{" "}
                  kəsişmə
                </p>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <button
                  onClick={exportSelected}
                  disabled={visiblePairCount === 0}
                  className="inline-flex items-center gap-2 rounded-xl border border-emerald-600 bg-white px-4 py-2 text-sm font-semibold text-emerald-700 transition hover:bg-emerald-50 disabled:cursor-not-allowed disabled:border-slate-200 disabled:text-slate-400"
                >
                  <DownloadIcon /> Seçilmişi endir ({visiblePairCount})
                </button>
                <button
                  onClick={exportAll}
                  disabled={result.pairs.length === 0}
                  className="inline-flex items-center gap-2 rounded-xl bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-soft transition hover:bg-emerald-700 disabled:cursor-not-allowed disabled:bg-slate-300"
                >
                  <DownloadIcon /> Hamısını endir ({result.pairs.length})
                </button>
              </div>
            </div>

            <div className="mt-4">
              <div className="relative w-full max-w-xs">
                <span className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400">
                  <SearchIcon />
                </span>
                <input
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Employee Badge axtar"
                  className="w-full rounded-lg border border-slate-200 bg-white py-2 pl-9 pr-3 text-sm shadow-sm focus:border-brand-400 focus:outline-none focus:ring-2 focus:ring-brand-200"
                />
              </div>
            </div>

            {result.pairs.length === 0 ? (
              <div className="mt-6 rounded-2xl border border-emerald-200 bg-emerald-50/60 p-8 text-center">
                <p className="text-lg font-semibold text-emerald-800">Kəsişən əmr tapılmadı</p>
                <p className="mt-1 text-sm text-emerald-700">
                  Bütün {result.totalRecords} qeyd üzrə heç bir işçidə tarixləri kəsişən əmr yoxdur.
                </p>
              </div>
            ) : visibleGroups.length === 0 ? (
              <div className="mt-6 rounded-2xl border border-slate-200 bg-white p-8 text-center text-sm text-slate-400">
                Bu axtarışa uyğun işçi yoxdur.
              </div>
            ) : (
              <div className="mt-4 space-y-4">
                {visibleGroups.map((g) => (
                  <GroupCard key={g.code} group={g} />
                ))}
              </div>
            )}
          </section>
        </>
      )}
    </>
  );
}

function GroupCard({ group }: { group: CrossGroup }) {
  return (
    <div className="overflow-hidden rounded-2xl border border-slate-200 bg-white">
      <div className="flex items-center justify-between gap-3 border-b border-slate-100 bg-slate-50 px-4 py-2.5">
        <div className="flex items-center gap-2">
          <span className="inline-flex h-7 w-7 items-center justify-center rounded-full bg-brand-100 text-xs font-bold text-brand-700">
            <UserIcon />
          </span>
          <span className="font-semibold text-slate-900">Badge {group.code}</span>
        </div>
        <span className="inline-flex items-center rounded-full bg-rose-100 px-2.5 py-0.5 text-xs font-semibold text-rose-800 ring-1 ring-inset ring-rose-200">
          {group.pairs.length} kəsişmə
        </span>
      </div>
      <div className="divide-y divide-slate-100">
        {group.pairs.map((p, i) => (
          <div
            key={i}
            className="flex flex-col items-stretch gap-2 px-4 py-3 md:flex-row md:items-center"
          >
            <OrderChip order={p.a} />
            <div className="flex flex-shrink-0 items-center justify-center text-rose-400 md:px-1">
              <CrossIcon />
            </div>
            <OrderChip order={p.b} />
          </div>
        ))}
      </div>
    </div>
  );
}

function OrderChip({ order }: { order: Order }) {
  return (
    <div className="flex flex-1 flex-wrap items-center gap-2 rounded-xl border border-slate-200 bg-slate-50/60 px-3 py-2">
      <span className="inline-flex items-center rounded-full bg-white px-2 py-0.5 text-xs font-medium text-slate-700 ring-1 ring-inset ring-slate-200">
        {order.type || "—"}
      </span>
      <span className="text-sm font-medium tabular-nums text-slate-800">
        {formatSerial(order.start)}
        <span className="mx-1.5 text-slate-400">→</span>
        {formatSerial(order.end)}
      </span>
    </div>
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

function UserIcon() {
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
      <path d="M20 21a8 8 0 0 0-16 0" />
      <circle cx="12" cy="7" r="4" />
    </svg>
  );
}

function CrossIcon() {
  return (
    <svg
      viewBox="0 0 24 24"
      className="h-5 w-5"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      <path d="M18 6 6 18" />
      <path d="m6 6 12 12" />
    </svg>
  );
}
