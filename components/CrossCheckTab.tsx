"use client";

import { useCallback, useMemo, useState } from "react";
import {
  processCrossCheck,
  exportCrossCheck,
  formatSerial,
  type CrossResult,
  type CrossGroup,
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
    exportCrossCheck(result, `${fileBase()}_crosscheck_secilmis.xlsx`, visibleGroups);
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
              <div className="mt-4 overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-sm">
                <div className="overflow-x-auto">
                  <table className="min-w-full border-separate border-spacing-0 text-sm">
                    <thead>
                      <tr className="bg-slate-50 text-left text-[11px] font-semibold uppercase tracking-wider text-slate-500">
                        <th className="sticky left-0 z-10 border-b border-slate-200 bg-slate-50 px-4 py-3">
                          Emp Badge
                        </th>
                        <th className="border-b border-slate-200 px-4 py-3">Kəsişən tarixlər</th>
                        <th className="border-b border-slate-200 px-4 py-3 text-center">
                          Kəsişən gün sayı
                        </th>
                        <th className="border-b border-slate-200 px-4 py-3">Əmrin növü</th>
                        <th className="border-b border-slate-200 px-4 py-3 text-center">Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      {visibleGroups.map((g, gi) => (
                        <GroupRow key={g.code} group={g} zebra={gi % 2 === 1} />
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </section>
        </>
      )}
    </>
  );
}

const ROW_H = "min-h-[2rem]";

function GroupRow({ group, zebra }: { group: CrossGroup; zebra: boolean }) {
  const base = zebra ? "bg-slate-50/40" : "bg-white";
  return (
    <tr className={`align-top transition hover:bg-brand-50/40 ${base}`}>
      <td
        className={`sticky left-0 z-10 whitespace-nowrap border-b border-slate-100 px-4 py-3 ${base}`}
      >
        <div className="flex items-center gap-2.5">
          <span className="inline-flex h-9 w-9 items-center justify-center rounded-xl bg-gradient-to-br from-brand-500 to-brand-700 text-white shadow-sm">
            <UserIcon />
          </span>
          <span className="text-base font-bold tabular-nums text-slate-900">{group.code}</span>
        </div>
      </td>

      <td className="border-b border-slate-100 px-4 py-3">
        <div className="flex flex-col gap-1.5">
          {group.orders.map((o, i) => (
            <span
              key={i}
              className={`inline-flex w-fit items-center gap-1.5 rounded-lg bg-slate-100/70 px-2.5 text-sm font-medium tabular-nums text-slate-800 ring-1 ring-inset ring-slate-200 ${ROW_H}`}
            >
              <CalendarIcon />
              {formatSerial(o.start)}
              <span className="text-slate-400">→</span>
              {formatSerial(o.end)}
            </span>
          ))}
        </div>
      </td>

      <td className="border-b border-slate-100 px-4 py-3">
        <div className="flex flex-col items-center gap-1.5">
          {group.orders.map((o, i) => (
            <span
              key={i}
              className={`inline-flex items-center justify-center rounded-full px-2.5 text-sm font-semibold tabular-nums ${dayTone(
                o.crossingDays,
              )} ${ROW_H} min-w-[2.5rem]`}
            >
              {o.crossingDays}
            </span>
          ))}
        </div>
      </td>

      <td className="border-b border-slate-100 px-4 py-3">
        <div className="flex flex-col gap-1.5">
          {group.orders.map((o, i) => {
            const t = typeTone(o.type);
            return (
              <span
                key={i}
                className={`inline-flex w-fit items-center gap-1.5 rounded-full px-2.5 text-sm font-medium ring-1 ring-inset ${t.cls} ${ROW_H}`}
              >
                <span className={`h-1.5 w-1.5 rounded-full ${t.dot}`} />
                {o.type || "—"}
              </span>
            );
          })}
        </div>
      </td>

      <td className="border-b border-slate-100 px-4 py-3 text-center">
        <span className="inline-flex items-center gap-1 rounded-full bg-rose-100 px-3 py-1 text-xs font-bold text-rose-800 ring-1 ring-inset ring-rose-200">
          {group.pairs.length} Kəsişmə
        </span>
      </td>
    </tr>
  );
}

function dayTone(days: number): string {
  if (days >= 5) return "bg-rose-100 text-rose-800 ring-1 ring-inset ring-rose-200";
  if (days >= 2) return "bg-amber-100 text-amber-800 ring-1 ring-inset ring-amber-200";
  return "bg-emerald-100 text-emerald-800 ring-1 ring-inset ring-emerald-200";
}

// Category coloring by keyword (works for both clean and lightly-garbled AZ text).
function typeTone(type: string): { cls: string; dot: string } {
  const t = (type || "").toLowerCase();
  if (/(m.?zuniyy|mezuniyy|məzuniyy)/.test(t))
    return { cls: "bg-indigo-50 text-indigo-700 ring-indigo-200", dot: "bg-indigo-500" };
  if (/ezam/.test(t))
    return { cls: "bg-amber-50 text-amber-700 ring-amber-200", dot: "bg-amber-500" };
  if (/overtime|bo\b/.test(t))
    return { cls: "bg-emerald-50 text-emerald-700 ring-emerald-200", dot: "bg-emerald-500" };
  if (/x.?st.?lik|xestelik|xəstəlik/.test(t))
    return { cls: "bg-rose-50 text-rose-700 ring-rose-200", dot: "bg-rose-500" };
  return { cls: "bg-slate-50 text-slate-700 ring-slate-200", dot: "bg-slate-400" };
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

function CalendarIcon() {
  return (
    <svg
      viewBox="0 0 24 24"
      className="h-3.5 w-3.5 text-slate-400"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      <rect x="3" y="4" width="18" height="18" rx="2" />
      <path d="M16 2v4M8 2v4M3 10h18" />
    </svg>
  );
}

