"use client";

import { useCallback, useMemo, useState } from "react";
import {
  DEFAULT_WORKYEAR_PARAMS,
  exportWorkyear,
  processWorkyear,
  todaySerial,
  type WorkyearParams,
  type WorkyearResult,
} from "@/lib/vacation-workyear";
import { fmtDate } from "@/lib/replacement/xl";
import { DownloadIcon, ErrorBox, FileDrop, PlayIcon, ProgressBar, SegmentBtn, Spinner, Stat } from "./ui";

type Status = "idle" | "processing" | "done" | "error";
type Preview = "checks" | "rows" | "logs";

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;

function serialToISO(s: number): string {
  return new Date(EXCEL_EPOCH_MS + s * MS_PER_DAY).toISOString().slice(0, 10);
}

function isoToSerial(iso: string): number | null {
  const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return Math.round((Date.UTC(+m[1], +m[2] - 1, +m[3]) - EXCEL_EPOCH_MS) / MS_PER_DAY);
}

export function VacationWorkyearTab() {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [progress, setProgress] = useState(0);
  const [stepText, setStepText] = useState("");
  const [result, setResult] = useState<WorkyearResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [preview, setPreview] = useState<Preview>("checks");
  const [search, setSearch] = useState("");

  const [asOf, setAsOf] = useState(serialToISO(todaySerial()));
  const [prorate, setProrate] = useState(DEFAULT_WORKYEAR_PARAMS.prorateLastPeriod);
  const [includeComp, setIncludeComp] = useState(DEFAULT_WORKYEAR_PARAMS.includeCompensation);
  const [masterSheet, setMasterSheet] = useState(DEFAULT_WORKYEAR_PARAMS.masterSheet);

  const reset = () => {
    setFile(null);
    setStatus("idle");
    setProgress(0);
    setStepText("");
    setResult(null);
    setError(null);
    setPreview("checks");
    setSearch("");
    setAsOf(serialToISO(todaySerial()));
    setProrate(DEFAULT_WORKYEAR_PARAMS.prorateLastPeriod);
    setIncludeComp(DEFAULT_WORKYEAR_PARAMS.includeCompensation);
    setMasterSheet(DEFAULT_WORKYEAR_PARAMS.masterSheet);
  };

  const onPick = (f: File | null) => {
    setFile(f);
    setResult(null);
    setError(null);
    setStatus("idle");
    setProgress(0);
  };

  const run = useCallback(async () => {
    if (!file) return;
    const params: WorkyearParams = {
      asOfDate: isoToSerial(asOf) ?? todaySerial(),
      prorateLastPeriod: prorate,
      includeCompensation: includeComp,
      masterSheet: masterSheet.trim() || DEFAULT_WORKYEAR_PARAMS.masterSheet,
    };
    setStatus("processing");
    setProgress(0);
    setError(null);
    setResult(null);
    setPreview("checks");
    try {
      const res = await processWorkyear(file, params, (step, pct) => {
        setStepText(step);
        setProgress(pct);
      });
      setResult(res);
      setStatus("done");
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : "Naməlum xəta baş verdi.");
      setStatus("error");
    }
  }, [file, asOf, prorate, includeComp, masterSheet]);

  const failedHard = result ? result.checks.filter((c) => c.hard && !c.ok) : [];

  const visibleRows = useMemo(() => {
    if (!result) return [];
    let list = result.rows;
    if (search.trim()) {
      const q = search.toLowerCase().trim();
      list = list.filter(
        (r) => r.empCode.toLowerCase().includes(q) || r.fullName.toLowerCase().includes(q),
      );
    }
    return list.slice(0, 300);
  }, [result, search]);

  return (
    <>
      <section className="glass mx-auto max-w-4xl rounded-3xl p-6 shadow-soft md:p-8">
        <div className="mb-6">
          <h2 className="text-lg font-bold text-slate-900">Vacation by Workyear</h2>
          <p className="mt-1 text-sm text-slate-500">
            vacation_days_by_period_SPEC.md-yə əsasən: hər əməkdaş × hər iş ili üçün bir sətir —
            tam illik hüquq (əsas + əlavə), istifadə FIFO ilə ən köhnə dövrdən silinir. Nəticə:{" "}
            <code className="rounded bg-slate-100 px-1.5">Vacation_days_by_period.xlsx</code> (tək
            səhifə).
          </p>
        </div>

        <FileDrop
          file={file}
          onPick={onPick}
          hint={
            <>
              Tələb olunan səhifələr: master (<code className="rounded bg-slate-100 px-1.5">{masterSheet || "2015-2026"}</code>),
              illik səhifələr (<code className="rounded bg-slate-100 px-1.5">2018…2026</code>) və{" "}
              <code className="rounded bg-slate-100 px-1.5">orders</code>
            </>
          }
        />

        <div className="mt-4 grid grid-cols-2 gap-3 rounded-2xl border border-slate-200 bg-white/60 p-4 md:grid-cols-4">
          <label className="flex flex-col gap-1 text-xs font-medium text-slate-500">
            AS_OF_DATE
            <input
              type="date"
              value={asOf}
              onChange={(e) => setAsOf(e.target.value)}
              className="rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-sm shadow-sm focus:border-brand-400 focus:outline-none focus:ring-2 focus:ring-brand-200"
            />
          </label>
          <label className="flex flex-col gap-1 text-xs font-medium text-slate-500">
            MASTER_SHEET
            <input
              value={masterSheet}
              onChange={(e) => setMasterSheet(e.target.value)}
              className="rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-sm shadow-sm focus:border-brand-400 focus:outline-none focus:ring-2 focus:ring-brand-200"
            />
          </label>
          <label className="flex items-center gap-2 self-end pb-1 text-sm text-slate-700">
            <input
              type="checkbox"
              checked={prorate}
              onChange={(e) => setProrate(e.target.checked)}
              className="h-4 w-4 rounded border-slate-300 text-brand-600 focus:ring-brand-500"
            />
            PRORATE_LAST_PERIOD
          </label>
          <label className="flex items-center gap-2 self-end pb-1 text-sm text-slate-700">
            <input
              type="checkbox"
              checked={includeComp}
              onChange={(e) => setIncludeComp(e.target.checked)}
              className="h-4 w-4 rounded border-slate-300 text-brand-600 focus:ring-brand-500"
            />
            INCLUDE_COMPENSATION
          </label>
        </div>

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
                <PlayIcon /> Hesabatı yarat
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
          <section className="mt-8 grid grid-cols-2 gap-3 md:grid-cols-5">
            <Stat label="Əməkdaş" value={result.stats.employees} />
            <Stat label="Dövr (sətir)" value={result.stats.periods} accent="brand" />
            <Stat label="İstifadə (gün)" value={result.stats.totalUsed} tone="amber" />
            <Stat label="Sayılan əmr" value={result.stats.ordersIncluded} />
            <Stat label="Used uzlaşması" value={`${result.stats.matchPct}%`} tone={result.stats.matchPct >= 95 ? "emerald" : "rose"} />
          </section>

          <section className="glass mt-6 rounded-3xl p-4 shadow-soft md:p-6">
            <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
              <div>
                <h2 className="text-lg font-bold text-slate-900">Vacation days by period</h2>
                <p className="text-sm text-slate-500">
                  AS_OF: <span className="font-semibold text-slate-700">{fmtDate(result.params.asOfDate)}</span> ·
                  illik səhifələr:{" "}
                  <span className="font-semibold text-slate-700">
                    {result.stats.yearSheets.length ? result.stats.yearSheets.join(", ") : "yoxdur"}
                  </span>{" "}
                  · müddət {(result.durationMs / 1000).toFixed(1)} san ·{" "}
                  {result.allPassed ? (
                    <span className="inline-flex items-center rounded-full bg-emerald-100 px-2 py-0.5 text-xs font-semibold text-emerald-800 ring-1 ring-inset ring-emerald-200">
                      YOXLAMALAR KEÇDİ
                    </span>
                  ) : (
                    <span className="inline-flex items-center rounded-full bg-rose-100 px-2 py-0.5 text-xs font-semibold text-rose-800 ring-1 ring-inset ring-rose-200">
                      {failedHard.length} YOXLAMA UĞURSUZ
                    </span>
                  )}
                </p>
              </div>
              <button
                onClick={() => exportWorkyear(result)}
                className="inline-flex items-center gap-2 rounded-xl bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-soft transition hover:bg-emerald-700"
              >
                <DownloadIcon /> Vacation_days_by_period.xlsx
              </button>
            </div>

            {!result.allPassed && (
              <div className="mb-4">
                <ErrorBox
                  message={`Diqqət: ${failedHard.map((c) => c.name).join("; ")} yoxlaması uğursuzdur — nəticəni yükləməzdən əvvəl aşağıdakı təfərrüatları və anomaliya loqunu nəzərdən keçirin.`}
                />
              </div>
            )}

            <div className="mb-4 flex flex-wrap items-center gap-3">
              <div className="inline-flex rounded-xl border border-slate-200 bg-white p-1">
                <SegmentBtn active={preview === "checks"} onClick={() => setPreview("checks")}>
                  Yoxlamalar ({result.checks.length})
                </SegmentBtn>
                <SegmentBtn active={preview === "rows"} onClick={() => setPreview("rows")}>
                  Dövrlər ({result.rows.length})
                </SegmentBtn>
                <SegmentBtn active={preview === "logs"} onClick={() => setPreview("logs")}>
                  Loq ({result.logs.length})
                </SegmentBtn>
              </div>
              {preview === "rows" && (
                <input
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Kod və ya ad axtar"
                  className="w-56 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-sm shadow-sm focus:border-brand-400 focus:outline-none focus:ring-2 focus:ring-brand-200"
                />
              )}
            </div>

            {preview === "checks" && (
              <div className="space-y-2">
                {result.checks.map((c) => (
                  <div
                    key={c.name}
                    className={`flex items-start gap-3 rounded-2xl border p-3 ${
                      c.ok ? "border-emerald-200 bg-emerald-50/50" : "border-rose-200 bg-rose-50/60"
                    }`}
                  >
                    <span
                      className={`mt-0.5 inline-flex h-6 min-w-[3.5rem] items-center justify-center rounded-full px-2 text-xs font-bold ${
                        c.ok ? "bg-emerald-600 text-white" : "bg-rose-600 text-white"
                      }`}
                    >
                      {c.ok ? "PASS" : "FAIL"}
                    </span>
                    <div>
                      <p className="text-sm font-semibold text-slate-800">
                        {c.name}
                        {!c.hard && (
                          <span className="ml-2 rounded bg-slate-200 px-1.5 py-0.5 text-[10px] font-medium uppercase text-slate-600">
                            info
                          </span>
                        )}
                      </p>
                      <p className="mt-0.5 text-xs leading-relaxed text-slate-600">{c.detail}</p>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {preview === "rows" && (
              <div className="overflow-x-auto rounded-2xl border border-slate-200 bg-white">
                <table className="min-w-full divide-y divide-slate-200 text-sm">
                  <thead className="bg-slate-50">
                    <tr className="text-left text-xs font-semibold uppercase tracking-wider text-slate-500">
                      {["Kod", "Ad", "Hire", "Dövr başlanğıcı", "Dövr sonu", "Main", "Add", "Hüquq", "Used main", "Used add", "Unused"].map((h) => (
                        <th key={h} className="px-3 py-3">
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {visibleRows.map((r, i) => {
                      const unused = Math.round((r.mainDays + r.addDays - r.usedMain - r.usedAdd) * 100) / 100;
                      return (
                        <tr key={i} className="hover:bg-slate-50/60">
                          <td className="px-3 py-1.5 font-medium tabular-nums text-slate-700">{r.empCode}</td>
                          <td className="px-3 py-1.5 text-slate-700">{r.fullName}</td>
                          <td className="px-3 py-1.5 tabular-nums text-slate-500">{fmtDate(r.hireDate)}</td>
                          <td className="px-3 py-1.5 tabular-nums text-slate-600">{fmtDate(r.periodStart)}</td>
                          <td className="px-3 py-1.5 tabular-nums text-slate-600">{fmtDate(r.periodEnd)}</td>
                          <td className="px-3 py-1.5 text-right tabular-nums text-slate-700">{r.mainDays}</td>
                          <td className="px-3 py-1.5 text-right tabular-nums text-slate-700">{r.addDays}</td>
                          <td className="px-3 py-1.5 text-right tabular-nums font-medium text-slate-800">
                            {Math.round((r.mainDays + r.addDays) * 100) / 100}
                          </td>
                          <td className="px-3 py-1.5 text-right tabular-nums text-slate-700">{r.usedMain}</td>
                          <td className="px-3 py-1.5 text-right tabular-nums text-slate-700">{r.usedAdd}</td>
                          <td className="px-3 py-1.5 text-right">
                            <span
                              className={`inline-flex min-w-[2.5rem] justify-center rounded-full px-2 py-0.5 text-xs font-semibold ring-1 ring-inset tabular-nums ${
                                unused < 0
                                  ? "bg-rose-100 text-rose-800 ring-rose-200"
                                  : unused === 0
                                    ? "bg-slate-100 text-slate-600 ring-slate-200"
                                    : "bg-emerald-100 text-emerald-800 ring-emerald-200"
                              }`}
                            >
                              {unused}
                            </span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                {result.rows.length > visibleRows.length && (
                  <p className="border-t border-slate-100 px-3 py-2 text-xs text-slate-400">
                    İlk {visibleRows.length} sətir göstərilir; tam siyahı export faylındadır.
                  </p>
                )}
              </div>
            )}

            {preview === "logs" && (
              <div className="overflow-x-auto rounded-2xl border border-slate-200 bg-white">
                <table className="min-w-full divide-y divide-slate-200 text-sm">
                  <thead className="bg-slate-50">
                    <tr className="text-left text-xs font-semibold uppercase tracking-wider text-slate-500">
                      <th className="px-3 py-3">#</th>
                      <th className="px-3 py-3">Tip</th>
                      <th className="px-3 py-3">Əməkdaş</th>
                      <th className="px-3 py-3">Təfərrüat</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {result.logs.length === 0 ? (
                      <tr>
                        <td colSpan={4} className="px-3 py-10 text-center text-sm text-slate-400">
                          Log qeydi yoxdur.
                        </td>
                      </tr>
                    ) : (
                      result.logs.slice(0, 300).map((l, i) => (
                        <tr key={i} className="hover:bg-slate-50/60">
                          <td className="px-3 py-1.5 tabular-nums text-slate-400">{i + 1}</td>
                          <td className="px-3 py-1.5">
                            <span className="inline-flex rounded-full bg-amber-100 px-2 py-0.5 text-xs font-medium text-amber-800 ring-1 ring-inset ring-amber-200">
                              {l.tip}
                            </span>
                          </td>
                          <td className="px-3 py-1.5 text-slate-700">{l.emp}</td>
                          <td className="max-w-xl px-3 py-1.5 text-xs text-slate-600">{l.detail}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            )}
          </section>
        </>
      )}
    </>
  );
}
