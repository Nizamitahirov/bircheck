"use client";

import { useCallback, useMemo, useState } from "react";
import {
  processAbsenteeism,
  exportAbsenteeism,
  formatExcelDate,
  type AbsenteeismResult,
  type NeticeRow,
} from "@/lib/absenteeism";
import {
  DownloadIcon,
  ErrorBox,
  FileDrop,
  PlayIcon,
  ProgressBar,
  Spinner,
  Stat,
} from "./ui";

type Status = "idle" | "processing" | "done" | "error";

export function AbsenteeismTab() {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [progress, setProgress] = useState(0);
  const [stepText, setStepText] = useState("");
  const [result, setResult] = useState<AbsenteeismResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState("");
  const [selected, setSelected] = useState<string | null>(null);

  const reset = () => {
    setFile(null);
    setStatus("idle");
    setProgress(0);
    setStepText("");
    setResult(null);
    setError(null);
    setSearch("");
    setSelected(null);
  };

  const onPick = (f: File | null) => {
    setError(null);
    setResult(null);
    setStatus("idle");
    setProgress(0);
    setSelected(null);
    setFile(f);
  };

  const run = useCallback(async () => {
    if (!file) return;
    setStatus("processing");
    setProgress(0);
    setError(null);
    setResult(null);
    setSelected(null);
    try {
      const res = await processAbsenteeism(file, (step, pct) => {
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

  const download = () => {
    if (!result) return;
    const base = file?.name?.replace(/\.(xlsx|xlsm|xls)$/i, "") ?? "BirCheck";
    exportAbsenteeism(result, `${base}_absenteeism.xlsx`);
  };

  const visibleNetice = useMemo(() => {
    if (!result) return [] as NeticeRow[];
    let list = result.netice;
    if (search.trim()) {
      const q = search.toLowerCase().trim();
      list = list.filter(
        (r) => r.code.toLowerCase().includes(q) || r.name.toLowerCase().includes(q),
      );
    }
    return list;
  }, [result, search]);

  const stats = useMemo(() => {
    if (!result) return { withAny: 0, totalSum: 0, max: 0 };
    let withAny = 0;
    let totalSum = 0;
    let max = 0;
    for (const r of result.netice) {
      if (r.total !== 0) withAny++;
      totalSum += r.total;
      if (r.total > max) max = r.total;
    }
    return { withAny, totalSum, max };
  }, [result]);

  const selectedDetails = useMemo(() => {
    if (!result || !selected) return [];
    return result.muqayise.filter((m) => m.code === selected);
  }, [result, selected]);

  return (
    <>
      <section className="glass mx-auto max-w-3xl rounded-3xl p-6 shadow-soft md:p-8">
        <FileDrop
          file={file}
          onPick={onPick}
          hint={
            <>
              Tələb olunan sheet:{" "}
              <code className="rounded bg-slate-100 px-1.5">Data</code> · Sütunlar: A=Personal kod,
              C=Növ kodu, E=Başlama, F=Bitmə · Aralıq: son 63 gün
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
                <PlayIcon /> Hesablamanı başlat
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
            <Stat label="Davamiyyət olan" value={stats.withAny} accent="brand" />
            <Stat label="Ümumi müddət" value={stats.totalSum} tone="amber" />
            <Stat label="Ən yüksək" value={stats.max} tone="rose" />
          </section>

          <section className="glass mt-6 rounded-3xl p-4 shadow-soft md:p-6">
            <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
              <div>
                <h2 className="text-lg font-bold text-slate-900">Nəticə</h2>
                <p className="text-sm text-slate-500">
                  Aralıq:{" "}
                  <span className="font-semibold text-slate-700">
                    {formatExcelDate(result.periodStart)} – {formatExcelDate(result.periodEnd)}
                  </span>{" "}
                  · <span className="font-semibold text-slate-700">{visibleNetice.length}</span>{" "}
                  işçi göstərilir
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <input
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Axtar (kod, ad)"
                  className="w-56 rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm focus:border-brand-400 focus:outline-none focus:ring-2 focus:ring-brand-200"
                />
                <button
                  onClick={download}
                  className="inline-flex items-center gap-2 rounded-xl bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-soft transition hover:bg-emerald-700"
                >
                  <DownloadIcon /> Endir (.xlsx)
                </button>
              </div>
            </div>

            <div className="grid gap-4 lg:grid-cols-[1.1fr_1fr]">
              <div className="overflow-x-auto rounded-2xl border border-slate-200 bg-white">
                <table className="min-w-full divide-y divide-slate-200 text-sm">
                  <thead className="bg-slate-50">
                    <tr className="text-left text-xs font-semibold uppercase tracking-wider text-slate-500">
                      <th className="px-3 py-3">#</th>
                      <th className="px-3 py-3">Personal kod</th>
                      <th className="px-3 py-3">Ad</th>
                      <th className="px-3 py-3 text-right">Müddət</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {visibleNetice.length === 0 ? (
                      <tr>
                        <td colSpan={4} className="px-3 py-10 text-center text-sm text-slate-400">
                          Heç nə tapılmadı.
                        </td>
                      </tr>
                    ) : (
                      visibleNetice.map((r, i) => {
                        const active = r.code === selected;
                        const tone =
                          r.total === 0
                            ? "bg-slate-100 text-slate-600 ring-slate-200"
                            : r.total >= 5
                              ? "bg-rose-100 text-rose-800 ring-rose-200"
                              : "bg-amber-100 text-amber-800 ring-amber-200";
                        return (
                          <tr
                            key={r.code}
                            onClick={() => setSelected(r.code)}
                            className={`cursor-pointer transition ${
                              active ? "bg-brand-50/70" : "hover:bg-slate-50/60"
                            }`}
                          >
                            <td className="px-3 py-2 tabular-nums text-slate-400">{i + 1}</td>
                            <td className="px-3 py-2 font-medium tabular-nums text-slate-700">
                              {r.code}
                            </td>
                            <td className="px-3 py-2 text-slate-700">{r.name}</td>
                            <td className="px-3 py-2 text-right">
                              <span
                                className={`inline-flex min-w-[2.5rem] justify-center rounded-full px-2 py-0.5 text-xs font-semibold ring-1 ring-inset tabular-nums ${tone}`}
                              >
                                {r.total}
                              </span>
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>

              <div className="rounded-2xl border border-slate-200 bg-white p-4">
                <h3 className="mb-3 text-sm font-semibold text-slate-700">
                  Müqayisə {selected ? `· ${selected}` : ""}
                </h3>
                {!selected ? (
                  <p className="text-sm text-slate-400">
                    Detalları görmək üçün sol cədvəldən bir işçi seçin.
                  </p>
                ) : selectedDetails.length === 0 ? (
                  <p className="text-sm text-slate-400">Bu işçi üçün davamiyyət qeydi yoxdur.</p>
                ) : (
                  <div className="overflow-y-auto" style={{ maxHeight: "420px" }}>
                    <table className="min-w-full divide-y divide-slate-200 text-sm">
                      <thead className="sticky top-0 bg-slate-50">
                        <tr className="text-left text-xs font-semibold uppercase tracking-wider text-slate-500">
                          <th className="px-3 py-2">Tarix</th>
                          <th className="px-3 py-2 text-right">Növ</th>
                          <th className="px-3 py-2 text-right">Müddət</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {selectedDetails.map((d, i) => (
                          <tr key={i} className="hover:bg-slate-50/60">
                            <td className="px-3 py-1.5 tabular-nums text-slate-700">
                              {formatExcelDate(d.dateSerial)}
                            </td>
                            <td className="px-3 py-1.5 text-right tabular-nums text-slate-600">
                              {d.type}
                            </td>
                            <td className="px-3 py-1.5 text-right tabular-nums">
                              <span
                                className={`inline-flex min-w-[2rem] justify-center rounded-full px-2 py-0.5 text-xs font-semibold ${
                                  d.weight < 0
                                    ? "bg-rose-100 text-rose-800"
                                    : d.weight >= 4
                                      ? "bg-amber-100 text-amber-800"
                                      : "bg-emerald-100 text-emerald-800"
                                }`}
                              >
                                {d.weight}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </div>
          </section>
        </>
      )}
    </>
  );
}
