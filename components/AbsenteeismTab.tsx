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
type Op = "none" | "lt" | "lte" | "eq" | "gte" | "gt" | "between";
type SortOrder = "desc" | "asc";

const OP_LABELS: Record<Op, string> = {
  none: "Filtr yoxdur",
  lt: "< (kiçik)",
  lte: "≤ (kiçik və ya bərabər)",
  eq: "= (bərabər)",
  gte: "≥ (böyük və ya bərabər)",
  gt: "> (böyük)",
  between: "Aralıq (between)",
};

interface AppliedFilter {
  op: Op;
  val1: number | null;
  val2: number | null;
}

const NO_FILTER: AppliedFilter = { op: "none", val1: null, val2: null };

export function AbsenteeismTab() {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [progress, setProgress] = useState(0);
  const [stepText, setStepText] = useState("");
  const [result, setResult] = useState<AbsenteeismResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState("");
  const [selected, setSelected] = useState<string | null>(null);

  // Filter form state
  const [op, setOp] = useState<Op>("none");
  const [val1Input, setVal1Input] = useState("");
  const [val2Input, setVal2Input] = useState("");
  const [applied, setApplied] = useState<AppliedFilter>(NO_FILTER);
  const [sortOrder, setSortOrder] = useState<SortOrder>("desc");

  const reset = () => {
    setFile(null);
    setStatus("idle");
    setProgress(0);
    setStepText("");
    setResult(null);
    setError(null);
    setSearch("");
    setSelected(null);
    setOp("none");
    setVal1Input("");
    setVal2Input("");
    setApplied(NO_FILTER);
    setSortOrder("desc");
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

  const applyFilter = () => {
    if (op === "none") {
      setApplied(NO_FILTER);
      return;
    }
    const v1 = val1Input.trim() === "" ? null : Number(val1Input);
    const v2 = val2Input.trim() === "" ? null : Number(val2Input);
    setApplied({ op, val1: v1, val2: v2 });
  };

  const clearFilter = () => {
    setOp("none");
    setVal1Input("");
    setVal2Input("");
    setApplied(NO_FILTER);
  };

  const visibleNetice = useMemo(() => {
    if (!result) return [] as NeticeRow[];
    let list = result.netice;

    if (applied.op !== "none") {
      const a = applied.val1;
      const b = applied.val2;
      list = list.filter((r) => {
        const t = r.total;
        switch (applied.op) {
          case "lt":
            return a !== null && t < a;
          case "lte":
            return a !== null && t <= a;
          case "eq":
            return a !== null && t === a;
          case "gte":
            return a !== null && t >= a;
          case "gt":
            return a !== null && t > a;
          case "between": {
            if (a === null || b === null) return true;
            const lo = Math.min(a, b);
            const hi = Math.max(a, b);
            return t >= lo && t <= hi;
          }
          default:
            return true;
        }
      });
    }

    if (search.trim()) {
      const q = search.toLowerCase().trim();
      list = list.filter(
        (r) => r.code.toLowerCase().includes(q) || r.name.toLowerCase().includes(q),
      );
    }

    const sorted = [...list].sort((x, y) =>
      sortOrder === "desc" ? y.total - x.total : x.total - y.total,
    );
    return sorted;
  }, [result, applied, search, sortOrder]);

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

  const fileBase = () => file?.name?.replace(/\.(xlsx|xlsm|xls)$/i, "") ?? "BirCheck";

  const exportSelected = () => {
    if (!result) return;
    const codes = new Set(visibleNetice.map((n) => n.code));
    exportAbsenteeism(result, `${fileBase()}_absenteeism_selected.xlsx`, {
      netice: visibleNetice,
      muqayise: result.muqayise.filter((m) => codes.has(m.code)),
    });
  };

  const exportAll = () => {
    if (!result) return;
    exportAbsenteeism(result, `${fileBase()}_absenteeism.xlsx`);
  };

  const filterActive = applied.op !== "none";

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
            <div className="mb-4 flex flex-col gap-3">
              <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                <div>
                  <h2 className="text-lg font-bold text-slate-900">Nəticə</h2>
                  <p className="text-sm text-slate-500">
                    Aralıq:{" "}
                    <span className="font-semibold text-slate-700">
                      {formatExcelDate(result.periodStart)} – {formatExcelDate(result.periodEnd)}
                    </span>{" "}
                    · <span className="font-semibold text-slate-700">{visibleNetice.length}</span>{" "}
                    / {result.totalEmployees} işçi
                    {filterActive && (
                      <span className="ml-2 inline-flex items-center gap-1 rounded-full bg-brand-50 px-2 py-0.5 text-xs font-medium text-brand-700 ring-1 ring-inset ring-brand-200">
                        Aktiv filtr: müddət {filterDescription(applied)}
                      </span>
                    )}
                  </p>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    onClick={exportSelected}
                    disabled={visibleNetice.length === 0}
                    className="inline-flex items-center gap-2 rounded-xl bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-soft transition hover:bg-emerald-700 disabled:cursor-not-allowed disabled:bg-slate-300"
                  >
                    <DownloadIcon /> Export selected ({visibleNetice.length})
                  </button>
                  <button
                    onClick={exportAll}
                    className="inline-flex items-center gap-2 rounded-xl border border-emerald-600 bg-white px-4 py-2 text-sm font-semibold text-emerald-700 transition hover:bg-emerald-50"
                  >
                    <DownloadIcon /> Export all ({result.totalEmployees})
                  </button>
                </div>
              </div>

              <div className="flex flex-wrap items-end gap-3 rounded-2xl border border-slate-200 bg-white p-3">
                <Field label="Axtar">
                  <input
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                    placeholder="Kod və ya ad"
                    className="w-52 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-sm shadow-sm focus:border-brand-400 focus:outline-none focus:ring-2 focus:ring-brand-200"
                  />
                </Field>

                <Field label="Müddət filtri">
                  <select
                    value={op}
                    onChange={(e) => setOp(e.target.value as Op)}
                    className="rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-sm shadow-sm focus:border-brand-400 focus:outline-none focus:ring-2 focus:ring-brand-200"
                  >
                    {(Object.keys(OP_LABELS) as Op[]).map((o) => (
                      <option key={o} value={o}>
                        {OP_LABELS[o]}
                      </option>
                    ))}
                  </select>
                </Field>

                {op !== "none" && (
                  <Field label={op === "between" ? "Min" : "Dəyər"}>
                    <input
                      type="number"
                      value={val1Input}
                      onChange={(e) => setVal1Input(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === "Enter") applyFilter();
                      }}
                      placeholder="0"
                      className="w-24 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-sm shadow-sm focus:border-brand-400 focus:outline-none focus:ring-2 focus:ring-brand-200"
                    />
                  </Field>
                )}

                {op === "between" && (
                  <Field label="Max">
                    <input
                      type="number"
                      value={val2Input}
                      onChange={(e) => setVal2Input(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === "Enter") applyFilter();
                      }}
                      placeholder="0"
                      className="w-24 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-sm shadow-sm focus:border-brand-400 focus:outline-none focus:ring-2 focus:ring-brand-200"
                    />
                  </Field>
                )}

                <Field label="Sıralama">
                  <select
                    value={sortOrder}
                    onChange={(e) => setSortOrder(e.target.value as SortOrder)}
                    className="rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-sm shadow-sm focus:border-brand-400 focus:outline-none focus:ring-2 focus:ring-brand-200"
                  >
                    <option value="desc">Largest → smallest</option>
                    <option value="asc">Smallest → largest</option>
                  </select>
                </Field>

                <div className="flex items-center gap-2 self-end">
                  <button
                    onClick={applyFilter}
                    disabled={op === "none"}
                    className="rounded-lg bg-brand-600 px-4 py-1.5 text-sm font-semibold text-white shadow-sm transition hover:bg-brand-700 disabled:cursor-not-allowed disabled:bg-slate-300"
                  >
                    Tətbiq et
                  </button>
                  <button
                    onClick={clearFilter}
                    disabled={!filterActive && op === "none"}
                    className="rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-sm font-medium text-slate-700 transition hover:bg-slate-50 disabled:opacity-50"
                  >
                    Sıfırla
                  </button>
                </div>
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
                          Bu filtrə uyğun sətir yoxdur.
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

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <label className="flex flex-col gap-1 text-xs font-medium text-slate-500">
      {label}
      {children}
    </label>
  );
}

function filterDescription(f: AppliedFilter): string {
  const a = f.val1 ?? "?";
  const b = f.val2 ?? "?";
  switch (f.op) {
    case "lt":
      return `< ${a}`;
    case "lte":
      return `≤ ${a}`;
    case "eq":
      return `= ${a}`;
    case "gte":
      return `≥ ${a}`;
    case "gt":
      return `> ${a}`;
    case "between":
      return `∈ [${Math.min(Number(a), Number(b))}; ${Math.max(Number(a), Number(b))}]`;
    default:
      return "";
  }
}
