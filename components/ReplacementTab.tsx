"use client";

import { useCallback, useMemo, useState } from "react";
import {
  DEFAULT_PARAMS,
  exportReplacement,
  fmtDate,
  fmtDateTime,
  processReplacement,
  type ReplacementParams,
  type ReplacementResult,
} from "@/lib/replacement";
import { DownloadIcon, ErrorBox, FileDrop, PlayIcon, ProgressBar, SegmentBtn, Spinner, Stat } from "./ui";

type Status = "idle" | "processing" | "done" | "error";
type Preview = "checks" | "xeta" | "dup" | "yekun";

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

interface Slot {
  key: "inzibatci" | "vacation" | "filial" | "resmi" | "vezifeMaas";
  title: string;
  hint: string;
}

const SLOTS: Slot[] = [
  { key: "inzibatci", title: "1. İnzibatçılıq prosesi", hint: "İnzibatciliq prosesi 2.xlsx — Sheet1" },
  {
    key: "vacation",
    title: "2. Vacation / Time off replacement",
    hint: "Vacation, Time off replacement.xls — hər iki sheet oxunur",
  },
  { key: "filial", title: "3. Filial baza", hint: "Filial baza .xlsx — Sheet1 (Müvəqqəti keçid)" },
  {
    key: "resmi",
    title: "4. Rəsmi əvəzetmə (position info)",
    hint: "Rəsmi əvəzetmə(position info).xlsx — qırmızı sətirlər avtomatik aşkarlanır",
  },
  { key: "vezifeMaas", title: "5. Vəzifə-Maaş", hint: "Vəzifə-Maaş.xlsx — Yanvar–İyun 2026 (yalnız lookup)" },
];

export function ReplacementTab() {
  const [files, setFiles] = useState<Record<Slot["key"], File | null>>({
    inzibatci: null,
    vacation: null,
    filial: null,
    resmi: null,
    vezifeMaas: null,
  });
  const [status, setStatus] = useState<Status>("idle");
  const [progress, setProgress] = useState(0);
  const [stepText, setStepText] = useState("");
  const [result, setResult] = useState<ReplacementResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [preview, setPreview] = useState<Preview>("checks");

  const [filterStart, setFilterStart] = useState(serialToISO(DEFAULT_PARAMS.filterStart));
  const [capDate, setCapDate] = useState(serialToISO(DEFAULT_PARAMS.capDate));
  const [winStart, setWinStart] = useState(serialToISO(DEFAULT_PARAMS.winStart));
  const [winEnd, setWinEnd] = useState(serialToISO(DEFAULT_PARAMS.winEnd));

  const allPicked = SLOTS.every((s) => files[s.key]);

  const winDays = useMemo(() => {
    const a = isoToSerial(winStart);
    const b = isoToSerial(winEnd);
    return a !== null && b !== null ? b - a + 1 : null;
  }, [winStart, winEnd]);

  const reset = () => {
    setFiles({ inzibatci: null, vacation: null, filial: null, resmi: null, vezifeMaas: null });
    setStatus("idle");
    setProgress(0);
    setStepText("");
    setResult(null);
    setError(null);
    setPreview("checks");
    setFilterStart(serialToISO(DEFAULT_PARAMS.filterStart));
    setCapDate(serialToISO(DEFAULT_PARAMS.capDate));
    setWinStart(serialToISO(DEFAULT_PARAMS.winStart));
    setWinEnd(serialToISO(DEFAULT_PARAMS.winEnd));
  };

  const pick = (key: Slot["key"]) => (f: File | null) => {
    setFiles((prev) => ({ ...prev, [key]: f }));
    setResult(null);
    setError(null);
    setStatus("idle");
    setProgress(0);
  };

  const run = useCallback(async () => {
    if (!allPicked) return;
    const params: ReplacementParams = {
      filterStart: isoToSerial(filterStart) ?? DEFAULT_PARAMS.filterStart,
      capDate: isoToSerial(capDate) ?? DEFAULT_PARAMS.capDate,
      winStart: isoToSerial(winStart) ?? DEFAULT_PARAMS.winStart,
      winEnd: isoToSerial(winEnd) ?? DEFAULT_PARAMS.winEnd,
    };
    setStatus("processing");
    setProgress(0);
    setError(null);
    setResult(null);
    setPreview("checks");
    try {
      const res = await processReplacement(
        {
          inzibatci: files.inzibatci!,
          vacation: files.vacation!,
          filial: files.filial!,
          resmi: files.resmi!,
          vezifeMaas: files.vezifeMaas!,
        },
        params,
        (step, pct) => {
          setStepText(step);
          setProgress(pct);
        },
      );
      setResult(res);
      setStatus("done");
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : "Naməlum xəta baş verdi.");
      setStatus("error");
    }
  }, [allPicked, files, filterStart, capDate, winStart, winEnd]);

  const failedHard = result ? result.checks.filter((c) => c.hard && !c.ok) : [];

  return (
    <>
      <section className="glass mx-auto max-w-6xl rounded-3xl p-6 shadow-soft md:p-8">
        <div className="mb-6">
          <h2 className="text-lg font-bold text-slate-900">Replacement check</h2>
          <p className="mt-1 text-sm text-slate-500">
            SPEC.md-yə əsasən 5 mənbə Excel faylı vahid sxemə salınır, duplikat / üst-üstə düşmə /
            iç-içə yoxlamaları aparılır və bütün məcburi yoxlamalar (bölmə 12) keçdikdə{" "}
            <code className="rounded bg-slate-100 px-1.5">Əvəzetmə_Hesabatı.xlsx</code> qaytarılır.
          </p>
        </div>

        <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-3">
          {SLOTS.map((s) => (
            <div key={s.key}>
              <p className="mb-2 text-sm font-semibold text-slate-700">{s.title}</p>
              <FileDrop file={files[s.key]} onPick={pick(s.key)} hint={s.hint} />
            </div>
          ))}

          <div>
            <p className="mb-2 text-sm font-semibold text-slate-700">Parametrlər</p>
            <div className="grid grid-cols-2 gap-3 rounded-2xl border border-slate-200 bg-white/60 p-4">
              <DateField label="Filtr başlanğıcı" value={filterStart} onChange={setFilterStart} />
              <DateField label="Hesablama sərhədi (CAP)" value={capDate} onChange={setCapDate} />
              <DateField label="Pəncərə başlanğıcı" value={winStart} onChange={setWinStart} />
              <DateField label="Pəncərə sonu" value={winEnd} onChange={setWinEnd} />
              <p
                className={`col-span-2 text-xs ${
                  winDays === 181 ? "text-slate-500" : "font-semibold text-rose-600"
                }`}
              >
                Pəncərə: {winDays ?? "?"} gün {winDays === 181 ? "✓" : "— dəqiq 181 gün olmalıdır"}
              </p>
            </div>
          </div>
        </div>

        <div className="mt-6 flex flex-col gap-3 sm:flex-row">
          <button
            onClick={run}
            disabled={!allPicked || status === "processing"}
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
          <section className="mt-8 grid grid-cols-2 gap-3 md:grid-cols-5">
            <Stat label="Normallaşdırılan sətir" value={result.counters.normalized} />
            <Stat label="Təmiz Data" value={result.clean.length} accent="brand" />
            <Stat label="Xərc yaradan (Bəli)" value={result.counters.beli} tone="amber" />
            <Stat label="Yekun sətri" value={result.yekun.length} tone="emerald" />
            <Stat label="Unikal əməkdaş" value={result.cem.length} />
          </section>
          <section className="mt-3 grid grid-cols-2 gap-3 md:grid-cols-5">
            <Stat label="Xəta Loqu" value={result.errors.length} tone="rose" />
            <Stat
              label="Duplikat (daxili/arası)"
              value={`${result.counters.dupIntra} / ${result.counters.dupCross}`}
            />
            <Stat
              label="Üst-üstə düşmə / iç-içə"
              value={`${result.counters.overlapRows} / ${result.counters.nestedRows}`}
            />
            <Stat label="Qırmızı sətir" value={result.counters.redRows} />
            <Stat label="Max unikal gün" value={result.counters.maxUnikalGun} />
          </section>

          <section className="glass mt-6 rounded-3xl p-4 shadow-soft md:p-6">
            <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
              <div>
                <h2 className="text-lg font-bold text-slate-900">Nəticə</h2>
                <p className="text-sm text-slate-500">
                  Məcburi yoxlamalar:{" "}
                  {result.allPassed ? (
                    <span className="inline-flex items-center rounded-full bg-emerald-100 px-2 py-0.5 text-xs font-semibold text-emerald-800 ring-1 ring-inset ring-emerald-200">
                      HAMISI KEÇDİ
                    </span>
                  ) : (
                    <span className="inline-flex items-center rounded-full bg-rose-100 px-2 py-0.5 text-xs font-semibold text-rose-800 ring-1 ring-inset ring-rose-200">
                      {failedHard.length} YOXLAMA UĞURSUZ
                    </span>
                  )}{" "}
                  · müddət {(result.durationMs / 1000).toFixed(1)} san
                </p>
              </div>
              <button
                onClick={() => exportReplacement(result)}
                disabled={!result.allPassed}
                title={
                  result.allPassed
                    ? "Əvəzetmə_Hesabatı.xlsx yüklə"
                    : "SPEC bölmə 12: məcburi yoxlamalardan biri uğursuzdursa nəticə qaytarılmır"
                }
                className="inline-flex items-center gap-2 rounded-xl bg-emerald-600 px-4 py-2 text-sm font-semibold text-white shadow-soft transition hover:bg-emerald-700 disabled:cursor-not-allowed disabled:bg-slate-300"
              >
                <DownloadIcon /> Əvəzetmə_Hesabatı.xlsx
              </button>
            </div>

            {!result.allPassed && (
              <div className="mb-4">
                <ErrorBox
                  message={`SPEC bölmə 12-yə əsasən nəticə qaytarılmır: ${failedHard
                    .map((c) => c.name)
                    .join("; ")}. Təfərrüatlar aşağıdakı yoxlama siyahısındadır.`}
                />
              </div>
            )}

            <div className="mb-4 inline-flex rounded-xl border border-slate-200 bg-white p-1">
              <SegmentBtn active={preview === "checks"} onClick={() => setPreview("checks")}>
                Yoxlamalar ({result.checks.length})
              </SegmentBtn>
              <SegmentBtn active={preview === "yekun"} onClick={() => setPreview("yekun")}>
                Yekun ({result.yekun.length})
              </SegmentBtn>
              <SegmentBtn active={preview === "xeta"} onClick={() => setPreview("xeta")}>
                Xəta Loqu ({result.errors.length})
              </SegmentBtn>
              <SegmentBtn active={preview === "dup"} onClick={() => setPreview("dup")}>
                Duplikatlar ({result.dups.length})
              </SegmentBtn>
            </div>

            {preview === "checks" && <ChecksList result={result} />}
            {preview === "yekun" && <YekunTable result={result} />}
            {preview === "xeta" && <XetaTable result={result} />}
            {preview === "dup" && <DupTable result={result} />}
          </section>
        </>
      )}
    </>
  );
}

function DateField({
  label,
  value,
  onChange,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
}) {
  return (
    <label className="flex flex-col gap-1 text-xs font-medium text-slate-500">
      {label}
      <input
        type="date"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className="rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-sm shadow-sm focus:border-brand-400 focus:outline-none focus:ring-2 focus:ring-brand-200"
      />
    </label>
  );
}

function ChecksList({ result }: { result: ReplacementResult }) {
  return (
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
  );
}

const PREVIEW_LIMIT = 200;

function TableShell({ head, children, total }: { head: string[]; children: React.ReactNode; total: number }) {
  return (
    <div className="overflow-x-auto rounded-2xl border border-slate-200 bg-white">
      <table className="min-w-full divide-y divide-slate-200 text-sm">
        <thead className="bg-slate-50">
          <tr className="text-left text-xs font-semibold uppercase tracking-wider text-slate-500">
            {head.map((h) => (
              <th key={h} className="px-3 py-3">
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody className="divide-y divide-slate-100">{children}</tbody>
      </table>
      {total > PREVIEW_LIMIT && (
        <p className="border-t border-slate-100 px-3 py-2 text-xs text-slate-400">
          İlk {PREVIEW_LIMIT} sətir göstərilir (cəmi {total}); tam siyahı export faylındadır.
        </p>
      )}
    </div>
  );
}

function YekunTable({ result }: { result: ReplacementResult }) {
  const rows = result.yekun.slice(0, PREVIEW_LIMIT);
  return (
    <TableShell
      head={["#", "Mənbə", "Sətir", "Kod", "Əvəz edən", "Növ", "Başlama", "Bitmə", "Gün", "Saat", "Qeydlər"]}
      total={result.yekun.length}
    >
      {rows.map((r, i) => (
        <tr key={i} className="hover:bg-slate-50/60">
          <td className="px-3 py-1.5 tabular-nums text-slate-400">{i + 1}</td>
          <td className="px-3 py-1.5 text-slate-700">{r.menbe}</td>
          <td className="px-3 py-1.5 tabular-nums text-slate-500">{r.menbeSetri}</td>
          <td className="px-3 py-1.5 font-medium tabular-nums text-slate-700">{r.evezEdenKod}</td>
          <td className="px-3 py-1.5 text-slate-700">{r.evezEdenAd}</td>
          <td className="px-3 py-1.5 text-slate-600">{r.evezetmeNovu}</td>
          <td className="px-3 py-1.5 tabular-nums text-slate-600">
            {r.dayBased ? fmtDate(r.baslama) : fmtDateTime(r.baslama)}
          </td>
          <td className="px-3 py-1.5 tabular-nums text-slate-600">
            {r.bitme === null ? "açıq" : r.dayBased ? fmtDate(r.bitme) : fmtDateTime(r.bitme)}
          </td>
          <td className="px-3 py-1.5 text-right tabular-nums text-slate-700">{r.dayCount}</td>
          <td className="px-3 py-1.5 text-right tabular-nums text-slate-700">{r.hourCount}</td>
          <td className="max-w-md px-3 py-1.5 text-xs text-slate-500">{r.qeydler.join(" || ")}</td>
        </tr>
      ))}
    </TableShell>
  );
}

function XetaTable({ result }: { result: ReplacementResult }) {
  const rows = result.errors.slice(0, PREVIEW_LIMIT);
  return (
    <TableShell head={["#", "Mənbə", "Sətir", "Xəta növü", "Təfərrüat"]} total={result.errors.length}>
      {rows.map((e, i) => (
        <tr key={i} className="hover:bg-slate-50/60">
          <td className="px-3 py-1.5 tabular-nums text-slate-400">{i + 1}</td>
          <td className="px-3 py-1.5 text-slate-700">{e.menbe}</td>
          <td className="px-3 py-1.5 tabular-nums text-slate-500">{e.setr}</td>
          <td className="px-3 py-1.5">
            <span className="inline-flex rounded-full bg-amber-100 px-2 py-0.5 text-xs font-medium text-amber-800 ring-1 ring-inset ring-amber-200">
              {e.novu}
            </span>
          </td>
          <td className="max-w-xl px-3 py-1.5 text-xs text-slate-600">{e.teferruat}</td>
        </tr>
      ))}
    </TableShell>
  );
}

function DupTable({ result }: { result: ReplacementResult }) {
  const rows = result.dups.slice(0, PREVIEW_LIMIT);
  return (
    <TableShell
      head={["#", "Növ", "Kod", "Əvəz edən", "Başlama", "Bitmə", "Saxlanılan", "Çıxarılan"]}
      total={result.dups.length}
    >
      {rows.map((d, i) => (
        <tr key={i} className="hover:bg-slate-50/60">
          <td className="px-3 py-1.5 tabular-nums text-slate-400">{i + 1}</td>
          <td className="px-3 py-1.5 text-xs text-slate-600">{d.novu}</td>
          <td className="px-3 py-1.5 font-medium tabular-nums text-slate-700">{d.kod}</td>
          <td className="px-3 py-1.5 text-slate-700">{d.ad}</td>
          <td className="px-3 py-1.5 tabular-nums text-slate-600">{fmtDate(d.baslama)}</td>
          <td className="px-3 py-1.5 tabular-nums text-slate-600">{d.bitme === null ? "açıq" : fmtDate(d.bitme)}</td>
          <td className="px-3 py-1.5 text-xs text-slate-600">{d.saxlanilanMenbe}</td>
          <td className="px-3 py-1.5 text-xs text-slate-600">{d.cixarilanMenbe}</td>
        </tr>
      ))}
    </TableShell>
  );
}
