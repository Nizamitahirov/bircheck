"use client";

import { useCallback, useMemo, useRef, useState } from "react";
import { processWorkbook, downloadWorkbook, type ProcessResult } from "@/lib/processor";

type Status = "idle" | "processing" | "done" | "error";

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [progress, setProgress] = useState(0);
  const [stepText, setStepText] = useState("");
  const [result, setResult] = useState<ProcessResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const reset = () => {
    setFile(null);
    setStatus("idle");
    setProgress(0);
    setStepText("");
    setResult(null);
    setError(null);
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
      const res = await processWorkbook(file, (step, pct) => {
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
    const name = file?.name?.replace(/\.(xlsx|xlsm|xls)$/i, "") ?? "BirCheck";
    downloadWorkbook(result.workbook, `${name}_yoxlanilmis.xlsx`);
  };

  const fileSize = useMemo(() => (file ? formatBytes(file.size) : ""), [file]);

  return (
    <main className="min-h-screen w-full px-4 py-10 md:py-16">
      <div className="mx-auto w-full max-w-3xl">
        <header className="mb-10 text-center">
          <div className="mx-auto mb-4 inline-flex items-center gap-2 rounded-full border border-brand-200/60 bg-white/60 px-3 py-1 text-xs font-medium text-brand-700 shadow-sm">
            <span className="inline-block h-1.5 w-1.5 rounded-full bg-brand-500"></span>
            VBA → Web · Vercel · Client-side
          </div>
          <h1 className="text-4xl font-extrabold tracking-tight text-slate-900 md:text-5xl">
            Bir<span className="text-brand-600">Check</span>
          </h1>
          <p className="mx-auto mt-3 max-w-xl text-slate-600">
            İş günü müqayisəsi və mezuniyyət fərqlərinin avtomatik yoxlanılması. 2 saatlıq
            makrosların yerinə bir neçə saniyə.
          </p>
        </header>

        <section className="glass rounded-3xl p-6 shadow-soft md:p-8">
          {/* Dropzone */}
          <label
            onDragOver={(e) => {
              e.preventDefault();
              setDragging(true);
            }}
            onDragLeave={() => setDragging(false)}
            onDrop={(e) => {
              e.preventDefault();
              setDragging(false);
              const f = e.dataTransfer.files?.[0];
              onPick(f ?? null);
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
                  Tələb olunan sheet-lər: <code className="rounded bg-slate-100 px-1.5">From_Excel_Baza</code>,{" "}
                  <code className="rounded bg-slate-100 px-1.5">From_MHMD</code>,{" "}
                  <code className="rounded bg-slate-100 px-1.5">From_HRB_Otpusk</code>,{" "}
                  <code className="rounded bg-slate-100 px-1.5">Problems</code>
                </p>
              </div>
            )}
          </label>

          {/* Action buttons */}
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

          {/* Progress */}
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

          {/* Error */}
          {error && (
            <div className="mt-6 rounded-xl border border-red-200 bg-red-50 p-4 text-sm text-red-800">
              <p className="font-semibold">Xəta</p>
              <p className="mt-1">{error}</p>
            </div>
          )}

          {/* Result */}
          {result && status === "done" && (
            <div className="mt-6 rounded-2xl border border-emerald-200 bg-emerald-50/60 p-5">
              <div className="flex items-center gap-3">
                <div className="inline-flex h-10 w-10 items-center justify-center rounded-full bg-emerald-500 text-white">
                  <CheckIcon />
                </div>
                <div>
                  <p className="font-semibold text-emerald-900">Hazırdır!</p>
                  <p className="text-sm text-emerald-800">
                    {(result.durationMs / 1000).toFixed(2)} saniyəyə tamamlandı
                  </p>
                </div>
              </div>
              <div className="mt-4 grid grid-cols-3 gap-3">
                <Stat label="İşçi sayı" value={result.totalEmployees} />
                <Stat label="Fərq tapıldı" value={result.rowsKept} accent="brand" />
                <Stat label="Uyğun (silindi)" value={result.rowsRemoved} muted />
              </div>
              <button
                onClick={download}
                className="mt-5 inline-flex w-full items-center justify-center gap-2 rounded-xl bg-emerald-600 px-5 py-3 font-semibold text-white shadow-soft transition hover:bg-emerald-700"
              >
                <DownloadIcon /> Nəticəni endir (.xlsx)
              </button>
            </div>
          )}
        </section>

        {/* Steps */}
        <section className="mt-10 grid gap-4 md:grid-cols-2">
          <InfoCard
            title="Necə işləyir?"
            body={
              <ol className="space-y-2 text-sm text-slate-600">
                {[
                  "Excel faylını yüklə (4 sheet-li).",
                  "From_HRB_Otpusk: J2 başlama, J3 bitmə, K2 ümumi iş günü, I2:I29 bayramlar.",
                  "Sistem MHMD-ni tarix aralığına görə təmizləyir.",
                  "Bazadakı hər işçi üçün iş günü və mezuniyyət hesablanır.",
                  "Fərqi 0 olan sətirlər silinir, fərqlilər mezuniyyət detalları ilə qalır.",
                  "Hazır faylı endir.",
                ].map((t, i) => (
                  <li key={i} className="flex items-start gap-2">
                    <span className="step-dot bg-brand-100 text-brand-700">{i + 1}</span>
                    <span>{t}</span>
                  </li>
                ))}
              </ol>
            }
          />
          <InfoCard
            title="Niyə sürətli?"
            body={
              <ul className="space-y-2 text-sm text-slate-600">
                <li className="flex items-start gap-2">
                  <Dot /> O(n×m) VBA loop-ları O(n) hash-lookup-larına çevrilib.
                </li>
                <li className="flex items-start gap-2">
                  <Dot /> Bütün hesablama brauzerdə işləyir — server gözləməsi yoxdur.
                </li>
                <li className="flex items-start gap-2">
                  <Dot /> Vercel-də static deploy: pulsuz, limitsiz fayl ölçüsü.
                </li>
                <li className="flex items-start gap-2">
                  <Dot /> Faylınız serverə yüklənmir — məxfilik tam qorunur.
                </li>
              </ul>
            }
          />
        </section>

        <footer className="mt-12 text-center text-xs text-slate-400">
          BirCheck · Excel makroslarının web variantı
        </footer>
      </div>
    </main>
  );
}

function Stat({
  label,
  value,
  muted,
  accent,
}: {
  label: string;
  value: number;
  muted?: boolean;
  accent?: "brand";
}) {
  const color = accent === "brand" ? "text-brand-700" : muted ? "text-slate-500" : "text-slate-900";
  return (
    <div className="rounded-xl bg-white/70 p-3 text-center">
      <div className={`text-2xl font-bold tabular-nums ${color}`}>{value}</div>
      <div className="mt-0.5 text-xs text-slate-500">{label}</div>
    </div>
  );
}

function InfoCard({ title, body }: { title: string; body: React.ReactNode }) {
  return (
    <div className="glass rounded-2xl p-5 shadow-sm">
      <h3 className="mb-3 font-semibold text-slate-900">{title}</h3>
      {body}
    </div>
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

function CheckIcon() {
  return (
    <svg viewBox="0 0 24 24" className="h-5 w-5" fill="none" stroke="currentColor" strokeWidth="3">
      <path d="m5 12 5 5L20 7" strokeLinecap="round" strokeLinejoin="round" />
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

function Dot() {
  return <span className="mt-1.5 inline-block h-1.5 w-1.5 flex-shrink-0 rounded-full bg-brand-500" />;
}

function formatBytes(n: number) {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(2)} MB`;
}
