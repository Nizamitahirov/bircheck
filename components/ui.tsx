"use client";

import { useRef, useState } from "react";

export function Spinner() {
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

export function PlayIcon() {
  return (
    <svg viewBox="0 0 24 24" className="h-4 w-4" fill="currentColor">
      <path d="M8 5v14l11-7z" />
    </svg>
  );
}

export function UploadIcon() {
  return (
    <svg
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

export function DownloadIcon() {
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

export function formatBytes(n: number) {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(2)} MB`;
}

export function formatNum(v: unknown): string {
  if (v === null || v === undefined || v === "") return "";
  const n = Number(v);
  if (!Number.isFinite(n)) return String(v);
  return Number.isInteger(n) ? String(n) : n.toFixed(2);
}

export function Stat({
  label,
  value,
  accent,
  tone,
}: {
  label: string;
  value: number | string;
  accent?: "brand";
  tone?: "amber" | "rose" | "emerald";
}) {
  let color = "text-slate-900";
  if (accent === "brand") color = "text-brand-700";
  else if (tone === "amber") color = "text-amber-700";
  else if (tone === "rose") color = "text-rose-700";
  else if (tone === "emerald") color = "text-emerald-700";
  return (
    <div className="glass rounded-2xl p-4 text-center shadow-sm">
      <div className={`text-3xl font-bold tabular-nums ${color}`}>{value}</div>
      <div className="mt-1 text-xs uppercase tracking-wide text-slate-500">{label}</div>
    </div>
  );
}

export function SegmentBtn({
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

export function ProgressBar({ pct, label }: { pct: number; label: string }) {
  return (
    <div>
      <div className="mb-2 flex items-center justify-between text-sm">
        <span className="font-medium text-slate-700">{label || "Hazır"}</span>
        <span className="tabular-nums text-slate-500">{pct}%</span>
      </div>
      <div className="h-2 w-full overflow-hidden rounded-full bg-slate-200">
        <div
          className="h-full rounded-full bg-gradient-to-r from-brand-400 to-brand-600 transition-[width] duration-300"
          style={{ width: `${pct}%` }}
        />
      </div>
    </div>
  );
}

export function ErrorBox({ message }: { message: string }) {
  return (
    <div className="rounded-xl border border-red-200 bg-red-50 p-4 text-sm text-red-800">
      <p className="font-semibold">Xəta</p>
      <p className="mt-1">{message}</p>
    </div>
  );
}

interface FileDropProps {
  file: File | null;
  onPick: (file: File | null) => void;
  hint: React.ReactNode;
}

export function FileDrop({ file, onPick, hint }: FileDropProps) {
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const pick = (f: File | null) => {
    if (!f) return;
    if (!/\.(xlsx|xlsm|xls)$/i.test(f.name)) return;
    onPick(f);
  };

  return (
    <label
      onDragOver={(e) => {
        e.preventDefault();
        setDragging(true);
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragging(false);
        pick(e.dataTransfer.files?.[0] ?? null);
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
        onChange={(e) => pick(e.target.files?.[0] ?? null)}
      />
      <div className="mb-3 inline-flex h-14 w-14 items-center justify-center rounded-2xl bg-brand-100 text-brand-600 transition group-hover:scale-105">
        <UploadIcon />
      </div>
      {file ? (
        <div className="space-y-1">
          <p className="font-medium text-slate-900">{file.name}</p>
          <p className="text-sm text-slate-500">{formatBytes(file.size)}</p>
          <p className="mt-2 text-xs text-brand-600">Başqa fayl seçmək üçün klikləyin</p>
        </div>
      ) : (
        <div className="space-y-1">
          <p className="font-semibold text-slate-800">
            Excel faylını bura sürüşdürün və ya seçin
          </p>
          <div className="text-sm text-slate-500">{hint}</div>
        </div>
      )}
    </label>
  );
}
