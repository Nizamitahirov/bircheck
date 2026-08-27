"use client";

import { FormEvent, useEffect, useState } from "react";
import { WorkingDaysTab } from "@/components/WorkingDaysTab";
import { AbsenteeismTab } from "@/components/AbsenteeismTab";
import { CrossCheckTab } from "@/components/CrossCheckTab";

type Tab = "working" | "absent" | "cross";

const TABS: { id: Tab; label: string }[] = [
  { id: "working", label: "Working days check" },
  { id: "absent", label: "Absenteeism" },
  { id: "cross", label: "Cross-checking" },
];

const APP_PASSWORD = "Payroll";
const AUTH_KEY = "bircheck-auth-v1";

export default function Home() {
  const [authed, setAuthed] = useState<boolean | null>(null);
  const [tab, setTab] = useState<Tab>("working");

  useEffect(() => {
    try {
      setAuthed(sessionStorage.getItem(AUTH_KEY) === "1");
    } catch {
      setAuthed(false);
    }
  }, []);

  if (authed === null) return <main className="min-h-screen" />;

  if (!authed) {
    return (
      <PasswordGate
        onSuccess={() => {
          try {
            sessionStorage.setItem(AUTH_KEY, "1");
          } catch {
            /* ignore */
          }
          setAuthed(true);
        }}
      />
    );
  }

  return (
    <main className="min-h-screen w-full px-4 py-10 md:py-16">
      <div className="mx-auto w-full max-w-6xl">
        <header className="mb-8 text-center">
          <h1 className="text-4xl font-extrabold tracking-tight text-slate-900 md:text-5xl">
            Bir<span className="text-brand-600">Check</span>
          </h1>
        </header>

        <nav className="mb-8 flex justify-center">
          <div className="glass inline-flex rounded-2xl p-1 shadow-soft">
            {TABS.map((t) => (
              <button
                key={t.id}
                onClick={() => setTab(t.id)}
                className={`rounded-xl px-5 py-2.5 text-sm font-semibold transition ${
                  tab === t.id
                    ? "bg-brand-600 text-white shadow-sm"
                    : "text-slate-600 hover:bg-slate-50"
                }`}
              >
                {t.label}
              </button>
            ))}
          </div>
        </nav>

        {tab === "working" && <WorkingDaysTab />}
        {tab === "absent" && <AbsenteeismTab />}
        {tab === "cross" && <CrossCheckTab />}

        <footer className="mt-12 text-center text-xs text-slate-400">
          Copyright created by Nizami Tahir
        </footer>
      </div>
    </main>
  );
}

function PasswordGate({ onSuccess }: { onSuccess: () => void }) {
  const [password, setPassword] = useState("");
  const [error, setError] = useState(false);
  const [shake, setShake] = useState(false);

  const submit = (e: FormEvent) => {
    e.preventDefault();
    if (password === APP_PASSWORD) {
      onSuccess();
    } else {
      setError(true);
      setShake(true);
      setPassword("");
      setTimeout(() => setShake(false), 500);
    }
  };

  return (
    <main className="flex min-h-screen w-full items-center justify-center px-4 py-10">
      <div
        className={`glass w-full max-w-sm rounded-3xl p-8 shadow-soft transition ${
          shake ? "animate-[shake_0.4s_ease-in-out]" : ""
        }`}
      >
        <div className="mb-6 text-center">
          <div className="mx-auto mb-4 inline-flex h-14 w-14 items-center justify-center rounded-2xl bg-brand-100 text-brand-600">
            <LockIcon />
          </div>
          <h1 className="text-3xl font-extrabold tracking-tight text-slate-900">
            Bir<span className="text-brand-600">Check</span>
          </h1>
          <p className="mt-1 text-sm text-slate-500">Giriş üçün parol daxil edin</p>
        </div>
        <form onSubmit={submit} className="space-y-3">
          <input
            type="password"
            value={password}
            onChange={(e) => {
              setPassword(e.target.value);
              if (error) setError(false);
            }}
            autoFocus
            placeholder="Parol"
            className={`w-full rounded-xl border px-4 py-3 text-sm shadow-sm transition focus:outline-none focus:ring-2 ${
              error
                ? "border-rose-300 focus:border-rose-400 focus:ring-rose-200"
                : "border-slate-200 focus:border-brand-400 focus:ring-brand-200"
            }`}
          />
          {error && (
            <p className="text-center text-sm text-rose-600">Parol yanlışdır. Yenidən yoxlayın.</p>
          )}
          <button
            type="submit"
            disabled={!password}
            className="inline-flex w-full items-center justify-center gap-2 rounded-xl bg-brand-600 px-5 py-3 font-semibold text-white shadow-soft transition hover:bg-brand-700 disabled:cursor-not-allowed disabled:bg-slate-300 disabled:shadow-none"
          >
            Daxil ol
          </button>
        </form>
        <p className="mt-6 text-center text-xs text-slate-400">
          Copyright created by Nizami Tahir
        </p>
      </div>
    </main>
  );
}

function LockIcon() {
  return (
    <svg
      viewBox="0 0 24 24"
      className="h-7 w-7"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      <rect x="4" y="11" width="16" height="9" rx="2" />
      <path d="M8 11V7a4 4 0 0 1 8 0v4" />
    </svg>
  );
}
