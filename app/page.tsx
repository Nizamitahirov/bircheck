"use client";

import { useState } from "react";
import { WorkingDaysTab } from "@/components/WorkingDaysTab";
import { AbsenteeismTab } from "@/components/AbsenteeismTab";
import { ReplacementTab } from "@/components/ReplacementTab";
import { VacationWorkyearTab } from "@/components/VacationWorkyearTab";

type Tab = "working" | "absent" | "replacement" | "workyear";

const TABS: { id: Tab; label: string }[] = [
  { id: "working", label: "Working days check" },
  { id: "absent", label: "Absenteeism" },
  { id: "replacement", label: "Replacement check" },
  { id: "workyear", label: "Vacation by Workyear" },
];

export default function Home() {
  const [tab, setTab] = useState<Tab>("working");

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

        {tab === "working" ? (
          <WorkingDaysTab />
        ) : tab === "absent" ? (
          <AbsenteeismTab />
        ) : tab === "replacement" ? (
          <ReplacementTab />
        ) : (
          <VacationWorkyearTab />
        )}

        <footer className="mt-12 text-center text-xs text-slate-400">
          Copyright created by Nizami Tahir
        </footer>
      </div>
    </main>
  );
}
