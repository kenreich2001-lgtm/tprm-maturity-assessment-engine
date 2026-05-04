import SectionCard from "./SectionCard.jsx";
import ScoreBadge from "./ScoreBadge.jsx";
import { DOMAINS, SCORE_LABELS } from "../utils/scoringEngine.js";

const EVIDENCE_OPTIONS = ["Weak", "Moderate", "Strong"];

export default function DomainAssessment({ domainScores, onChange }) {
  const updateDomain = (id, patch) => {
    onChange({
      ...domainScores,
      [id]: { ...domainScores[id], ...patch },
    });
  };

  return (
    <SectionCard
      title="Domain maturity scoring"
      subtitle="Score each capability area from 1 (Initial) to 5 (Optimized). Capture evidence strength and notes."
    >
      <div className="space-y-4">
        {DOMAINS.map((d, idx) => {
          const row = domainScores[d.id] || {
            score: 3,
            evidence: "Moderate",
            notes: "",
          };
          return (
            <div
              key={d.id}
              className="rounded-lg border border-slate-100 bg-slate-50/50 p-4 transition hover:border-slate-200"
            >
              <div className="flex flex-wrap items-start justify-between gap-3 border-b border-slate-100 pb-3">
                <div className="flex items-start gap-3">
                  <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-md bg-white text-xs font-bold text-accent ring-1 ring-slate-200">
                    {idx + 1}
                  </span>
                  <div>
                    <h3 className="font-medium text-navy-900">{d.label}</h3>
                    <p className="mt-0.5 text-xs text-slate-500">
                      Maturity level: {SCORE_LABELS[row.score]}
                    </p>
                  </div>
                </div>
                <ScoreBadge score={row.score} />
              </div>
              <div className="mt-4 grid gap-4 lg:grid-cols-12">
                <div className="lg:col-span-3">
                  <label className="mb-1.5 block text-xs font-medium uppercase tracking-wide text-slate-500">
                    Score (1–5)
                  </label>
                  <select
                    value={row.score}
                    onChange={(e) =>
                      updateDomain(d.id, { score: Number(e.target.value) })
                    }
                    className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm outline-none ring-accent/10 focus:border-accent focus:ring-2 focus:ring-accent/20"
                  >
                    {[1, 2, 3, 4, 5].map((n) => (
                      <option key={n} value={n}>
                        {n} — {SCORE_LABELS[n]}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="lg:col-span-3">
                  <label className="mb-1.5 block text-xs font-medium uppercase tracking-wide text-slate-500">
                    Evidence strength
                  </label>
                  <select
                    value={row.evidence}
                    onChange={(e) =>
                      updateDomain(d.id, { evidence: e.target.value })
                    }
                    className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm outline-none focus:border-accent focus:ring-2 focus:ring-accent/20"
                  >
                    {EVIDENCE_OPTIONS.map((ev) => (
                      <option key={ev} value={ev}>
                        {ev}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="lg:col-span-6">
                  <label className="mb-1.5 block text-xs font-medium uppercase tracking-wide text-slate-500">
                    Notes & evidence references
                  </label>
                  <textarea
                    rows={2}
                    value={row.notes}
                    onChange={(e) =>
                      updateDomain(d.id, { notes: e.target.value })
                    }
                    placeholder="Artifacts reviewed, interviews, system extracts…"
                    className="w-full resize-y rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm outline-none focus:border-accent focus:ring-2 focus:ring-accent/20"
                  />
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </SectionCard>
  );
}
