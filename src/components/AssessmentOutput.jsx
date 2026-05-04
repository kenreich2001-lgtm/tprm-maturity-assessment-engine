import SectionCard from "./SectionCard.jsx";

function Pill({ children, tone = "slate" }) {
  const tones = {
    slate: "bg-slate-100 text-slate-700 ring-slate-200",
    amber: "bg-amber-50 text-amber-900 ring-amber-200",
    blue: "bg-blue-50 text-blue-900 ring-blue-200",
    emerald: "bg-emerald-50 text-emerald-900 ring-emerald-200",
  };
  return (
    <span
      className={`inline-flex rounded-md px-2.5 py-1 text-xs font-semibold ring-1 ring-inset ${tones[tone] || tones.slate}`}
    >
      {children}
    </span>
  );
}

export default function AssessmentOutput({ result, visible }) {
  if (!visible || !result) {
    return (
      <SectionCard
        title="Assessment output"
        subtitle="Generated maturity narrative and roadmap — appears after you run the assessment."
      >
        <div className="rounded-lg border border-dashed border-slate-200 bg-slate-50/80 px-6 py-12 text-center">
          <p className="text-sm text-slate-500">
            Complete the profile and domain scores, then click{" "}
            <span className="font-semibold text-navy-900">
              Generate assessment
            </span>{" "}
            to produce the structured output.
          </p>
        </div>
      </SectionCard>
    );
  }

  const { ai } = result;
  const tierTone =
    result.tierColor === "amber"
      ? "amber"
      : result.tierColor === "blue"
        ? "blue"
        : "emerald";

  return (
    <SectionCard
      title="Assessment output"
      subtitle="Narrative assembled from structured scores and profile only — no external facts."
    >
      <p className="mb-6 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-600">
        <span className="font-semibold text-navy-800">Source: </span>
        {ai.meta.source}
      </p>

      <div className="space-y-10">
        {/* Executive narrative */}
        <div className="rounded-xl bg-gradient-to-br from-navy-900 to-slate-800 p-6 text-white shadow-lg">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="flex flex-wrap items-center gap-2">
              <h3 className="text-sm font-semibold uppercase tracking-wider text-white/70">
                Executive narrative
              </h3>
              <Pill tone="slate">
                <span className="text-navy-900">Structured inputs</span>
              </Pill>
            </div>
            <Pill tone={tierTone}>
              {result.maturityBand} · {result.tier}
            </Pill>
          </div>
          <div className="mt-4 space-y-3 text-sm leading-relaxed text-white/90">
            {ai.executiveNarrative.map((para, i) => (
              <p key={i}>{para}</p>
            ))}
          </div>
          <div className="mt-6 flex flex-wrap items-baseline gap-2 border-t border-white/10 pt-4">
            <span className="text-sm text-white/60">Overall maturity score</span>
            <span className="text-3xl font-bold tracking-tight">
              {result.overallScore}
            </span>
            <span className="text-sm text-white/50">/ 5.0 mean domain score</span>
          </div>
        </div>

        {/* Domain findings */}
        <div>
          <h4 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500">
            Domain findings
          </h4>
          <ul className="space-y-3">
            {ai.domainFindings.map((d) => (
              <li
                key={d.domain}
                className="rounded-lg border border-slate-200 bg-white p-4 text-sm shadow-sm"
              >
                <div className="flex flex-wrap items-baseline justify-between gap-2">
                  <span className="font-semibold text-navy-900">{d.domain}</span>
                  <span className="text-xs font-medium text-slate-500">
                    {d.score}/5 · {d.levelLabel} · Evidence: {d.evidence}
                  </span>
                </div>
                <p className="mt-2 leading-relaxed text-slate-700">{d.finding}</p>
              </li>
            ))}
          </ul>
        </div>

        {/* Risk implications */}
        <div>
          <h4 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500">
            Risk implications
          </h4>
          <ul className="space-y-3">
            {ai.riskImplications.map((row, i) => (
              <li
                key={i}
                className="rounded-lg border border-amber-100 bg-amber-50/40 px-4 py-3 text-sm text-amber-950"
              >
                <p className="flex gap-2">
                  <span className="font-bold text-amber-600">!</span>
                  <span>{row.implication}</span>
                </p>
                <p className="mt-2 border-t border-amber-100/80 pt-2 text-xs text-amber-900/80">
                  <span className="font-semibold">Basis: </span>
                  {row.basis}
                </p>
              </li>
            ))}
          </ul>
        </div>

        {/* Prioritized recommendations */}
        <div className="grid gap-8 lg:grid-cols-2">
          <div>
            <h4 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500">
              Prioritized recommendations — gap order
            </h4>
            <ol className="space-y-3">
              {ai.prioritizedRecommendations.gapPriorities.map((g) => (
                <li
                  key={g.rank}
                  className="rounded-lg border border-slate-100 bg-slate-50/80 p-4 text-sm"
                >
                  <div className="flex items-center justify-between gap-2">
                    <span className="font-semibold text-navy-900">
                      {g.rank}. {g.title}
                    </span>
                  </div>
                  <p className="mt-2 text-slate-700">{g.detail}</p>
                  <p className="mt-2 text-xs text-slate-500">{g.basis}</p>
                </li>
              ))}
            </ol>
          </div>
          <div>
            <h4 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500">
              Program-level actions
            </h4>
            <ol className="space-y-3">
              {ai.prioritizedRecommendations.programActions.map((a) => (
                <li
                  key={a.rank}
                  className="rounded-lg border border-accent/15 bg-accent-soft/50 p-4 text-sm"
                >
                  <div className="font-semibold text-navy-900">
                    {a.rank}. {a.title}
                  </div>
                  <p className="mt-2 text-slate-700">{a.detail}</p>
                  <p className="mt-2 text-xs text-slate-500">{a.basis}</p>
                </li>
              ))}
            </ol>
          </div>
        </div>

        {/* Roadmap */}
        <div>
          <h4 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-500">
            Roadmap
          </h4>
          <p className="mb-4 text-sm leading-relaxed text-slate-600">
            {ai.roadmapNarrative.intro}
          </p>
          <div className="grid gap-4 md:grid-cols-3">
            {[
              { label: "30 days", items: ai.roadmapNarrative.horizon30 },
              { label: "60 days", items: ai.roadmapNarrative.horizon60 },
              { label: "90 days", items: ai.roadmapNarrative.horizon90 },
            ].map((col) => (
              <div
                key={col.label}
                className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm"
              >
                <div className="mb-3 text-xs font-bold uppercase tracking-wider text-accent">
                  {col.label}
                </div>
                <ul className="space-y-2 text-sm text-slate-600">
                  {col.items.map((item, i) => (
                    <li key={i} className="flex gap-2">
                      <span className="text-slate-400">•</span>
                      <span>{item}</span>
                    </li>
                  ))}
                </ul>
              </div>
            ))}
          </div>
        </div>

        {/* Audit trace */}
        <details className="rounded-lg border border-slate-200 bg-slate-50/80 p-4 text-sm">
          <summary className="cursor-pointer font-semibold text-navy-900">
            Scoring trace (verbatim engine summary)
          </summary>
          <p className="mt-3 text-slate-700">{result.executiveSummary}</p>
          <ul className="mt-4 divide-y divide-slate-200 rounded-lg border border-slate-200 bg-white">
            {result.domainSummaries.map((d, i) => (
              <li
                key={i}
                className="flex flex-wrap items-start justify-between gap-3 px-4 py-3 text-sm"
              >
                <span className="font-medium text-navy-900">{d.domain}</span>
                <span className="text-slate-600">{d.summary}</span>
              </li>
            ))}
          </ul>
        </details>
      </div>
    </SectionCard>
  );
}
