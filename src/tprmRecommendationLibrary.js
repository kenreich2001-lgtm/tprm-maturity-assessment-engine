/**
 * Reusable TPRM recommendation library: typical gaps, standard remediation patterns,
 * and structured language. Used to match low-scoring domains and reduce one-off wording drift.
 */

/** @typedef {{ id: string, label: string, score: number, evidence: string }} DomainEntry */

export const TPRM_DOMAIN_LIBRARY = {
  governance: {
    commonGaps: [
      "Unclear committee ownership and escalation paths for third-party risk.",
      "Policies exist but are not consistently operationalized across BU/legal entities.",
    ],
    templates: [
      "Formalize {domain}: assign named sponsor(s), quarterly steering forum, and documented escalation into enterprise risk committee.",
      "Publish an integrated third-party risk policy set aligned to tier definitions, residual appetite, and exception governance.",
    ],
  },
  inventory: {
    commonGaps: [
      "Incomplete vendor register across entities or legacy acquisitions.",
      "Weak linkage between legal entity, vendor ID, and internal risk records.",
    ],
    templates: [
      "Stand up a golden vendor record with ownership, refresh cadence, and reconciliation across procurement, legal, and risk systems.",
      "Implement automated mismatch detection for dormant vendors and Tier-1 concentration IDs.",
    ],
  },
  segmentation: {
    commonGaps: [
      "Tiering methodology inconsistent or decoupled from residual risk and control intensity.",
    ],
    templates: [
      "Recalibrate tiering criteria against inherent/residual risk and material outsourcing definitions; publish QA sampling.",
      "Link tier decisions explicitly to diligence depth, monitoring cadence, and committee thresholds.",
    ],
  },
  dueDiligence: {
    commonGaps: [
      "Inconsistent diligence depth at onboarding or periodic refresh by tier.",
    ],
    templates: [
      "Deploy tier-standard diligence packages with QA gates; track SLA adherence for onboarding and refresh cycles.",
      "Integrate cyber and resilience assessment outputs into diligence scoring with dispute resolution owners.",
    ],
  },
  contracting: {
    commonGaps: [
      "Contract language misaligned to tier risk or jurisdictional requirements.",
    ],
    templates: [
      "Embed audit/cooperation, subprocessors, data localization, and exit clauses proportionate to tier and sensitivity.",
      "Route Tier-1 contracts through legal/risk standard clauses with breach notification time-bound commitments.",
    ],
  },
  monitoring: {
    commonGaps: [
      "Monitoring KRIs without actionable thresholds or independent validation.",
    ],
    templates: [
      "Define KRIs with tier-specific thresholds, owners, and attestations; automate telemetry feeds where feasible.",
      "Institute breach narratives and ageing discipline for exceptions—avoid passive dashboard greens.",
    ],
  },
  issues: {
    commonGaps: [
      "Issue remediation ageing without SLA transparency or closure validation.",
    ],
    templates: [
      "Implement ageing dashboards with SLA bridges to governance; validate closure with independent QA sampling.",
      "Tie issue reopen triggers to monitoring KRIs and diligence findings.",
    ],
  },
  reporting: {
    commonGaps: [
      "Board/MI packs narrative-heavy without traceability to controls or tiers.",
    ],
    templates: [
      "Produce tier-aware MI one-pagers with concentration metrics and remediation transparency.",
      "Align disclosure narratives to substantiated evidence artifacts—not attestations alone.",
    ],
  },
  technology: {
    commonGaps: [
      "Fragmented workflow tooling and lineage gaps across diligence and monitoring.",
    ],
    templates: [
      "Prioritize workflow orchestration and lineage from diligence → contract → monitoring triggers.",
      "Pilot automation on highest-volume tier segments after inventory and tier accuracy are validated.",
    ],
  },
  regulatory: {
    commonGaps: [
      "Mapping to supervisory expectations incomplete or not operationally tested.",
    ],
    templates: [
      "Build a defensible control-to-expectation map with exam readiness pack and issue evidence trail.",
      "For high regulatory intensity, add jurisdiction-specific addenda and materiality-based testing scope.",
    ],
  },
};

const TIER_CROSS_CUTTING = {
  Low: [
    "Fast-track minimum viable TPRM operating model: complete inventory, defensible tiering, and named owners for the weakest domains.",
    "Unify control and data taxonomies across diligence, contracting, and monitoring so outcomes are testable end-to-end.",
  ],
  Moderate: [
    "Industrialize playbooks and independent QA on diligence and monitoring evidence; time-bound compensating controls for known gaps.",
    "Institute monthly governance forum with metrics; lock issue SLAs to leadership visibility.",
  ],
  Advanced: [
    "Shift toward predictive KRIs, continuous control testing, and automated attestations where materiality supports it.",
    "Sustain scorecard reporting to leadership; fund continuous improvement rather than one-time remediation waves.",
  ],
};

const RAPID_DIAGNOSTIC_ADDON =
  "Scope the diagnostic to Tier-1 (or material) relationships; validate monitoring telemetry and evidence depth within 30 days.";

/**
 * @param {string} template
 * @param {DomainEntry} entry
 */
function applyTemplate(template, entry) {
  return template.replace(/\{domain\}/g, entry.label);
}

/**
 * Slight variation for evidence quality (reduces copy-paste feel without rewriting from scratch).
 * @param {string} line
 * @param {DomainEntry} entry
 */
function evidenceQualify(line, entry) {
  if (entry.evidence === "Weak") {
    return `${line} Prioritize artifact depth and contemporaneous documentation suitable for supervisory or audit challenge.`;
  }
  if (entry.evidence === "Moderate") {
    return `${line} Tighten sampling and QA on documented controls within two refresh cycles.`;
  }
  return line;
}

function dedupePush(arr, seen, text) {
  const key = text.replace(/\s+/g, " ").trim().toLowerCase();
  if (seen.has(key)) return;
  seen.add(key);
  arr.push(text);
}

/**
 * Primary remediation list for engine output, exports, and packs — matched by lowest-scoring domains.
 * @param {DomainEntry[]} sortedByGap ascending gap priority
 * @param {{ tier: string }} band
 * @param {Record<string, unknown>} profile
 */
export function buildRemediationPlan(sortedByGap, band, profile) {
  const out = [];
  const seen = new Set();

  const worst = sortedByGap[0];
  if (worst) {
    const lib = TPRM_DOMAIN_LIBRARY[worst.id];
    if (lib) {
      lib.templates.slice(0, 2).forEach((tpl, i) => {
        const base = applyTemplate(tpl, worst);
        const line = i === 0 ? evidenceQualify(base, worst) : base;
        dedupePush(out, seen, line);
      });
    } else {
      dedupePush(
        out,
        seen,
        `Establish an accountable remediation lane for ${worst.label}: sponsor, 90-day outcomes, and KPIs tied to oversight.`
      );
    }
  }

  sortedByGap
    .filter((e) => e.score <= 3)
    .slice(0, 6)
    .forEach((entry, idx) => {
      if (worst && entry.id === worst.id && idx === 0) return;
      const lib = TPRM_DOMAIN_LIBRARY[entry.id];
      if (!lib || !lib.templates.length) return;
      const tpl = lib.templates[Math.min(idx % lib.templates.length, lib.templates.length - 1)];
      const line = evidenceQualify(applyTemplate(tpl, entry), entry);
      dedupePush(out, seen, line);
    });

  const tierLines = TIER_CROSS_CUTTING[band.tier] || TIER_CROSS_CUTTING.Moderate;
  tierLines.forEach((t) => dedupePush(out, seen, t));

  if (profile.assessmentType === "Rapid diagnostic") {
    dedupePush(out, seen, RAPID_DIAGNOSTIC_ADDON);
  }

  return out.slice(0, 12);
}

/**
 * Multi-line recommendations for AI placeholder when OpenAI is unavailable.
 * @param {object} d — from buildAssessmentDataForNarrative including topGapsForLibrary
 */
export function buildPlaceholderRecommendationLines(d) {
  const lines = [];
  const seen = new Set();
  const gaps = d.topGapsForLibrary || [];

  gaps.slice(0, 5).forEach((g, idx) => {
    const entry = {
      id: g.domainId,
      label: g.domain,
      score: g.score,
      evidence: g.evidence,
    };
    const lib = TPRM_DOMAIN_LIBRARY[g.domainId];
    if (lib?.templates?.length) {
      const tpl = lib.templates[idx % lib.templates.length];
      const line = evidenceQualify(applyTemplate(tpl, entry), entry);
      dedupePush(lines, seen, line);
    }
  });

  if (lines.length < 4 && gaps.length) {
    const entries = gaps.map((g) => ({
      id: g.domainId,
      label: g.domain,
      score: g.score,
      evidence: g.evidence,
    }));
    const sorted = [...entries].sort((a, b) =>
      a.score !== b.score ? a.score - b.score : String(a.evidence).localeCompare(String(b.evidence))
    );
    const tier = d.maturityTier || "Moderate";
    buildRemediationPlan(sorted, { tier }, { assessmentType: d.assessmentTypeHint || "" }).forEach((x) =>
      dedupePush(lines, seen, x)
    );
  }

  dedupePush(
    lines,
    seen,
    "Sequence investment after inventory completeness and tier accuracy—avoid automation spend ahead of foundational data quality."
  );

  return lines.slice(0, 8);
}

/**
 * Short hints for OpenAI user prompt — steer model toward library-aligned actions without pasting the whole library.
 * @param {Array<{ domainId: string, domain: string, score: number, evidence: string }>} topGaps
 */
export function getLibraryHintsForOpenAI(topGaps) {
  const parts = [];
  (topGaps || []).slice(0, 5).forEach((g) => {
    const lib = TPRM_DOMAIN_LIBRARY[g.domainId];
    if (!lib) return;
    const g1 = lib.commonGaps[0];
    const a1 = lib.templates[0]?.replace("{domain}", g.domain) || "";
    parts.push(`- ${g.domain}: typical gap — ${g1} Standard pattern: ${a1}`);
  });
  if (!parts.length) return "";
  return ["Align recommendations where appropriate with these proven TPRM patterns:", ...parts].join("\n");
}
