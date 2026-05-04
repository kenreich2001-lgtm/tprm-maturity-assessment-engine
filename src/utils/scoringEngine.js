/**
 * Deterministic "mock AI" assessment from domain scores and client profile.
 */

export const DOMAINS = [
  { id: "governance", label: "Governance and ownership" },
  { id: "inventory", label: "Third-party inventory" },
  { id: "segmentation", label: "Risk segmentation" },
  { id: "dueDiligence", label: "Due diligence" },
  { id: "contracting", label: "Contracting and control requirements" },
  { id: "monitoring", label: "Ongoing monitoring" },
  { id: "issues", label: "Issue management" },
  { id: "reporting", label: "Reporting and oversight" },
  { id: "technology", label: "Technology and data enablement" },
  { id: "regulatory", label: "Regulatory alignment" },
];

export const SCORE_LABELS = {
  1: "Initial / Ad hoc",
  2: "Developing",
  3: "Defined",
  4: "Managed",
  5: "Optimized",
};

export const EVIDENCE_WEIGHT = {
  Weak: 0,
  Moderate: 0.15,
  Strong: 0.25,
};

function tierFromAverage(avg) {
  if (avg < 2.5) return { tier: "Low", band: "Low maturity", color: "amber" };
  if (avg <= 3.5)
    return { tier: "Moderate", band: "Moderate maturity", color: "blue" };
  return { tier: "Advanced", band: "Advanced maturity", color: "emerald" };
}

function gapPriority(domainScore, evidence) {
  const ew = EVIDENCE_WEIGHT[evidence] ?? 0;
  return domainScore - ew;
}

export function generateAssessment(clientProfile, domainScores) {
  const entries = DOMAINS.map((d) => {
    const row = domainScores[d.id] || {
      score: 3,
      evidence: "Moderate",
      notes: "",
    };
    const score = Number(row.score);
    const evidence = row.evidence || "Moderate";
    const weighted = gapPriority(score, evidence);
    return {
      ...d,
      score,
      evidence,
      notes: row.notes || "",
      labelScore: SCORE_LABELS[score] || "",
      weighted,
    };
  });

  const avg =
    entries.reduce((s, e) => s + e.score, 0) / Math.max(entries.length, 1);
  const band = tierFromAverage(avg);

  const sortedGaps = [...entries].sort((a, b) => {
    if (a.score !== b.score) return a.score - b.score;
    return a.weighted - b.weighted;
  });

  const topGaps = sortedGaps.slice(0, 5).map((e) => ({
    domain: e.label,
    score: e.score,
    evidence: e.evidence,
    rationale: buildGapRationale(e),
  }));

  const themes = deriveRiskThemes(entries, clientProfile);
  const remediation = deriveRemediation(sortedGaps[0], band, clientProfile);
  const roadmap = buildRoadmap(band, sortedGaps, clientProfile);
  const executiveSummary = buildExecutiveSummary(
    clientProfile,
    band,
    avg,
    entries,
    themes
  );

  return {
    overallScore: Math.round(avg * 10) / 10,
    maturityBand: band.band,
    tier: band.tier,
    tierColor: band.color,
    domainSummaries: entries.map((e) => ({
      domain: e.label,
      score: e.score,
      evidence: e.evidence,
      summary: `${e.label} is at ${SCORE_LABELS[e.score]} (${e.score}/5) with ${e.evidence.toLowerCase()} evidence strength.`,
    })),
    topGaps,
    riskThemes: themes,
    remediation,
    roadmap,
    executiveSummary,
  };
}

function buildGapRationale(e) {
  if (e.score <= 2)
    return "Score and evidence indicate foundational gaps; prioritize governance and repeatable processes.";
  if (e.score === 3)
    return "Defined but inconsistent execution—tighten ownership, metrics, and assurance.";
  return "Relative strength area—use as anchor to uplift adjacent domains.";
}

function deriveRiskThemes(entries, profile) {
  const themes = [];
  const low = entries.filter((e) => e.score <= 2);
  const reg = profile.regulatoryIntensity;
  const vol = profile.thirdPartyVolume;

  if (low.some((e) => e.id === "governance"))
    themes.push(
      "Governance fragmentation increases approval drift and inconsistent risk appetite application."
    );
  if (low.some((e) => e.id === "inventory" || e.id === "segmentation"))
    themes.push(
      "Incomplete inventory or segmentation elevates concentration and residual risk in the third-party portfolio."
    );
  if (low.some((e) => e.id === "monitoring" || e.id === "issues"))
    themes.push(
      "Monitoring and issue closure gaps create latent operational and compliance exposure between diligence cycles."
    );
  if (low.some((e) => e.id === "technology"))
    themes.push(
      "Technology and data gaps reduce traceability and extend time-to-detect for vendor-related incidents."
    );
  if (
    (reg === "High" || reg === "Very High") &&
    entries.find((e) => e.id === "regulatory")?.score <= 3
  ) {
    themes.push(
      "Regulatory intensity outpaces documented alignment—expect supervisory scrutiny on third-party oversight."
    );
  }
  if (
    (vol === "High" || vol === "Very High") &&
    entries.some((e) => e.score <= 3 && e.id === "dueDiligence")
  ) {
    themes.push(
      "High third-party volume strains due diligence throughput—queue risk and inconsistent depth likely."
    );
  }

  if (themes.length === 0) {
    themes.push(
      "Portfolio-level concentration and control efficacy remain the primary enterprise themes—validate with sampling."
    );
  }

  return themes.slice(0, 6);
}

function deriveRemediation(worstGap, band, profile) {
  const actions = [];
  if (worstGap) {
    actions.push(
      `Establish an accountable remediation lane for ${worstGap.label}: executive sponsor, 90-day control targets, and measurable KPIs.`
    );
  }
  actions.push(
    "Stand up a unified risk taxonomy across diligence, contracting, and monitoring so scoring translates into enforcement."
  );
  if (band.tier === "Low") {
    actions.push(
      "Fast-track minimum viable governance: inventory completeness, critical-tier identification, and monitoring triggers."
    );
  } else if (band.tier === "Moderate") {
    actions.push(
      "Industrialize playbooks and QA on diligence packs; tie exceptions to compensating controls with expiry dates."
    );
  } else {
    actions.push(
      "Shift from compliance milestones to predictive indicators—KRIs, automated attestations, and continuous control testing where material."
    );
  }
  if (profile.assessmentType === "Rapid diagnostic") {
    actions.push(
      "Scope rapid diagnostic to top critical Tier-1 relationships and validate monitoring telemetry within 30 days."
    );
  }
  return actions;
}

function buildRoadmap(band, sortedByGap, profile) {
  const worst = sortedByGap[0]?.label || "priority domains";
  const d30 = [
    `Confirm ownership and success metrics for ${worst}; baseline current-state artifacts.`,
    "Publish interim third-party policy clarifications and escalation paths.",
    "Pilot enhanced evidence capture for one critical tier segment.",
  ];
  const d60 = [
    "Deploy standardized diligence and monitoring checklists tied to risk tier.",
    "Integrate issue tracking with SLA dashboards; begin monthly governance forum.",
    "Automate inventory refresh triggers where source systems exist.",
  ];
  const d90 = [
    "Run first independent QA sample across diligence and monitoring outputs.",
    "Publish maturity scorecard to leadership; lock FY roadmap funding.",
    "Expand tooling pilots if technology domain remains below target.",
  ];
  if (band.tier === "Advanced") {
    d30.unshift(
      "Benchmark KRIs against peers; refine Tier-0 concentration triggers."
    );
  }
  if (profile.companySize === "Global Enterprise") {
    d60.push("Align regional operating models to global minimum standards with local lawful exceptions register.");
  }
  return { d30, d60, d90 };
}

function buildExecutiveSummary(profile, band, avg, entries, themes) {
  const industry = profile.industry || "the sector";
  const reg = profile.regulatoryIntensity || "current";
  const footprint = profile.geographicFootprint || "operating";

  return (
    `Overall maturity is assessed at ${band.band.toLowerCase()} (mean domain score ${avg.toFixed(1)}/5). ` +
    `For a ${profile.companySize?.toLowerCase() || "mid-market"} organization in ${industry} with ${reg.toLowerCase()} regulatory intensity ` +
    `and ${footprint.toLowerCase()} footprint, the program ${avg < 3 ? "requires foundational strengthening and clearer accountability" : avg < 4 ? "shows definitional strength but needs consistent execution and assurance" : "demonstrates strong discipline with opportunities to industrialize analytics and predictive monitoring"}. ` +
    `Priority attention should flow to the weakest domains while ensuring evidence depth keeps pace with third-party scale. ` +
    `Key themes: ${themes.slice(0, 2).join(" ")}`
  );
}
