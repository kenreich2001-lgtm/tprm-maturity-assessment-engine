/**
 * AI-style narrative assembly from structured inputs only.
 * All statements must trace to: client profile fields, numeric domain scores,
 * evidence selections, user notes (verbatim when present), and derived
 * assessment fields (means, bands, orderings). No external or invented facts.
 */

import { DOMAINS, SCORE_LABELS } from "./scoringEngine.js";

/**
 * @param {object} clientProfile
 * @param {object} domainScores - raw state { [id]: { score, evidence, notes } }
 * @param {object} assessment - output of generateAssessment()
 */
export function buildAINarrative(clientProfile, domainScores, assessment) {
  const entries = DOMAINS.map((d) => {
    const row = domainScores[d.id] || {};
    const score = Number(row.score ?? 3);
    const evidence = row.evidence || "Moderate";
    const notes = typeof row.notes === "string" ? row.notes.trim() : "";
    return {
      id: d.id,
      label: d.label,
      score,
      evidence,
      notes,
      levelLabel: SCORE_LABELS[score] || String(score),
    };
  });

  const sortedAsc = [...entries].sort((a, b) =>
    a.score !== b.score ? a.score - b.score : a.label.localeCompare(b.label)
  );
  const sortedDesc = [...entries].sort((a, b) =>
    a.score !== b.score ? b.score - a.score : a.label.localeCompare(b.label)
  );
  const weakest = sortedAsc.slice(0, Math.min(3, sortedAsc.length));
  const strongest = sortedDesc.slice(0, Math.min(3, sortedDesc.length));

  return {
    meta: {
      source:
        "Generated only from fields in this assessment (profile dropdowns, domain scores 1–5, evidence strength, optional notes). No external data.",
    },
    executiveNarrative: buildExecutiveNarrative(
      clientProfile,
      assessment,
      entries,
      weakest,
      strongest
    ),
    domainFindings: entries.map((e) => domainFinding(e)),
    riskImplications: buildRiskImplications(clientProfile, entries, assessment),
    prioritizedRecommendations: buildPrioritizedRecommendations(assessment),
    roadmapNarrative: buildRoadmapNarrative(
      clientProfile,
      assessment,
      sortedAsc[0]
    ),
  };
}

function buildExecutiveNarrative(profile, assessment, entries, weakest, strongest) {
  const paras = [];

  paras.push(
    `This narrative reflects the selected assessment type (${profile.assessmentType}), industry (${profile.industry}), company size (${profile.companySize}), regulatory intensity (${profile.regulatoryIntensity}), third-party volume (${profile.thirdPartyVolume}), and geographic footprint (${profile.geographicFootprint}). The computed mean domain score is ${assessment.overallScore} on a 1–5 scale, placing overall maturity in the "${assessment.maturityBand}" band (${assessment.tier} tier per internal scoring rules).`
  );

  paras.push(
    `Across the ten domains, recorded scores range from ${Math.min(...entries.map((e) => e.score))}/5 to ${Math.max(...entries.map((e) => e.score))}/5. The lowest recorded scores appear in: ${weakest.map((w) => `${w.label} (${w.score}/5)`).join(", ")}. The highest recorded scores appear in: ${strongest.map((s) => `${s.label} (${s.score}/5)`).join(", ")}.`
  );

  paras.push(
    `Prioritization should follow the ordering implied by recorded scores and evidence strength (see prioritized recommendations). Themes listed under risk implications are conditional statements tied only to those inputs—not independent factual claims about the organization.`
  );

  return paras;
}

function domainFinding(e) {
  const strengthBand =
    e.score <= 2 ? "priority attention" : e.score === 3 ? "stabilization" : "sustain or refine";

  let finding = `${e.label}: recorded maturity score ${e.score}/5 (${e.levelLabel}), with evidence strength rated ${e.evidence}. Based solely on these inputs, this domain maps to ${strengthBand} within the assessment model.`;

  if (e.notes) {
    finding += ` Assessor notes (verbatim): "${e.notes}"`;
  } else {
    finding += ` No free-text notes were entered for this domain.`;
  }

  return {
    domain: e.label,
    score: e.score,
    levelLabel: e.levelLabel,
    evidence: e.evidence,
    finding,
    hasNotes: Boolean(e.notes),
  };
}

function buildRiskImplications(profile, entries, assessment) {
  /** Each item ties to profile fields, scores, evidence, or engine rules — no external facts. */
  const items = [];

  items.push({
    implication:
      "Downstream conclusions should be read as conditional on the profile and scores entered in this session—not as verified statements about real-world incidents or supervisory outcomes.",
    basis: `Profile: regulatory intensity = ${profile.regulatoryIntensity}; third-party volume = ${profile.thirdPartyVolume}; geographic footprint = ${profile.geographicFootprint}. Mean domain score = ${assessment.overallScore}.`,
  });

  const low = entries.filter((e) => e.score <= 2);
  if (low.length > 0) {
    items.push({
      implication:
        "Recorded scores at or below 2/5 indicate comparatively weak structured capability in the named domains within this rubric; sequencing should reflect these numeric gaps.",
      basis: `Domains with score ≤2: ${low.map((e) => `${e.label} (${e.score}/5)`).join("; ")}.`,
    });
  }

  const weakEv = entries.filter((e) => e.evidence === "Weak");
  if (weakEv.length > 0) {
    items.push({
      implication:
        "Evidence strength marked Weak lowers assurance implied by the numeric score alone for those domains.",
      basis: `Domains with Weak evidence: ${weakEv.map((e) => `${e.label} (${e.score}/5)`).join("; ")}.`,
    });
  }

  assessment.riskThemes.forEach((theme, i) => {
    items.push({
      implication: theme,
      basis: `Theme ${i + 1} from scoring-engine rules applied to this profile and recorded domain scores (conditional implication, not an independent factual assertion).`,
    });
  });

  return items;
}

function buildPrioritizedRecommendations(assessment) {
  const gapPriorities = assessment.topGaps.map((g, i) => ({
    rank: i + 1,
    kind: "gap",
    title: g.domain,
    detail: `Score ${g.score}/5 · ${g.evidence} evidence. ${g.rationale}`,
    basis:
      "Ranked by ascending domain score and evidence weighting in the assessment engine.",
  }));

  const programActions = assessment.remediation.map((line, i) => ({
    rank: i + 1,
    kind: "program",
    title: `Program-level action ${i + 1}`,
    detail: line,
    basis:
      "Derived from lowest-scoring domain, maturity tier, and assessment type in the profile.",
  }));

  return { gapPriorities, programActions };
}

function buildRoadmapNarrative(profile, assessment, lowestDomain) {
  const focus = lowestDomain
    ? `${lowestDomain.label} (${lowestDomain.score}/5)`
    : "lowest-scoring domains";

  return {
    intro: `Roadmap items below are generated templates keyed to assessment type (${profile.assessmentType}) and lowest recorded domain focus (${focus}). They do not prescribe legal obligations or project dates—only structured next steps consistent with inputs.`,
    horizon30: assessment.roadmap.d30,
    horizon60: assessment.roadmap.d60,
    horizon90: assessment.roadmap.d90,
  };
}
