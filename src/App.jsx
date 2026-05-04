import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import pptxgen from "pptxgenjs";
import {
  buildPlaceholderRecommendationLines,
  buildRemediationPlan,
  getLibraryHintsForOpenAI,
  TPRM_DOMAIN_LIBRARY,
} from "./tprmRecommendationLibrary.js";

// --- Domain catalog & scoring engine (embedded) ---

const DOMAINS = [
  {
    id: "governance",
    label: "Governance & ownership",
    description:
      "Board / committee oversight, policy ownership, and accountability for third-party risk appetite.",
  },
  {
    id: "inventory",
    label: "Third-party inventory",
    description:
      "Completeness and accuracy of the third-party register across legal entities and systems.",
  },
  {
    id: "segmentation",
    label: "Risk segmentation",
    description:
      "Tiering / classification methodology and linkage to residual risk and control intensity.",
  },
  {
    id: "dueDiligence",
    label: "Due diligence",
    description:
      "Scope, depth, and QA of diligence for onboarding and periodic refresh by materiality.",
  },
  {
    id: "contracting",
    label: "Contracting & controls",
    description:
      "Contractual security, operational resilience, and control language aligned to risk tier.",
  },
  {
    id: "monitoring",
    label: "Ongoing monitoring",
    description:
      "Monitoring coverage, KRIs, attestations, and escalation when triggers breach thresholds.",
  },
  {
    id: "issues",
    label: "Issue management",
    description:
      "Finding lifecycle, remediation SLAs, validation of closure, and linkage to governance.",
  },
  {
    id: "reporting",
    label: "Reporting & oversight",
    description:
      "Management and board reporting, MI quality, and audit trail for third-party decisions.",
  },
  {
    id: "technology",
    label: "Technology & data",
    description:
      "Workflow tooling, integrations, data lineage, and automation for monitoring at scale.",
  },
  {
    id: "regulatory",
    label: "Regulatory alignment",
    description:
      "Mapping to supervisory expectations, jurisdictional nuance, and demonstrable compliance.",
  },
];

const SCORE_LABELS = {
  1: "Initial / Ad hoc",
  2: "Developing",
  3: "Defined",
  4: "Managed",
  5: "Optimized",
};

/** Industry benchmark per domain (static reference overlay) */
const BENCHMARK_SCORE = 3.5;

const EVIDENCE_WEIGHT = { Weak: 0, Moderate: 0.15, Strong: 0.25 };

const PROFILE_OPTIONS = {
  industry: [
    "Financial Services",
    "Healthcare",
    "Technology",
    "Manufacturing",
    "Retail",
    "Energy",
    "Public Sector",
    "Other",
  ],
  companySize: ["Small", "Mid-Market", "Enterprise", "Global Enterprise"],
  regulatoryIntensity: ["Low", "Medium", "High", "Very High"],
  thirdPartyVolume: ["Low", "Medium", "High", "Very High"],
  geographicFootprint: ["Domestic", "Regional", "Global"],
};

const ASSESSMENT_TYPES = [
  "Full maturity assessment",
  "Rapid diagnostic",
  "Targeted domain review",
];

function tierFromAverage(avg) {
  if (avg < 2.5) return { tier: "Low", band: "Low maturity" };
  if (avg <= 3.5) return { tier: "Moderate", band: "Moderate maturity" };
  return { tier: "Advanced", band: "Advanced maturity" };
}

function gapPriority(score, evidence) {
  return score - (EVIDENCE_WEIGHT[evidence] ?? 0);
}

function buildGapRationale(e) {
  if (e.score <= 2)
    return "Foundational gaps—prioritize repeatable processes and clear ownership.";
  if (e.score === 3)
    return "Defined but inconsistent—tighten metrics, assurance, and evidence depth.";
  return "Relative strength—use as anchor to uplift adjacent domains.";
}

function deriveRiskThemes(entries, profile) {
  const themes = [];
  const reg = profile.regulatoryIntensity;
  const vol = profile.thirdPartyVolume;

  if (entries.some((e) => e.id === "governance" && e.score <= 2))
    themes.push(
      "Governance fragmentation may drive inconsistent risk appetite application."
    );
  if (
    entries.some(
      (e) => (e.id === "inventory" || e.id === "segmentation") && e.score <= 2
    )
  )
    themes.push(
      "Inventory or segmentation gaps elevate concentration and residual portfolio risk."
    );
  if (
    entries.some(
      (e) => (e.id === "monitoring" || e.id === "issues") && e.score <= 2
    )
  )
    themes.push(
      "Monitoring and issue closure gaps increase exposure between diligence cycles."
    );
  if (entries.some((e) => e.id === "technology" && e.score <= 2))
    themes.push(
      "Technology gaps reduce traceability and extend time-to-detect for vendor issues."
    );
  if (
    (reg === "High" || reg === "Very High") &&
    entries.find((e) => e.id === "regulatory")?.score <= 3
  ) {
    themes.push(
      "Regulatory intensity may outpace documented alignment—focus supervisory readiness."
    );
  }
  if (
    (vol === "High" || vol === "Very High") &&
    entries.some((e) => e.score <= 3 && e.id === "dueDiligence")
  ) {
    themes.push(
      "High third-party volume strains diligence throughput and depth consistency."
    );
  }
  if (themes.length === 0) {
    themes.push(
      "Portfolio concentration and control efficacy remain primary themes—validate with sampling."
    );
  }
  return themes.slice(0, 6);
}

function buildRoadmap(band, sortedByGap, profile) {
  const worst = sortedByGap[0]?.label || "priority domains";
  const d30 = [
    `Confirm ownership and metrics for ${worst}; baseline current-state artifacts.`,
    "Publish interim policy clarifications and escalation paths.",
    "Pilot enhanced evidence capture for one critical tier segment.",
  ];
  const d60 = [
    "Deploy standardized diligence and monitoring checklists by tier.",
    "Integrate issue SLAs with dashboards; monthly governance forum.",
    "Automate inventory refresh triggers where systems allow.",
  ];
  const d90 = [
    "Independent QA sample across diligence and monitoring outputs.",
    "Publish maturity scorecard to leadership; lock roadmap funding.",
    "Expand tooling pilots if technology remains below target.",
  ];
  if (band.tier === "Advanced") {
    d30.unshift("Benchmark KRIs; refine Tier-0 concentration triggers.");
  }
  if (profile.companySize === "Global Enterprise") {
    d60.push(
      "Align regional models to global minimums with lawful exceptions register."
    );
  }
  return { d30, d60, d90 };
}

function buildExecutiveSummary(profile, band, avg, themes) {
  const industry = profile.industry || "the sector";
  const reg = profile.regulatoryIntensity || "current";
  const footprint = profile.geographicFootprint || "operating";
  const posture =
    avg < 3
      ? "requires foundational strengthening and clearer accountability"
      : avg < 4
        ? "shows definitional strength but needs consistent execution and assurance"
        : "demonstrates strong discipline with room to industrialize analytics";
  return (
    `Overall maturity is ${band.band.toLowerCase()} (mean domain score ${avg.toFixed(1)}/5). ` +
    `For a ${(profile.companySize || "mid-market").toLowerCase()} organization in ${industry} with ${reg.toLowerCase()} regulatory intensity ` +
    `and ${footprint.toLowerCase()} footprint, the program ${posture}. ` +
    `Priority flows to weakest domains while evidence depth scales with third-party volume. ` +
    `Themes: ${themes.slice(0, 2).join(" ")}`
  );
}

function buildAssessment(profile, domainRows) {
  const entries = DOMAINS.map((d) => {
    const row = domainRows[d.id] || {};
    const score = Number(row.score ?? 3);
    const evidence = row.evidence || "Moderate";
    return {
      ...d,
      score,
      evidence,
      notes: (row.notes || "").trim(),
      weighted: gapPriority(score, evidence),
    };
  });

  const avg =
    entries.reduce((s, e) => s + e.score, 0) / Math.max(entries.length, 1);
  const band = tierFromAverage(avg);
  const sortedGaps = [...entries].sort((a, b) => {
    if (a.score !== b.score) return a.score - b.score;
    return a.weighted - b.weighted;
  });

  const topGapsAll = sortedGaps.map((e) => ({
    domain: e.label,
    domainId: e.id,
    score: e.score,
    evidence: e.evidence,
    rationale: buildGapRationale(e),
  }));

  const top3 = topGapsAll.slice(0, 3);
  const weakCount = entries.filter((e) => e.score <= 2).length;

  const themes = deriveRiskThemes(entries, profile);
  const remediation = buildRemediationPlan(sortedGaps, band, profile);
  const roadmap = buildRoadmap(band, sortedGaps, profile);
  const executiveSummary = buildExecutiveSummary(profile, band, avg, themes);

  return {
    overallScore: Math.round(avg * 10) / 10,
    maturityBand: band.band,
    maturityTier: band.tier,
    weakCount,
    top3,
    topGapsAll,
    themes,
    remediation,
    roadmap,
    executiveSummary,
    entries,
    assessmentType: profile.assessmentType ?? "",
  };
}

/** Structured input for AI / placeholder narrative generation */
function buildAssessmentDataForNarrative(assessment) {
  const entries = assessment.entries;
  const weakDomains = entries.filter((e) => e.score <= 2);
  const weakList = weakDomains.map((e) => e.label).join("; ") || "None";
  let nWeakEv = 0;
  let nModEv = 0;
  let nStrongEv = 0;
  entries.forEach((e) => {
    if (e.evidence === "Weak") nWeakEv++;
    else if (e.evidence === "Moderate") nModEv++;
    else nStrongEv++;
  });
  const evidenceSummary = `${nWeakEv} Weak, ${nModEv} Moderate, ${nStrongEv} Strong (${entries.length} domains)`;
  const domainScoresLine = entries.map((e) => `${e.label}: ${e.score}/5 (${e.evidence})`).join("\n");
  const topGapsSummary = assessment.topGapsAll
    .slice(0, 5)
    .map((g) => `${g.domain} ${g.score}/5 (${g.evidence})`)
    .join("; ");
  const topGapsForLibrary = assessment.topGapsAll.slice(0, 5).map((g) => ({
    domainId: g.domainId,
    domain: g.domain,
    score: g.score,
    evidence: g.evidence,
  }));
  return {
    overallScore: assessment.overallScore,
    maturityLevel: assessment.maturityBand,
    maturityTier: assessment.maturityTier,
    weakDomainsList: weakList,
    evidenceSummary,
    domainScoresLine,
    topGapsSummary,
    themesLine: (assessment.themes || []).join("; "),
    topGapsForLibrary,
    assessmentTypeHint: assessment.assessmentType ?? "",
  };
}

function generatePlaceholderNarrative(d) {
  const score = d.overallScore;
  const weak = d.weakDomainsList && d.weakDomainsList !== "None" ? d.weakDomainsList : "limited pockets";
  let exec;
  let risk;
  if (score < 3) {
    exec = `The assessment indicates significant maturity gaps across the third-party risk lifecycle (mean domain score ${score}/5). Leadership should assume supervisory and operational exposure where controls are thin, evidence is uneven, and tiering may not reflect residual concentration. Near-term priority is to stabilize governance, inventory integrity, and monitoring triggers before expanding scope or automation spend. Cross-functional sponsorship is required so remediation does not stall in siloed workstreams.`;
    risk = `Risk exposure is elevated: domains at score ≤2 include ${weak}. Evidence strength is summarized as ${d.evidenceSummary}. Control weaknesses typically manifest as delayed detection of vendor failures, inconsistent diligence depth at refresh, and audit trails that may not withstand challenge. Until foundational artifacts and ownership are strengthened, residual risk from critical third parties may exceed risk appetite in practice—even where policy appears adequate on paper.`;
  } else if (score < 4) {
    exec = `The program shows developing capabilities (mean ${score}/5) with definitional assets in place but uneven execution and assurance. Several domains likely constrain scale—particularly where evidence is moderate or weak relative to materiality. The organization should industrialize playbooks, tighten QA on diligence and monitoring outputs, and align investment to tiered exposure rather than uniform uplift.`;
    risk = `Risk is moderate and unevenly distributed: domains requiring attention include ${weak === "None" ? "limited low-score lanes only; validate tiering" : weak}. With ${d.evidenceSummary}, assurance quality may lag policy intent—creating windows where attestations outrun artifacts. Key implications include reactive oversight, limited predictive MI, and potential regulatory scrutiny where Tier-1 concentration meets thin documentation.`;
  } else {
    exec = `Overall maturity presents a strong foundation (mean ${score}/5) with opportunity to industrialize analytics and continuous monitoring where material. Focus shifts from wholesale redesign to targeted uplift in any residual weak domains, sustained evidence cadence, and hardening automation before vendor volume scales further.`;
    risk = `Residual risk is comparatively contained; remaining exposure clusters where evidence or operating cadence still trails tier expectations (${d.evidenceSummary}). The primary risk is complacency—assuming green dashboards without independent sampling—or underinvesting in data lineage as models consolidate.`;
  }
  const rec = buildPlaceholderRecommendationLines(d).join("\n");
  return {
    executiveSummary: exec,
    riskNarrative: risk,
    recommendationsText: rec,
    source: "placeholder",
  };
}

function stripJsonFence(s) {
  let t = String(s || "").trim();
  if (t.startsWith("```")) {
    t = t.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/u, "");
  }
  return t.trim();
}

/**
 * Produces executive summary, risk narrative, and recommendations (OpenAI when configured, else rule-based placeholder).
 * @param {object} memoryOptions Optional `{ profile, domainRows }` for assessment memory prompt enrichment.
 */
async function generateNarrative(assessmentData, memoryOptions = null) {
  const apiKey = (import.meta.env.VITE_OPENAI_API_KEY || "").trim();
  const memoryBlock =
    memoryOptions?.profile && memoryOptions?.domainRows
      ? formatMemoryLibraryForPrompt(memoryOptions.profile, memoryOptions.domainRows)
      : "";

  if (!apiKey) {
    const base = generatePlaceholderNarrative(assessmentData);
    if (memoryBlock) {
      const note = `\n\n[Institutional memory — similar prior assessments]\n${memoryBlock.slice(0, 1200)}`;
      return {
        ...base,
        executiveSummary: base.executiveSummary + note,
        source: "placeholder",
      };
    }
    return base;
  }

  const userPrompt = `You are a senior Digital Risk consultant.

Based on the following TPRM maturity assessment data, generate:

1. Executive Summary (5–6 sentences, clear and concise, suitable for senior stakeholders)
2. Risk Narrative (explain the key risk themes, control weaknesses, and implications)
3. Recommendations (specific, actionable, prioritized improvements)

Assessment Data:
- Overall Maturity Score: ${assessmentData.overallScore}
- Maturity Level: ${assessmentData.maturityLevel} (${assessmentData.maturityTier})
- Weak Domains: ${assessmentData.weakDomainsList}
- Evidence Strength: ${assessmentData.evidenceSummary}
- Domain scores:
${assessmentData.domainScoresLine}
- Top gaps: ${assessmentData.topGapsSummary}
- Risk themes (engine): ${assessmentData.themesLine || "N/A"}

${getLibraryHintsForOpenAI(assessmentData.topGapsForLibrary || [])}

${memoryBlock ? `${memoryBlock}\n\n` : ""}

Guidelines:
- Use professional consulting tone
- Be direct and structured
- Avoid generic language
- Tie insights to business impact
- Make recommendations practical and prioritized`;

  try {
    const res = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: "gpt-4o-mini",
        temperature: 0.35,
        response_format: { type: "json_object" },
        messages: [
          {
            role: "system",
            content:
              'Respond only with a valid JSON object with exactly these string keys: "executiveSummary", "riskNarrative", "recommendationsText". recommendationsText may use newline-separated items for distinct actions.',
          },
          { role: "user", content: userPrompt },
        ],
      }),
    });

    if (!res.ok) {
      const errText = await res.text();
      throw new Error(`OpenAI API ${res.status}: ${errText.slice(0, 200)}`);
    }

    const payload = await res.json();
    const raw = payload?.choices?.[0]?.message?.content;
    const parsed = JSON.parse(stripJsonFence(raw));
    const executiveSummary = String(parsed.executiveSummary || "").trim();
    const riskNarrative = String(parsed.riskNarrative || "").trim();
    const recommendationsText = String(parsed.recommendationsText || "").trim();
    if (!executiveSummary || !riskNarrative || !recommendationsText) {
      throw new Error("Incomplete narrative JSON");
    }
    return { executiveSummary, riskNarrative, recommendationsText, source: "openai" };
  } catch (e) {
    console.warn("generateNarrative falling back to placeholder:", e);
    const fallback = generatePlaceholderNarrative(assessmentData);
    if (memoryBlock) {
      const note = `\n\n[Institutional memory — similar prior assessments]\n${memoryBlock.slice(0, 1200)}`;
      return { ...fallback, executiveSummary: fallback.executiveSummary + note, source: "placeholder" };
    }
    return { ...fallback, source: "placeholder" };
  }
}

function splitRecommendationLines(text) {
  return String(text || "")
    .split(/\n+/)
    .map((l) => l.replace(/^\s*[\d]+[.)]\s*/, "").trim())
    .filter(Boolean);
}

function clampScore(score) {
  const n = Number(score);
  if (!Number.isFinite(n)) return 3;
  if (n < 1) return 1;
  if (n > 5) return 5;
  return Math.round(n);
}

function normalizeEvidenceAiSuggestion(raw) {
  const src = raw || {};
  const suggestedScore = clampScore(src.suggestedScore);
  const confRaw = String(src.confidence || "Medium").trim().toLowerCase();
  const confidence =
    confRaw === "low" ? "Low" : confRaw === "high" ? "High" : "Medium";
  const asLines = (v, fallback) => {
    if (!v) return fallback;
    if (Array.isArray(v))
      return v.map((x) => String(x || "").trim()).filter(Boolean);
    return String(v)
      .split(/\n+/)
      .map((x) => x.trim())
      .filter(Boolean);
  };
  return {
    suggestedScore,
    confidence,
    evidenceSummary: String(src.evidenceSummary || "No evidence summary provided.").trim(),
    scoringRationale: String(src.scoringRationale || "No rationale returned.").trim(),
    evidenceUsed: asLines(src.evidenceUsed, ["No direct evidence provided."]),
    confidenceReason: String(
      src.confidenceReason || "Confidence reason not provided; defaulted from available evidence depth."
    ).trim(),
    evidenceGaps: asLines(src.evidenceGaps, ["No direct evidence provided."]),
    keyGaps: asLines(src.keyGaps, ["No key gaps provided."]),
    recommendedActions: asLines(src.recommendedActions, ["No recommended actions provided."]),
    missingEvidence: asLines(src.missingEvidence, ["No missing evidence noted."]),
  };
}

function buildClientProfileForAi(profile, assessmentType) {
  return [
    `Client: ${reportClientLine(profile)}`,
    `Industry: ${profile.industry}`,
    `Company size: ${profile.companySize}`,
    `Regulatory intensity: ${profile.regulatoryIntensity}`,
    `Third-party volume: ${profile.thirdPartyVolume}`,
    `Geographic footprint: ${profile.geographicFootprint}`,
    `Assessment type: ${assessmentType}`,
  ].join("\n");
}

function buildEvidencePrompt({
  clientProfile,
  domainName,
  evidenceText,
  evidenceStrength,
}) {
  return `You are a senior TPRM maturity assessment specialist.

Assess the submitted evidence for the domain below.

Client profile:
${clientProfile}

Domain:
${domainName}

Scoring rubric:
1 = Initial / Ad hoc
2 = Developing
3 = Defined
4 = Managed
5 = Optimized

Evidence:
${evidenceText}

Evidence strength selected by analyst:
${evidenceStrength}

Return JSON only:
{
  "suggestedScore": number,
  "confidence": "Low | Medium | High",
  "evidenceSummary": "string",
  "scoringRationale": "string",
  "evidenceUsed": ["string"],
  "confidenceReason": "string",
  "evidenceGaps": ["string"],
  "keyGaps": ["string"],
  "recommendedActions": ["string"],
  "missingEvidence": ["string"]
}

Rules:
- Do not invent evidence.
- If evidence is weak or missing, lower confidence.
- Tie score directly to evidence.
- Be conservative where evidence is incomplete.
- Explain the rationale clearly.
- Output JSON only.`;
}

function mockEvidenceSuggestion(domain, row, clientProfile, assessmentType, supportingText) {
  const notes = String(row?.notes || "").trim();
  const score = clampScore(row?.score ?? 3);
  const evidence = row?.evidence || "Moderate";
  const supportLen = String(supportingText || "").trim().length;
  let suggested = score;
  if (notes.length < 35) suggested -= 1;
  if (notes.length > 220) suggested += 1;
  if (supportLen > 600) suggested += 1;
  if (evidence === "Weak") suggested -= 1;
  if (evidence === "Strong") suggested += 1;
  suggested = clampScore(suggested);
  const confidence =
    evidence === "Strong" && notes.length > 120
      ? "High"
      : evidence === "Weak" || notes.length < 35
        ? "Low"
        : "Medium";

  const keyGaps = [];
  if (!notes) keyGaps.push("No domain-specific evidence notes were provided.");
  if (evidence === "Weak") keyGaps.push("Evidence strength is marked Weak.");
  if (notes.length < 80) keyGaps.push("Evidence narrative is brief; control operation depth is unclear.");
  if (keyGaps.length === 0) keyGaps.push("No critical gap inferred from mock logic; validate with artifact sampling.");

  const missingEvidence = [];
  if (!notes) missingEvidence.push("Detailed control evidence notes for this domain.");
  if (supportLen < 120) missingEvidence.push("Additional supporting text or artifacts mapped to this domain.");
  missingEvidence.push("Recent QA sample results and exception closure proof.");
  const evidenceUsed = [];
  if (notes) evidenceUsed.push(`Assessor notes excerpt: ${notes.slice(0, 220)}`);
  if (String(supportingText || "").trim()) {
    evidenceUsed.push(`Supporting text excerpt: ${String(supportingText).trim().slice(0, 220)}`);
  }
  if (evidenceUsed.length === 0) evidenceUsed.push("No direct evidence provided.");

  const evidenceGaps = [];
  if (!notes) evidenceGaps.push("No direct evidence provided.");
  if (!String(supportingText || "").trim()) {
    evidenceGaps.push("No pasted supporting text was provided for this domain.");
  }
  if (evidence === "Weak") {
    evidenceGaps.push("Evidence strength is Weak; objective control artifacts are limited.");
  }
  if (evidenceGaps.length === 0) {
    evidenceGaps.push("No direct evidence provided.");
  }
  const confidenceReason =
    confidence === "High"
      ? "High confidence because evidence strength is Strong and notes are sufficiently detailed."
      : confidence === "Low"
        ? "Low confidence because direct evidence is limited, weak, or too brief to fully support scoring."
        : "Medium confidence because evidence provides partial support but leaves notable verification gaps.";

  return normalizeEvidenceAiSuggestion({
    suggestedScore: suggested,
    confidence,
    evidenceSummary: notes
      ? `Using analyst notes and ${evidence} evidence setting for ${domain.label}.`
      : `Limited direct notes for ${domain.label}; recommendation based on sparse inputs and ${evidence} evidence setting.`,
    scoringRationale: `Mock scoring logic considered current score (${score}/5), notes length (${notes.length} chars), supporting text volume (${supportLen} chars), and evidence strength (${evidence}). Conservative adjustments were applied when evidence was thin.`,
    keyGaps,
    recommendedActions: [
      `Validate ${domain.label} controls against the 1–5 rubric before finalizing score ${suggested}/5.`,
      "Attach objective artifacts and add concise control effectiveness notes.",
      "Run reviewer challenge session before committee sign-off.",
    ],
    evidenceUsed,
    confidenceReason,
    evidenceGaps,
    missingEvidence,
    _meta: { source: "mock", assessmentType, clientProfile },
  });
}

/** Build per-domain AI suggestion payloads for demo load (reuses mock analyzer output shape). */
function buildDemoAiSuggestionsFromWorkspace(profile, assessmentType, rows, supportingText) {
  const out = {};
  const clientProfile = buildClientProfileForAi(profile, assessmentType);
  for (const d of DOMAINS) {
    const row = rows[d.id];
    const suggestion = mockEvidenceSuggestion(d, row, clientProfile, assessmentType, supportingText);
    out[d.id] = { status: "Suggested", suggestion, source: "mock" };
  }
  return out;
}

async function analyzeDomainEvidenceWithLlm({
  profile,
  assessmentType,
  domain,
  row,
  supportingText,
}) {
  const clientProfile = buildClientProfileForAi(profile, assessmentType);
  const notes = String(row?.notes || "").trim();
  const evidenceText = [
    notes ? `Domain notes: ${notes}` : "Domain notes: (none provided)",
    supportingText?.trim()
      ? `Supporting text: ${supportingText.trim()}`
      : "Supporting text: (none provided)",
  ].join("\n\n");
  const apiKey = (import.meta.env.VITE_OPENAI_API_KEY || "").trim();
  if (!apiKey) {
    return mockEvidenceSuggestion(domain, row, clientProfile, assessmentType, supportingText);
  }

  const prompt = buildEvidencePrompt({
    clientProfile,
    domainName: domain.label,
    evidenceText,
    evidenceStrength: row?.evidence || "Moderate",
  });

  const res = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: "gpt-4o-mini",
      temperature: 0.2,
      response_format: { type: "json_object" },
      messages: [
        {
          role: "system",
          content:
            'Return only valid JSON with keys: "suggestedScore", "confidence", "evidenceSummary", "scoringRationale", "evidenceUsed", "confidenceReason", "evidenceGaps", "keyGaps", "recommendedActions", "missingEvidence". Do not invent evidence; if none is directly provided, set evidenceUsed and evidenceGaps entries to "No direct evidence provided."',
        },
        { role: "user", content: prompt },
      ],
    }),
  });

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`OpenAI API ${res.status}: ${errText.slice(0, 200)}`);
  }
  const payload = await res.json();
  const raw = payload?.choices?.[0]?.message?.content;
  return normalizeEvidenceAiSuggestion(JSON.parse(stripJsonFence(raw)));
}

/** Matrix maturity: 1–2 Low, 3 Moderate, 4–5 Advanced */
function matrixMaturityBand(score) {
  if (score <= 2)
    return { label: "Low", badgeBg: "#fee2e2", badgeFg: "#991b1b", badgeBd: "#fecaca" };
  if (score === 3)
    return { label: "Moderate", badgeBg: "#fef3c7", badgeFg: "#b45309", badgeBd: "#fcd34d" };
  return { label: "Advanced", badgeBg: "#dcfce7", badgeFg: "#15803d", badgeBd: "#86efac" };
}

function matrixPriority(score, evidence) {
  if (score <= 2 && evidence === "Weak")
    return { label: "Critical", color: "#991b1b", fontWeight: 800 };
  if (score <= 2) return { label: "High", color: "#ea580c", fontWeight: 700 };
  if (score === 3) return { label: "Medium", color: "#d97706", fontWeight: 600 };
  return { label: "Low", color: "#16a34a", fontWeight: 600 };
}

function formatReportDate() {
  return new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
}

function reportClientLine(profile) {
  const name = (profile.clientName || "").trim();
  if (name) return name;
  return `${profile.industry || "Organization"} · ${profile.companySize || "Enterprise"}`;
}

/**
 * From raw inputs (notes, evidence, score) produce consulting-style finding, risk, and domain recommendation.
 */
function deriveDomainConsultingBlocks(entry, profile) {
  const mat = matrixMaturityBand(entry.score);
  const pri = matrixPriority(entry.score, entry.evidence);
  const lib = TPRM_DOMAIN_LIBRARY[entry.id];
  const notes = (entry.notes || "").trim();
  const typical = lib?.commonGaps?.[0] ?? "";

  let specificFinding;
  if (notes) {
    specificFinding = `Documented observations for ${entry.label}: "${notes}" Taken with the scored posture (${entry.score}/5, ${entry.evidence} evidence; ${mat.label} maturity band; ${pri.label} priority), the specific finding is that this lifecycle area ${
      entry.score <= 2
        ? "shows material gaps warranting prioritized remediation and sponsorship."
        : entry.score === 3
          ? "delivers inconsistent execution—metrics and assurance cadence should be tightened to match policy intent."
          : "performs at a relative strength suitable to anchor uplift in weaker adjacent domains."
    }`;
  } else {
    specificFinding = `Without free-text notes, the structured assessment positions ${entry.label} at ${entry.score}/5 with ${entry.evidence} evidence (${mat.label}; ${pri.label} priority).${typical ? ` Benchmark gap pattern often includes: ${typical}` : ""} Summary implication: ${buildGapRationale(entry)}`;
  }

  let riskImplication;
  if (entry.score <= 2 && entry.evidence === "Weak") {
    riskImplication = `Compounded exposure—low scored maturity plus weak evidence depth impairs defensibility of tiering, residual risk, and supervisory narratives for relationships depending on ${entry.label}.`;
  } else if (entry.score <= 2) {
    riskImplication = `Elevated third-party lifecycle risk: shortcomings in ${entry.label} propagate into diligence, contractual leverage, monitoring triggers, and exit readiness—particularly for tier-critical vendors.`;
  } else if (entry.score === 3) {
    riskImplication = `Moderate enterprise risk: uneven execution may produce attestations or dashboards that outrun contemporaneous artifacts—raising audit and regulatory challenge exposure under scrutiny.`;
  } else {
    riskImplication = `Residual concern is primarily drift—${entry.label} should be leveraged as a control anchor without diverting investment from weaker lifecycle lanes.`;
  }
  if (
    (profile.regulatoryIntensity === "High" || profile.regulatoryIntensity === "Very High") &&
    entry.id === "regulatory" &&
    entry.score <= 3
  ) {
    riskImplication +=
      " Supervisory intensity amplifies any documented misalignment between policy and operating proof.";
  }

  let recommendation;
  const tplIdx =
    entry.score <= 2 ? 0 : Math.min(1, Math.max(0, (lib?.templates?.length ?? 1) - 1));
  const tpl = lib?.templates?.[tplIdx];
  if (tpl) {
    recommendation = tpl.replace(/\{domain\}/g, entry.label);
    if (entry.evidence === "Weak") {
      recommendation +=
        " Prioritize contemporaneous evidence packs and sampling discipline suitable for regulatory or audit inquiry.";
    } else if (entry.evidence === "Moderate") {
      recommendation +=
        " Stand up QA sampling on artifacts underlying assertions within the next two governance cycles.";
    }
  } else {
    recommendation = `Assign executive sponsor, 90-day measurable outcomes, and committee-visible KPIs for ${entry.label}, scaled to tier materiality.`;
  }

  return {
    specificFinding,
    riskImplication,
    recommendation,
  };
}

/** Minimum note length to treat low scores as adequately explained without Strong evidence */
const MIN_LOW_SCORE_NOTE_CHARS = 40;

const DOMAIN_REC_WEAK_SUFFIX =
  " Prioritize contemporaneous evidence packs and sampling discipline suitable for regulatory or audit inquiry.";
const DOMAIN_REC_MOD_SUFFIX =
  " Stand up QA sampling on artifacts underlying assertions within the next two governance cycles.";

function stripDomainRecommendationSuffix(recommendation) {
  let s = String(recommendation || "");
  if (s.endsWith(DOMAIN_REC_WEAK_SUFFIX)) s = s.slice(0, -DOMAIN_REC_WEAK_SUFFIX.length);
  else if (s.endsWith(DOMAIN_REC_MOD_SUFFIX)) s = s.slice(0, -DOMAIN_REC_MOD_SUFFIX.length);
  return s.trim();
}

/**
 * Per-domain pre-export / QA messages (same strings as shown on Results).
 */
function collectDomainPreReportQaMessages(e, profile, aiSuggestionsByDomain = {}) {
  const warnings = [];
  const notes = (e.notes || "").trim();

  if (!notes) {
    warnings.push(`${e.label}: missing assessor notes.`);
  }

  if (e.score <= 2 && !notes) {
    warnings.push(
      `[Pre-report] ${e.label}: score ≤2 with no assessor notes — add evidence before export.`
    );
  }

  if (e.score <= 2) {
    const strongExplanation =
      notes.length >= MIN_LOW_SCORE_NOTE_CHARS || e.evidence === "Strong";
    if (!strongExplanation) {
      warnings.push(
        `${e.label}: score ≤2 without strong explanation — add at least ${MIN_LOW_SCORE_NOTE_CHARS} characters of notes or set evidence to Strong if artifacts substantiate the rating.`
      );
    }
  }

  if (e.evidence === "Weak" && e.score >= 4) {
    warnings.push(
      `[Pre-report] ${e.label}: Weak evidence with high maturity score (${e.score}/5) — reconcile maturity score with artifact depth before export.`
    );
  }

  const aiReview = aiSuggestionsByDomain[e.id];
  const aiSug = aiReview?.suggestion;
  if (aiSug?.confidence === "Low") {
    warnings.push(
      `[Pre-report] ${e.label}: AI confidence is Low — review evidence and scoring before export.`
    );
  }
  if (aiSug && typeof aiSug.suggestedScore === "number") {
    const suggested = clampScore(aiSug.suggestedScore);
    if (Math.abs(e.score - suggested) >= 2) {
      warnings.push(
        `[Pre-report] ${e.label}: manual score (${e.score}/5) differs from AI suggested (${suggested}/5) by ≥2 — reconcile before export.`
      );
    }
  }

  const blocks = deriveDomainConsultingBlocks(e, profile);
  const tpls = TPRM_DOMAIN_LIBRARY[e.id]?.templates;
  if (tpls?.length) {
    const expectedIdx = e.score <= 2 ? 0 : Math.min(1, tpls.length - 1);
    const expectedBase = tpls[expectedIdx].replace(/\{domain\}/g, e.label);
    const core = stripDomainRecommendationSuffix(blocks.recommendation);
    if (expectedBase && core !== expectedBase) {
      warnings.push(
        `${e.label}: recommendation language may not match this score tier — verify score, evidence, and inputs for consistency.`
      );
    }
  }

  return warnings;
}

/**
 * Pre-export checks: missing notes, underspecified low scores, score/evidence mismatch,
 * alignment between score tier and generated domain recommendation text, AI confidence,
 * and manual vs. AI score divergence.
 */
function computeReportQualityWarnings(profile, liveAssessment, aiSuggestionsByDomain = {}) {
  return liveAssessment.entries.flatMap((e) =>
    collectDomainPreReportQaMessages(e, profile, aiSuggestionsByDomain)
  );
}

/** Shared signals for Review Queue and Command Center (needs review, AI deltas, QA counts). */
function computeDomainReviewSignals(e, profile, aiSuggestionsByDomain, reviewFollowUpByDomain = {}) {
  const aiReview = aiSuggestionsByDomain[e.id];
  const aiSug = aiReview?.suggestion;
  const suggested = aiSug ? clampScore(aiSug.suggestedScore) : null;
  const diff = suggested !== null ? e.score - suggested : null;
  const qaMessages = collectDomainPreReportQaMessages(e, profile, aiSuggestionsByDomain);
  const qaCount = qaMessages.length;
  const pri = matrixPriority(e.score, e.evidence);
  const bigDiff = suggested !== null && Math.abs(diff) >= 2;
  const lowConf = aiSug?.confidence === "Low";
  const weakHigh = e.evidence === "Weak" && e.score >= 4;
  const priHot = pri.label === "Critical" || pri.label === "High";
  const follow = !!reviewFollowUpByDomain[e.id];
  const aiStatus = aiReview?.status || "Not analyzed";
  const needsReview =
    follow ||
    qaCount > 0 ||
    bigDiff ||
    lowConf ||
    weakHigh ||
    priHot ||
    (aiStatus === "Suggested" && !!aiSug);
  const hasCriticalPreReportQa = qaMessages.some((m) => m.includes("[Pre-report]"));
  const criticalQaUnresolved =
    hasCriticalPreReportQa || (pri.label === "Critical" && qaCount > 0);
  return {
    aiReview,
    aiSug,
    suggested,
    diff,
    qaMessages,
    qaCount,
    pri,
    bigDiff,
    lowConf,
    weakHigh,
    priHot,
    follow,
    aiStatus,
    needsReview,
    criticalQaUnresolved,
  };
}

function computeEngagementCommandCenterModel({
  profile,
  liveAssessment,
  aiSuggestionsByDomain,
  reviewFollowUpByDomain,
  supportingEvidenceText,
  scopeNotes,
  managerReviewStatus,
  completionPct,
}) {
  const entries = liveAssessment.entries;
  const signals = entries.map((e) => ({
    e,
    ...computeDomainReviewSignals(e, profile, aiSuggestionsByDomain, reviewFollowUpByDomain),
  }));

  const domainsNeedingReview = signals.filter((s) => s.needsReview).length;
  const criticalGaps = entries.filter((e) => {
    const p = matrixPriority(e.score, e.evidence);
    return p.label === "Critical" || (e.score <= 2 && e.evidence === "Weak");
  }).length;

  const unresolvedCriticalQa = signals.some((s) => s.criticalQaUnresolved);
  const allDomainsScored = completionPct === 100;
  const managerReviewOk =
    managerReviewStatus === "complete" || managerReviewStatus === "not_required";

  const reportReady = allDomainsScored && !unresolvedCriticalQa && managerReviewOk;
  const reportReadinessReasons = [];
  if (!allDomainsScored) reportReadinessReasons.push("Complete score and evidence strength for every domain.");
  if (unresolvedCriticalQa)
    reportReadinessReasons.push(
      "Resolve critical QA items (Critical priority with open QA, or any [Pre-report] warning on a domain)."
    );
  if (!managerReviewOk)
    reportReadinessReasons.push('Mark "Manager review" complete on the Command Center, or set it to not required.');

  const profileFieldsOk = ["industry", "companySize", "regulatoryIntensity", "thirdPartyVolume", "geographicFootprint"].every(
    (k) => Boolean(profile[k])
  );
  const profileEnriched =
    profileFieldsOk &&
    (Boolean((profile.clientName || "").trim()) || (scopeNotes || "").trim().length >= 12);

  const evidenceSubmitted =
    (supportingEvidenceText || "").trim().length >= 40 ||
    entries.every((e) => (e.notes || "").trim().length >= 15);

  const aiAnalysisComplete = DOMAINS.every((d) => Boolean(aiSuggestionsByDomain[d.id]?.suggestion));

  const managerReviewStageComplete =
    managerReviewStatus === "complete" ||
    managerReviewStatus === "not_required" ||
    (domainsNeedingReview === 0 && aiAnalysisComplete);

  const qaChecksComplete = !unresolvedCriticalQa;

  const stages = [
    { key: "profile", label: "Client profile complete", done: profileEnriched },
    { key: "evidence", label: "Evidence submitted", done: evidenceSubmitted },
    { key: "ai", label: "AI analysis complete", done: aiAnalysisComplete },
    { key: "mgr", label: "Manager review complete", done: managerReviewStageComplete },
    { key: "qa", label: "QA checks complete", done: qaChecksComplete },
    { key: "report", label: "Report ready", done: reportReady },
  ];

  const topRiskDomains = [...entries]
    .sort((a, b) => {
      const pa = matrixPriority(a.score, a.evidence);
      const pb = matrixPriority(b.score, b.evidence);
      const rank = { Critical: 0, High: 1, Medium: 2, Low: 3 };
      const d = rank[pa.label] - rank[pb.label];
      if (d !== 0) return d;
      return a.score - b.score;
    })
    .slice(0, 5)
    .map((e) => ({
      id: e.id,
      label: e.label,
      score: e.score,
      evidence: e.evidence,
      priority: matrixPriority(e.score, e.evidence).label,
    }));

  let evidenceGapDomains = 0;
  let evidenceGapBullets = 0;
  entries.forEach((e) => {
    const s = aiSuggestionsByDomain[e.id]?.suggestion;
    if (!s) return;
    const merged = [...(s.evidenceGaps || []), ...(s.missingEvidence || [])].filter((x) => {
      const t = String(x || "").trim();
      return t && t !== "No direct evidence provided.";
    });
    if (merged.length) {
      evidenceGapDomains += 1;
      evidenceGapBullets += merged.length;
    }
  });

  const belowBenchmark = entries.filter((e) => e.score < BENCHMARK_SCORE).map((e) => e.label);

  return {
    signals,
    domainsNeedingReview,
    criticalGaps,
    reportReady,
    reportReadinessReasons,
    stages,
    topRiskDomains,
    evidenceGapDomains,
    evidenceGapBullets,
    belowBenchmark,
    unresolvedCriticalQa,
    allDomainsScored,
    managerReviewOk,
  };
}

/**
 * Full structured client-ready first draft after Generate assessment (analyst edit vs write from scratch).
 */
function buildClientAssessmentDraft(profile, liveAssessment, aiNarratives) {
  const exec =
    aiNarratives?.executiveSummary?.trim() || liveAssessment.executiveSummary;
  const domainFindings = liveAssessment.entries.map((e) => {
    const mat = matrixMaturityBand(e.score);
    const pri = matrixPriority(e.score, e.evidence);
    const blocks = deriveDomainConsultingBlocks(e, profile);
    return {
      id: e.id,
      title: e.label,
      score: e.score,
      evidence: e.evidence,
      maturityLabel: mat.label,
      priority: pri.label,
      lifecycleScope: e.description || "",
      ...blocks,
    };
  });

  const benchGap = liveAssessment.overallScore - BENCHMARK_SCORE;
  const keyRisks = liveAssessment.themes.map((text, id) => ({ id: `theme-${id}`, text }));
  keyRisks.push({
    id: "benchmark",
    text: `Portfolio mean score ${liveAssessment.overallScore}/5 versus sector benchmark ${BENCHMARK_SCORE} (gap ${benchGap >= 0 ? "+" : ""}${benchGap.toFixed(1)}).`,
  });
  const weakCt = liveAssessment.entries.filter((x) => x.evidence === "Weak").length;
  if (weakCt > 0) {
    keyRisks.push({
      id: "evidence",
      text: `${weakCt} domain(s) rely on Weak evidence—prioritize artifact programs before relying on scores for regulatory or audit narratives.`,
    });
  }

  const recommendations = aiNarratives?.recommendationsText?.trim()
    ? splitRecommendationLines(aiNarratives.recommendationsText)
    : [...liveAssessment.remediation];

  const roadmap = liveAssessment.roadmap;

  return {
    meta: {
      clientName: reportClientLine(profile),
      dateStr: formatReportDate(),
      assessmentType: profile.assessmentType || "TPRM maturity assessment",
      overallScore: liveAssessment.overallScore,
      maturityBand: liveAssessment.maturityBand,
      maturityTier: liveAssessment.maturityTier,
      benchmark: BENCHMARK_SCORE,
      benchGap,
    },
    executiveSummary: exec,
    domainFindings,
    keyRisks,
    recommendations,
    roadmap,
    narrativeSource: aiNarratives?.source || null,
  };
}

function formatClientDraftPlainText(draft) {
  if (!draft) return "";
  const m = draft.meta;
  let s = "";
  s += "TPRM MATURITY ASSESSMENT — CLIENT DRAFT (FIRST PASS)\r\n\r\n";
  s += `Client / organization: ${m.clientName}\r\n`;
  s += `Date: ${m.dateStr}\r\n`;
  s += `Assessment type: ${m.assessmentType}\r\n`;
  s += `Overall maturity: ${m.overallScore}/5 · ${m.maturityBand} (${m.maturityTier} tier)\r\n\r\n`;

  s += "EXECUTIVE SUMMARY\r\n";
  s += `${draft.executiveSummary}\r\n\r\n`;

  s += "DOMAIN-BY-DOMAIN FINDINGS\r\n";
  draft.domainFindings.forEach((d, i) => {
    s += `\r\n${i + 1}. ${d.title}\r\n`;
    s += `   Inputs: score ${d.score}/5 · Evidence strength: ${d.evidence} · Band: ${d.maturityLabel} · Priority: ${d.priority}\r\n`;
    if (d.lifecycleScope) s += `   Lifecycle scope (reference): ${d.lifecycleScope}\r\n`;
    s += `   Specific finding:\r\n   ${d.specificFinding}\r\n`;
    s += `   Risk implication:\r\n   ${d.riskImplication}\r\n`;
    s += `   Recommendation:\r\n   ${d.recommendation}\r\n`;
  });

  s += "\r\nKEY RISKS\r\n";
  draft.keyRisks.forEach((r) => {
    s += `• ${r.text}\r\n`;
  });

  s += "\r\nRECOMMENDATIONS\r\n";
  draft.recommendations.forEach((line, i) => {
    s += `${i + 1}. ${line}\r\n`;
  });

  s += "\r\nROADMAP — 30 / 60 / 90 DAYS\r\n";
  s += "\r\n30 days — stabilize critical gaps\r\n";
  draft.roadmap.d30.forEach((x) => {
    s += `  • ${x}\r\n`;
  });
  s += "\r\n60 days — standardize controls and evidence\r\n";
  draft.roadmap.d60.forEach((x) => {
    s += `  • ${x}\r\n`;
  });
  s += "\r\n90 days — optimize reporting and monitoring\r\n";
  draft.roadmap.d90.forEach((x) => {
    s += `  • ${x}\r\n`;
  });

  s += "\r\n— End of draft — analysts should tailor jurisdiction, materiality, and names before client delivery.\r\n";
  return s;
}

function ClientAssessmentDraftSection({ draft, narrativeLoading }) {
  if (!draft) return null;
  const m = draft.meta;
  return (
    <SectionCard
      title="Client-ready assessment draft"
      subtitle="Each domain converts score, evidence, and assessor notes into finding, risk implication, and recommendation — edit to finalize."
    >
      <div style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 16, alignItems: "center" }}>
        <button
          type="button"
          onClick={() => {
            navigator.clipboard?.writeText(formatClientDraftPlainText(draft)).catch(() => {});
          }}
          style={{
            padding: "8px 14px",
            borderRadius: 8,
            border: `1px solid ${C.border}`,
            background: C.surface,
            fontWeight: 700,
            fontSize: 12,
            cursor: "pointer",
          }}
        >
          Copy full draft (plain text)
        </button>
        <button
          type="button"
          onClick={() => {
            const blob = new Blob([formatClientDraftPlainText(draft)], {
              type: "text/plain;charset=utf-8",
            });
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = "TPRM_Client_Assessment_Draft.txt";
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(url), 1000);
          }}
          style={{
            padding: "8px 14px",
            borderRadius: 8,
            border: `1px solid ${C.border}`,
            background: "#f8fafc",
            fontWeight: 700,
            fontSize: 12,
            cursor: "pointer",
          }}
        >
          Download draft (.txt)
        </button>
        {narrativeLoading && (
          <span style={{ fontSize: 12, color: C.muted, fontWeight: 600 }}>
            AI narrative sections updating…
          </span>
        )}
        {draft.narrativeSource === "openai" && (
          <span style={{ fontSize: 11, fontWeight: 800, color: "#15803d" }}>Includes AI-assisted narrative</span>
        )}
      </div>

      <div
        style={{
          borderRadius: 12,
          border: `1px solid ${C.border}`,
          background: "#fafafa",
          padding: "18px 20px",
          fontSize: 13,
          lineHeight: 1.65,
          color: C.text,
        }}
      >
        <p style={{ margin: "0 0 8px", fontSize: 11, fontWeight: 800, letterSpacing: "0.1em", color: C.muted }}>
          COVER
        </p>
        <p style={{ margin: 0, fontWeight: 700 }}>{m.clientName}</p>
        <p style={{ margin: "6px 0 0", color: C.muted }}>
          {m.dateStr} · {m.assessmentType}
        </p>
        <p style={{ margin: "10px 0 0" }}>
          Overall maturity <strong>{m.overallScore}/5</strong> · {m.maturityBand} ({m.maturityTier} tier) · Benchmark gap{" "}
          <strong style={{ color: m.benchGap >= 0 ? "#15803d" : "#b91c1c" }}>
            {m.benchGap >= 0 ? "+" : ""}
            {m.benchGap.toFixed(1)}
          </strong>{" "}
          vs {m.benchmark}
        </p>

        <h3 style={{ margin: "22px 0 10px", fontSize: 14, fontWeight: 800, color: C.accent }}>1. Executive summary</h3>
        <p style={{ margin: 0, whiteSpace: "pre-wrap" }}>{draft.executiveSummary}</p>

        <h3 style={{ margin: "22px 0 10px", fontSize: 14, fontWeight: 800, color: C.accent }}>
          2. Domain-by-domain findings
        </h3>
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          {draft.domainFindings.map((d, i) => (
            <div
              key={d.id}
              style={{
                padding: 14,
                borderRadius: 10,
                background: "#fff",
                border: `1px solid ${C.border}`,
              }}
            >
              <div style={{ fontWeight: 800, marginBottom: 6 }}>
                {i + 1}. {d.title}
              </div>
              <div style={{ fontSize: 12, color: C.muted, marginBottom: 8 }}>
                Inputs · Score {d.score}/5 · Evidence strength {d.evidence} · {d.maturityLabel} · {d.priority}
              </div>
              {d.lifecycleScope ? (
                <p style={{ margin: "0 0 10px", fontSize: 11, color: "#64748b", fontStyle: "italic" }}>
                  Lifecycle reference: {d.lifecycleScope}
                </p>
              ) : null}
              <p style={{ margin: "0 0 10px", fontSize: 13 }}>
                <strong style={{ color: C.accent }}>Specific finding.</strong> {d.specificFinding}
              </p>
              <p style={{ margin: "0 0 10px", fontSize: 13 }}>
                <strong style={{ color: "#b45309" }}>Risk implication.</strong> {d.riskImplication}
              </p>
              <p style={{ margin: 0, fontSize: 13 }}>
                <strong style={{ color: "#15803d" }}>Recommendation.</strong> {d.recommendation}
              </p>
            </div>
          ))}
        </div>

        <h3 style={{ margin: "22px 0 10px", fontSize: 14, fontWeight: 800, color: C.accent }}>3. Key risks</h3>
        <ul style={{ margin: 0, paddingLeft: 20 }}>
          {draft.keyRisks.map((r) => (
            <li key={r.id} style={{ marginBottom: 8 }}>
              {r.text}
            </li>
          ))}
        </ul>

        <h3 style={{ margin: "22px 0 10px", fontSize: 14, fontWeight: 800, color: C.accent }}>4. Recommendations</h3>
        <ol style={{ margin: 0, paddingLeft: 22 }}>
          {draft.recommendations.map((line, i) => (
            <li key={i} style={{ marginBottom: 8 }}>
              {line}
            </li>
          ))}
        </ol>

        <h3 style={{ margin: "22px 0 10px", fontSize: 14, fontWeight: 800, color: C.accent }}>
          5. Roadmap — 30 / 60 / 90 days
        </h3>
        <div style={{ display: "grid", gap: 16, gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))" }}>
          {[
            { k: "30 days — stabilize critical gaps", items: draft.roadmap.d30 },
            { k: "60 days — standardize controls and evidence", items: draft.roadmap.d60 },
            { k: "90 days — optimize reporting and monitoring", items: draft.roadmap.d90 },
          ].map((col) => (
            <div key={col.k}>
              <div style={{ fontSize: 11, fontWeight: 800, color: C.muted, marginBottom: 8 }}>{col.k}</div>
              <ul style={{ margin: 0, paddingLeft: 18, fontSize: 12 }}>
                {col.items.map((x, i) => (
                  <li key={i} style={{ marginBottom: 6 }}>
                    {x}
                  </li>
                ))}
              </ul>
            </div>
          ))}
        </div>
      </div>
    </SectionCard>
  );
}

function heatmapCellStyle(score) {
  const m = matrixMaturityBand(score);
  return { background: m.badgeBg, border: `1px solid ${m.badgeBd}`, color: m.badgeFg };
}

function evidenceConfidenceNumeric(ev) {
  if (ev === "Weak") return 1;
  if (ev === "Moderate") return 2;
  return 3;
}

function computeExecutiveKPIs(entries, overallScore) {
  const belowBench = entries.filter((e) => e.score < BENCHMARK_SCORE).length;
  const criticalGaps = entries.filter((e) => e.score <= 2 && e.evidence === "Weak").length;
  const avgEv =
    entries.reduce((s, e) => s + evidenceConfidenceNumeric(e.evidence), 0) /
    Math.max(entries.length, 1);
  const confidencePct = Math.round((avgEv / 3) * 100);
  return { belowBench, criticalGaps, confidencePct, overallScore };
}

/** Y-axis for risk map: Weak=1 (bottom), Moderate=3, Strong=5 (top) */
function evidenceAxisScore(ev) {
  if (ev === "Weak") return 1;
  if (ev === "Moderate") return 3;
  return 5;
}

/** Deterministic jitter 0–0.3 so layout is stable between renders */
function stableJitter(id, axis) {
  let h = 0;
  const s = `${id}:${axis}`;
  for (let i = 0; i < s.length; i++) {
    h = (h << 5) - h + s.charCodeAt(i);
    h |= 0;
  }
  return (Math.abs(h) % 301) / 1000;
}

/** Bubble diameter: low maturity scores = larger (more prominent risk) */
function bubbleDiameterByScore(score) {
  if (score <= 2) return 46;
  if (score === 3) return 30;
  return 20;
}

const RISK_MAP_ABBREV = {
  governance: "Gov",
  inventory: "Inv",
  segmentation: "Seg",
  dueDiligence: "Due",
  contracting: "Ctr",
  monitoring: "Mon",
  issues: "Iss",
  reporting: "Rep",
  technology: "Tech",
  regulatory: "Reg",
};

function riskMapShortLabel(id) {
  return RISK_MAP_ABBREV[id] ?? id.slice(0, 3);
}

function RiskPriorityMapSvg({ entries }) {
  const W = 560;
  const H = 400;
  const padL = 56;
  const padR = 24;
  const padT = 36;
  const padB = 56;
  const plotW = W - padL - padR;
  const plotH = H - padT - padB;
  const xMin = 0.75;
  const xMax = 5.45;
  const yMin = 0.75;
  const yMax = 5.45;
  const n = entries.length;

  function toX(xm) {
    return padL + ((xm - xMin) / (xMax - xMin)) * plotW;
  }
  function toY(ym) {
    return padT + plotH - ((ym - yMin) / (yMax - yMin)) * plotH;
  }

  const midXm = 3;
  const midYm = 3;
  const midPx = toX(midXm);
  const midPy = toY(midYm);

  const points = entries.map((e, idx) => {
    const jx = stableJitter(e.id, "x");
    const jy = stableJitter(e.id, "y");
    const ring = 0.14;
    const angle = (2 * Math.PI * idx) / Math.max(n, 1) + stableJitter(e.id, "ang") * 0.4;
    let xm = e.score + jx + Math.cos(angle) * ring;
    let ym = evidenceAxisScore(e.evidence) + jy + Math.sin(angle) * ring;
    xm = Math.min(Math.max(xm, xMin + 0.06), xMax - 0.06);
    ym = Math.min(Math.max(ym, yMin + 0.06), yMax - 0.06);
    const cx = toX(xm);
    const cy = toY(ym);
    const r = bubbleDiameterByScore(e.score) / 2;
    const pri = matrixPriority(e.score, e.evidence);
    const crit = pri.label === "Critical";
    const fill =
      e.score <= 2 ? "rgba(254,202,202,0.92)" : e.score === 3 ? "rgba(253,224,71,0.88)" : "rgba(187,247,208,0.92)";
    return { e, cx, cy, r, pri, crit, fill };
  });

  return (
    <div style={{ width: "100%", overflow: "auto" }}>
      <svg
        width="100%"
        viewBox={`0 0 ${W} ${H}`}
        style={{ display: "block", maxHeight: 440, borderRadius: 16, border: `1px solid ${C.border}`, background: "#fafafa" }}
      >
        <defs>
          <clipPath id="riskPlotClip">
            <rect x={padL} y={padT} width={plotW} height={plotH} rx="8" />
          </clipPath>
        </defs>

        <g clipPath="url(#riskPlotClip)">
          <rect x={padL} y={midPy} width={midPx - padL} height={padT + plotH - midPy} fill="rgba(254,202,202,0.45)" />
          <rect x={midPx} y={padT} width={padL + plotW - midPx} height={midPy - padT} fill="rgba(187,247,208,0.4)" />
          <rect x={padL} y={padT} width={midPx - padL} height={midPy - padT} fill="rgba(254,243,199,0.3)" />
          <rect x={midPx} y={midPy} width={padL + plotW - midPx} height={padT + plotH - midPy} fill="rgba(241,245,249,0.55)" />
        </g>

        {[1, 2, 3, 4, 5].map((tick) => (
          <g key={`gx-${tick}`}>
            <line x1={toX(tick)} y1={padT} x2={toX(tick)} y2={padT + plotH} stroke="#e2e8f0" strokeWidth="1" />
            <line x1={padL} y1={toY(tick)} x2={padL + plotW} y2={toY(tick)} stroke="#e2e8f0" strokeWidth="1" />
          </g>
        ))}

        <line x1={padL} y1={padT + plotH} x2={padL + plotW} y2={padT + plotH} stroke="#64748b" strokeWidth="1.5" />
        <line x1={padL} y1={padT} x2={padL} y2={padT + plotH} stroke="#64748b" strokeWidth="1.5" />

        {[1, 2, 3, 4, 5].map((t) => (
          <text key={`xt-${t}`} x={toX(t)} y={H - 32} textAnchor="middle" fontSize="10" fill={C.muted} fontWeight="700">
            {t}
          </text>
        ))}
        <text x={padL + plotW / 2} y={H - 12} textAnchor="middle" fontSize="11" fill={C.muted} fontWeight="700">
          Domain score (1–5)
        </text>

        <text x={18} y={padT + plotH / 2} textAnchor="middle" fontSize="10" fill={C.muted} fontWeight="700" transform={`rotate(-90,18,${padT + plotH / 2})`}>
          Evidence strength
        </text>
        <text x={padL - 10} y={toY(1) + 4} textAnchor="end" fontSize="9" fill="#b91c1c" fontWeight="700">
          Weak (1)
        </text>
        <text x={padL - 10} y={toY(3) + 4} textAnchor="end" fontSize="9" fill={C.muted} fontWeight="600">
          Mod (3)
        </text>
        <text x={padL - 10} y={toY(5) + 4} textAnchor="end" fontSize="9" fill="#15803d" fontWeight="700">
          Strong (5)
        </text>

        <text x={padL + 6} y={padT + plotH - 6} fontSize="9" fill="#991b1b" fontWeight="700">
          Higher risk
        </text>
        <text x={padL + plotW - 6} y={padT + 14} textAnchor="end" fontSize="9" fill="#15803d" fontWeight="700">
          Lower risk
        </text>

        {points.map(({ e, cx, cy, r, pri, crit, fill }) => (
          <g key={e.id}>
            <circle
              cx={cx}
              cy={cy}
              r={r}
              fill={fill}
              stroke={crit ? "#991b1b" : pri.color}
              strokeWidth={crit ? 3 : 2}
              style={crit ? { filter: "drop-shadow(0 0 6px rgba(153,27,27,0.45))" } : undefined}
            >
              <title>
                {`${e.label}\nScore: ${e.score}/5 · Evidence: ${e.evidence} (axis ${evidenceAxisScore(e.evidence)})\nPriority: ${pri.label}`}
              </title>
            </circle>
            <text x={cx} y={cy + r + 13} textAnchor="middle" fontSize="10" fontWeight="800" fill="#334155">
              {riskMapShortLabel(e.id)}
            </text>
          </g>
        ))}
      </svg>
    </div>
  );
}

function buildExecutiveInsights(entries) {
  const out = [];
  const belowBench = entries.filter((e) => e.score < BENCHMARK_SCORE).length;
  const weakEv = entries.filter((e) => e.evidence === "Weak").length;

  if (belowBench >= 5) {
    out.push(
      "A majority of domains sit below the sector benchmark—this pattern suggests systemic maturity gaps that span multiple TPRM lifecycle stages, not isolated fixes."
    );
  } else if (belowBench >= 3) {
    out.push(
      "Several domains trail the benchmark—prioritize cross-cutting governance and evidence cadence before scaling automation."
    );
  }

  if (weakEv >= 4) {
    out.push(
      "Evidence strength is Weak across multiple domains—interpret scores as directional until artifact depth is uplifted; assessment confidence is limited."
    );
  } else if (weakEv >= 2) {
    out.push(
      "Mixed evidence quality signals pockets where attestations may not yet withstand regulatory or audit challenge."
    );
  }

  const gov = entries.find((e) => e.id === "governance");
  const seg = entries.find((e) => e.id === "segmentation");
  if (gov && seg && gov.score <= 2 && seg.score <= 2) {
    out.push(
      "Governance and risk segmentation are both stretched—foundational controls (tiering, ownership, appetite translation) should take priority over peripheral tooling spend."
    );
  }

  const mon = entries.find((e) => e.id === "monitoring");
  const rep = entries.find((e) => e.id === "reporting");
  if ((mon && mon.score <= 2) || (rep && rep.score <= 2)) {
    out.push(
      "Monitoring and/or reporting maturity is low—oversight is likely reactive; institute MI cadence, breach narratives, and ageing discipline before expanding scope."
    );
  }

  const tech = entries.find((e) => e.id === "technology");
  if (tech && tech.score <= 2) {
    out.push(
      "Technology and data enablement are behind—scaling monitoring and assurance across volume will be constrained without workflow and lineage uplift."
    );
  }

  if (out.length === 0) {
    out.push(
      "Portfolio posture is relatively balanced versus benchmark—focus on sustaining evidence depth and industrializing playbooks rather than wholesale redesign."
    );
  }

  return out.slice(0, 5);
}

/** Radar helpers: polar coords, angle i of n domains starting top */
function radarPoint(cx, cy, R, valueOn5, index, n) {
  const angle = (2 * Math.PI * index) / n - Math.PI / 2;
  const r = (valueOn5 / 5) * R;
  return {
    x: cx + r * Math.cos(angle),
    y: cy + r * Math.sin(angle),
    lx: cx + (R + 36) * Math.cos(angle),
    ly: cy + (R + 36) * Math.sin(angle),
    angle,
  };
}

function polygonPointsFromEntries(entries, cx, cy, R, valueFn) {
  const n = entries.length;
  return entries
    .map((e, i) => {
      const { x, y } = radarPoint(cx, cy, R, valueFn(e), i, n);
      return `${x},${y}`;
    })
    .join(" ");
}

// --- Theme tokens (inline reuse) ---

const C = {
  bg: "#f1f5f9",
  surface: "#ffffff",
  border: "#e2e8f0",
  text: "#0f172a",
  muted: "#64748b",
  accent: "#1e40af",
  accentSoft: "#eff6ff",
  sidebar: "#0f172a",
  sidebarMuted: "#94a3b8",
  danger: "#b45309",
  success: "#047857",
};

function initialDomainRows() {
  const o = {};
  DOMAINS.forEach((d) => {
    o[d.id] = { score: 3, evidence: "Moderate", notes: "", recommendationDraft: "" };
  });
  return o;
}

const DEMO_ASSESSMENT_PROFILE = {
  clientName: "Summit Metropolitan Financial Group",
  industry: "Financial Services",
  companySize: "Enterprise",
  regulatoryIntensity: "Very High",
  thirdPartyVolume: "Very High",
  geographicFootprint: "Regional",
};

const DEMO_ASSESSMENT_TYPE = "Full maturity assessment";

const DEMO_SCOPE_NOTES =
  "FY26 enterprise TPRM cycle covering Tier-1 material outsourcers, critical SaaS concentration, payment processor resilience, and cross-border data flows for retail and commercial banking.";

const DEMO_SUPPORTING_TEXT = `Board-approved third-party risk appetite statement (FY26) and concentration limits. Enterprise vendor inventory reconciled Q4: 412 active vendors; 38 Tier-1 / material relationships identified across legal entities.
Recent regulatory dialogue referenced cloud service provider concentration and incomplete exit / transition planning for two legacy payment processors.
Artifacts on file: vendor risk committee minutes (redacted), SOC 2 Type II summaries for the top five processors, integrated resilience test outcomes, and enterprise issue-management aging (90-day roll-up).`;

/** Realistic per-domain scores, evidence strength, and assessor notes for the demo workspace. */
const DEMO_DOMAIN_SEED = [
  {
    id: "governance",
    score: 3,
    evidence: "Moderate",
    notes:
      "TPRM steering forum meets quarterly; charter exists but board reporting packs vary in depth by quarter. Exception governance is documented; two business units show waiver aging beyond policy SLAs.",
    recommendationDraft: "",
  },
  {
    id: "inventory",
    score: 2,
    evidence: "Weak",
    notes:
      "Golden vendor record initiative underway; legal-entity / vendor-ID mismatches persist after acquisitions. Dormant vendor detection runs monthly but closure discipline is uneven across regions.",
    recommendationDraft: "",
  },
  {
    id: "segmentation",
    score: 3,
    evidence: "Moderate",
    notes:
      "Tiering criteria published and mapped to diligence depth. Residual risk linkage is improving but not consistently enforced for fast-growing SaaS vendors added via shadow IT channels.",
    recommendationDraft: "",
  },
  {
    id: "dueDiligence",
    score: 2,
    evidence: "Moderate",
    notes:
      "Tier-standard diligence packages exist; QA sampling for onboarding files shows inconsistent depth for non-US entities. Cyber sub-assessments sometimes lag contract execution.",
    recommendationDraft: "",
  },
  {
    id: "contracting",
    score: 4,
    evidence: "Strong",
    notes:
      "Security and resilience clauses aligned to tier for most Tier-1 contracts; enterprise clause library refreshed annually. Subsidiary addenda occasionally lag template updates by one release.",
    recommendationDraft: "",
  },
  {
    id: "monitoring",
    score: 3,
    evidence: "Moderate",
    notes:
      "KRIs and attestations in place for Tier-1; monitoring cadence slips for lower tiers during peak procurement cycles. Automated connectors cover ~60% of critical vendors.",
    recommendationDraft: "",
  },
  {
    id: "issues",
    score: 2,
    evidence: "Weak",
    notes:
      "Issue lifecycle documented; validation of closure evidence is inconsistent for Sev-2 items. Linkage from monitoring alerts into formal issues is manual for two core platforms.",
    recommendationDraft: "",
  },
  {
    id: "reporting",
    score: 3,
    evidence: "Moderate",
    notes:
      "Management MI pack reviewed monthly; board deck narrative quality improved. Audit trail for tier changes is partial—some decisions captured only in email threads.",
    recommendationDraft: "",
  },
  {
    id: "technology",
    score: 4,
    evidence: "Moderate",
    notes:
      "Workflow tooling deployed for Tier-1 lifecycle; API integrations to CMDB and GRC progressing. Data lineage for vendor attributes still fragmented between procurement and risk systems.",
    recommendationDraft: "",
  },
  {
    id: "regulatory",
    score: 3,
    evidence: "Strong",
    notes:
      "Supervisory expectations mapped to control owners; jurisdictional nuance documented for primary markets. Evidence packs for exams are strong but assembly time remains high.",
    recommendationDraft: "",
  },
];

function buildDemoDomainRows() {
  const o = initialDomainRows();
  for (const e of DEMO_DOMAIN_SEED) {
    const ev = ["Weak", "Moderate", "Strong"].includes(e.evidence) ? e.evidence : "Moderate";
    o[e.id] = {
      score: clampScore(e.score),
      evidence: ev,
      notes: typeof e.notes === "string" ? e.notes : "",
      recommendationDraft: typeof e.recommendationDraft === "string" ? e.recommendationDraft : "",
    };
  }
  return o;
}

const defaultProfile = {
  clientName: "",
  industry: "Financial Services",
  companySize: "Enterprise",
  regulatoryIntensity: "High",
  thirdPartyVolume: "High",
  geographicFootprint: "Regional",
};

const TPRM_ASSESSMENT_DRAFT_KEY = "tprmAssessmentDraft";
const PERSISTENCE_DRAFT_VERSION = 2;

function mergeLoadedDomainRows(raw) {
  const base = initialDomainRows();
  if (!raw || typeof raw !== "object") return base;
  DOMAINS.forEach((d) => {
    const row = raw[d.id];
    if (!row || typeof row !== "object") return;
    const ev = ["Weak", "Moderate", "Strong"].includes(row.evidence) ? row.evidence : base[d.id].evidence;
    base[d.id] = {
      score: clampScore(row.score ?? base[d.id].score),
      evidence: ev,
      notes: typeof row.notes === "string" ? row.notes : "",
      recommendationDraft:
        typeof row.recommendationDraft === "string" ? row.recommendationDraft : base[d.id].recommendationDraft || "",
    };
  });
  return base;
}

function mergeLoadedProfile(raw) {
  if (!raw || typeof raw !== "object") return { ...defaultProfile };
  return {
    ...defaultProfile,
    clientName: typeof raw.clientName === "string" ? raw.clientName : defaultProfile.clientName,
    industry: PROFILE_OPTIONS.industry.includes(raw.industry) ? raw.industry : defaultProfile.industry,
    companySize: PROFILE_OPTIONS.companySize.includes(raw.companySize) ? raw.companySize : defaultProfile.companySize,
    regulatoryIntensity: PROFILE_OPTIONS.regulatoryIntensity.includes(raw.regulatoryIntensity)
      ? raw.regulatoryIntensity
      : defaultProfile.regulatoryIntensity,
    thirdPartyVolume: PROFILE_OPTIONS.thirdPartyVolume.includes(raw.thirdPartyVolume)
      ? raw.thirdPartyVolume
      : defaultProfile.thirdPartyVolume,
    geographicFootprint: PROFILE_OPTIONS.geographicFootprint.includes(raw.geographicFootprint)
      ? raw.geographicFootprint
      : defaultProfile.geographicFootprint,
  };
}

function mergeLoadedAiSuggestionsByDomain(raw) {
  const out = {};
  if (!raw || typeof raw !== "object") return out;
  DOMAINS.forEach((d) => {
    const cell = raw[d.id];
    if (!cell || typeof cell !== "object") return;
    const sug = cell.suggestion;
    if (!sug || typeof sug !== "object") return;
    const st = String(cell.status || "Suggested");
    const status = st === "Accepted" || st === "Overridden" || st === "Suggested" ? st : "Suggested";
    out[d.id] = {
      status,
      suggestion: normalizeEvidenceAiSuggestion(sug),
      source: cell.source === "openai" ? "openai" : "mock",
    };
  });
  return out;
}

function mergeLoadedReviewFollowUp(raw) {
  const out = {};
  if (!raw || typeof raw !== "object") return out;
  DOMAINS.forEach((d) => {
    if (raw[d.id] === true) out[d.id] = true;
  });
  return out;
}

function mergeLoadedReviewerNotes(raw) {
  const out = {};
  if (!raw || typeof raw !== "object") return out;
  DOMAINS.forEach((d) => {
    if (typeof raw[d.id] === "string") out[d.id] = raw[d.id];
  });
  return out;
}

function mergeLoadedAiNarratives(raw) {
  if (!raw || typeof raw !== "object") return null;
  const executiveSummary = typeof raw.executiveSummary === "string" ? raw.executiveSummary : "";
  const riskNarrative = typeof raw.riskNarrative === "string" ? raw.riskNarrative : "";
  const recommendationsText = typeof raw.recommendationsText === "string" ? raw.recommendationsText : "";
  const source = raw.source === "openai" ? "openai" : "placeholder";
  if (!executiveSummary.trim() && !riskNarrative.trim() && !recommendationsText.trim()) return null;
  return { executiveSummary, riskNarrative, recommendationsText, source };
}

function mergeLoadedLibrarySuggestions(raw) {
  const out = {};
  if (!raw || typeof raw !== "object") return out;
  DOMAINS.forEach((d) => {
    const cell = raw[d.id];
    if (!cell || typeof cell !== "object") return;
    const dismissed = Boolean(cell.dismissed);
    const p = cell.payload;
    let payload = null;
    if (p === false || p?.match === false) {
      payload = { match: false };
    } else if (p && typeof p === "object" && p.match === true) {
      const conf = ["Strong match", "Moderate match", "Weak match"].includes(p.confidence)
        ? p.confidence
        : "Weak match";
      const ev = Array.isArray(p.evidenceExamples) ? p.evidenceExamples.map(String).filter(Boolean).slice(0, 8) : [];
      payload = {
        match: true,
        confidence: conf,
        findingText: typeof p.findingText === "string" ? p.findingText : "",
        recommendationText: typeof p.recommendationText === "string" ? p.recommendationText : "",
        evidenceExamples: ev,
        roadmapAction: typeof p.roadmapAction === "string" ? p.roadmapAction : "",
        similarPatternNote: typeof p.similarPatternNote === "string" ? p.similarPatternNote : "",
      };
    }
    out[d.id] = { dismissed, payload };
  });
  return out;
}

const NAV_ITEMS = [
  { id: "commandCenter", label: "Command Center" },
  { id: "assessment", label: "Assessment" },
  { id: "results", label: "Results" },
  { id: "analytics", label: "Visual Analytics" },
  { id: "library", label: "Library" },
  { id: "roadmap", label: "Roadmap" },
  { id: "reviewQueue", label: "Review Queue" },
  { id: "settings", label: "Settings" },
];

const NAV_PAGE_IDS = new Set(NAV_ITEMS.map((i) => i.id));

function sanitizeActivePage(p) {
  return typeof p === "string" && NAV_PAGE_IDS.has(p) ? p : "assessment";
}

/** Assessment memory — reusable snippets from completed runs (localStorage). */
const LS_FINDINGS_LIBRARY = "findingsLibrary";
const LS_RECOMMENDATIONS_LIBRARY = "recommendationsLibrary";
const LS_NARRATIVES_LIBRARY = "narrativesLibrary";
const MAX_ASSESSMENT_MEMORY_ITEMS = 200;

function readJsonLibrary(key) {
  if (typeof localStorage === "undefined") return [];
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function writeJsonLibrary(key, arr) {
  if (typeof localStorage === "undefined") return;
  const trimmed = arr.slice(-MAX_ASSESSMENT_MEMORY_ITEMS);
  localStorage.setItem(key, JSON.stringify(trimmed));
}

function appendAssessmentMemoryLibraries(profile, assessment, narratives) {
  if (typeof localStorage === "undefined") return;
  const findings = readJsonLibrary(LS_FINDINGS_LIBRARY);
  const recommendations = readJsonLibrary(LS_RECOMMENDATIONS_LIBRARY);
  const narrativesLib = readJsonLibrary(LS_NARRATIVES_LIBRARY);
  const ts = Date.now();

  assessment.entries.forEach((e, idx) => {
    const mat = matrixMaturityBand(e.score);
    const blocks = deriveDomainConsultingBlocks(e, profile);
    const evPat = `score=${e.score}; evidence=${e.evidence}; notesChars=${(e.notes || "").length}`;
    const baseId = `${ts}-${e.id}-${idx}`;
    const ctx = {
      industry: profile.industry || "",
      assessmentType: profile.assessmentType || "",
    };
    findings.push({
      id: `${baseId}-f`,
      domainId: e.id,
      domain: e.label,
      maturityLevel: mat.label,
      score: e.score,
      evidenceStrength: e.evidence,
      findingText: String(blocks.specificFinding || "").slice(0, 4000),
      recommendationText: String(blocks.recommendation || "").slice(0, 4000),
      evidencePatterns: evPat,
      savedAt: new Date().toISOString(),
      ...ctx,
    });
    recommendations.push({
      id: `${baseId}-r`,
      domainId: e.id,
      domain: e.label,
      maturityLevel: mat.label,
      score: e.score,
      evidenceStrength: e.evidence,
      findingText: "",
      recommendationText: String(blocks.recommendation || "").slice(0, 4000),
      evidencePatterns: evPat,
      savedAt: new Date().toISOString(),
      ...ctx,
    });
  });

  const exec = String(narratives?.executiveSummary || "").trim();
  const risk = String(narratives?.riskNarrative || "").trim();
  const recsText = String(narratives?.recommendationsText || "").trim();
  const portCtx = {
    industry: profile.industry || "",
    assessmentType: profile.assessmentType || "",
  };
  if (exec) {
    narrativesLib.push({
      id: `${ts}-exec`,
      domainId: "portfolio",
      domain: "Executive summary",
      maturityLevel: assessment.maturityBand,
      score: Math.round(Number(assessment.overallScore) || 0),
      evidenceStrength: "Portfolio",
      findingText: exec.slice(0, 4500),
      recommendationText: "",
      evidencePatterns: `overallScore=${assessment.overallScore}`,
      savedAt: new Date().toISOString(),
      ...portCtx,
    });
  }
  if (risk) {
    narrativesLib.push({
      id: `${ts}-risk`,
      domainId: "portfolio",
      domain: "Risk narrative",
      maturityLevel: assessment.maturityBand,
      score: Math.round(Number(assessment.overallScore) || 0),
      evidenceStrength: "Portfolio",
      findingText: risk.slice(0, 4500),
      recommendationText: "",
      evidencePatterns: `overallScore=${assessment.overallScore}`,
      savedAt: new Date().toISOString(),
      ...portCtx,
    });
  }
  if (recsText) {
    splitRecommendationLines(recsText).forEach((line, i) => {
      narrativesLib.push({
        id: `${ts}-rec-${i}`,
        domainId: "portfolio",
        domain: "Recommendations (portfolio)",
        maturityLevel: assessment.maturityBand,
        score: Math.round(Number(assessment.overallScore) || 0),
        evidenceStrength: "Portfolio",
        findingText: line.slice(0, 2000),
        recommendationText: line.slice(0, 2000),
        evidencePatterns: `overallScore=${assessment.overallScore}`,
        savedAt: new Date().toISOString(),
        ...portCtx,
      });
    });
  }

  writeJsonLibrary(LS_FINDINGS_LIBRARY, findings);
  writeJsonLibrary(LS_RECOMMENDATIONS_LIBRARY, recommendations);
  writeJsonLibrary(LS_NARRATIVES_LIBRARY, narrativesLib);
}

function formatMemoryLibraryForPrompt(profile, domainRows) {
  if (typeof localStorage === "undefined") return "";
  const findings = readJsonLibrary(LS_FINDINGS_LIBRARY);
  const recommendations = readJsonLibrary(LS_RECOMMENDATIONS_LIBRARY);
  const narratives = readJsonLibrary(LS_NARRATIVES_LIBRARY);
  const chunks = [];

  DOMAINS.forEach((d) => {
    const row = domainRows[d.id];
    if (!row) return;
    const band = matrixMaturityBand(clampScore(row.score)).label;
    const ev = row.evidence || "Moderate";
    const match = (item) =>
      (item.domainId === d.id || item.domain === d.label) &&
      item.maturityLevel === band &&
      item.evidenceStrength === ev;

    findings.filter(match).slice(0, 2).forEach((it) => {
      chunks.push(
        `[Domain: ${it.domain} | ${it.maturityLevel} | ${it.evidenceStrength}]\nFinding: ${it.findingText.slice(0, 700)}`
      );
    });
    recommendations.filter(match).slice(0, 2).forEach((it) => {
      chunks.push(
        `[Domain: ${it.domain} | ${it.maturityLevel} | ${it.evidenceStrength}]\nRecommendation: ${it.recommendationText.slice(0, 700)}`
      );
    });
  });

  narratives
    .filter((it) => it.domainId === "portfolio")
    .slice(-4)
    .forEach((it) => {
      chunks.push(`[Portfolio narrative — ${it.maturityLevel}]\n${it.findingText.slice(0, 650)}`);
    });

  if (chunks.length === 0) return "";
  const body = chunks.join("\n---\n").slice(0, 4500);
  return `Here are similar past findings and recommendations:\n${body}\n\nUse these as reference, adapt to current assessment.`;
}

const LIBRARY_CONFIDENCE_ORDER = { "Strong match": 3, "Moderate match": 2, "Weak match": 1 };

function maxLibraryConfidence(a, b) {
  return LIBRARY_CONFIDENCE_ORDER[a] >= LIBRARY_CONFIDENCE_ORDER[b] ? a : b;
}

function sameDomainLibraryItem(domainMeta, item) {
  if (!item) return false;
  return item.domainId === domainMeta.id || item.domain === domainMeta.label;
}

function maturitySimilarForLibrary(row, item) {
  const band = matrixMaturityBand(clampScore(row.score)).label;
  if (item.maturityLevel === band) return true;
  const order = { Low: 0, Moderate: 1, Advanced: 2 };
  const a = order[item.maturityLevel];
  const b = order[band];
  if (a !== undefined && b !== undefined) return Math.abs(a - b) <= 1;
  const rs = clampScore(row.score);
  const is = Number(item.score);
  if (Number.isFinite(is)) return Math.abs(rs - is) <= 1;
  return false;
}

/** Strong / Moderate / Weak per spec: weak = same domain only (floor). */
function classifyLibraryItemConfidence(row, item) {
  const band = matrixMaturityBand(clampScore(row.score)).label;
  const ev = row.evidence || "Moderate";
  if (item.maturityLevel === band && item.evidenceStrength === ev) return "Strong match";
  if (maturitySimilarForLibrary(row, item)) return "Moderate match";
  return "Weak match";
}

function libraryContextBoost(profile, assessmentType, item) {
  let n = 0;
  if (item.industry && profile?.industry && item.industry === profile.industry) n += 2;
  if (item.assessmentType && assessmentType && item.assessmentType === assessmentType) n += 2;
  return n;
}

function pickBestLibraryFieldByTier(pool, row, field, profile, assessmentType) {
  for (const tier of ["Strong match", "Moderate match", "Weak match"]) {
    const tierItems = pool.filter((it) => classifyLibraryItemConfidence(row, it) === tier);
    if (!tierItems.length) continue;
    const sorted = [...tierItems].sort(
      (a, b) =>
        libraryContextBoost(profile, assessmentType, b) - libraryContextBoost(profile, assessmentType, a)
    );
    const texts = sorted
      .map((it) => it[field])
      .filter((t) => String(t || "").trim().length > 12);
    if (texts.length) {
      const text = [...texts].sort((a, b) => String(b).length - String(a).length)[0];
      return { text: String(text).trim(), tier };
    }
  }
  return { text: "", tier: "Weak match" };
}

function buildSuggestedRoadmapLine(domainLabel, liveAssessment) {
  const road = liveAssessment?.roadmap;
  const first =
    (road?.d30 && road.d30[0]) || (road?.d60 && road.d60[0]) || (road?.d90 && road.d90[0]) || null;
  if (first) {
    return `${first} (Tailor timing and owners to ${domainLabel} materiality.)`;
  }
  return `Assign an accountable owner, 30-day outcomes, and evidence milestones for ${domainLabel}, scaled to tier and regulatory intensity.`;
}

function collectEvidenceExamples(pool, row, domainId) {
  const seen = new Set();
  const out = [];
  const sorted = [...pool].sort(
    (a, b) =>
      LIBRARY_CONFIDENCE_ORDER[classifyLibraryItemConfidence(row, b)] -
      LIBRARY_CONFIDENCE_ORDER[classifyLibraryItemConfidence(row, a)]
  );
  for (const it of sorted) {
    const ex = String(it.evidencePatterns || "").trim();
    if (ex && !seen.has(ex)) {
      seen.add(ex);
      out.push(ex);
      if (out.length >= 4) break;
    }
  }
  const gap0 = TPRM_DOMAIN_LIBRARY[domainId]?.commonGaps?.[0];
  if (gap0 && !seen.has(gap0)) out.unshift(`Common gap pattern (reference): ${gap0}`);
  return out.slice(0, 5);
}

/**
 * Returns reusable suggestion from local library or null when no same-domain entries exist.
 */
function buildLibrarySuggestionForDomain(domainMeta, row, profile, assessmentType, liveAssessment) {
  const findings = readJsonLibrary(LS_FINDINGS_LIBRARY);
  const recs = readJsonLibrary(LS_RECOMMENDATIONS_LIBRARY);
  const nars = readJsonLibrary(LS_NARRATIVES_LIBRARY);
  const pool = [...findings, ...recs, ...nars].filter((it) => sameDomainLibraryItem(domainMeta, it));
  if (pool.length === 0) return null;

  const fPick = pickBestLibraryFieldByTier(pool, row, "findingText", profile, assessmentType);
  const rPick = pickBestLibraryFieldByTier(pool, row, "recommendationText", profile, assessmentType);
  let findingText = fPick.text;
  let recommendationText = rPick.text;
  if (!findingText) findingText = pool.map((it) => it.findingText).find((t) => String(t || "").trim().length > 12) || "";
  if (!recommendationText) {
    recommendationText =
      pool.map((it) => it.recommendationText).find((t) => String(t || "").trim().length > 12) || "";
  }

  const panelConfidence = maxLibraryConfidence(fPick.tier, rPick.tier);

  const evidenceExamples = collectEvidenceExamples(pool, row, domainMeta.id);
  const roadmapAction = buildSuggestedRoadmapLine(domainMeta.label, liveAssessment);

  const indHits = pool.some((it) => it.industry && profile?.industry && it.industry === profile.industry);
  const typeHits = pool.some(
    (it) => it.assessmentType && assessmentType && it.assessmentType === assessmentType
  );
  const similarPatternNote = [
    `Similar pattern: ${panelConfidence} for ${domainMeta.label} against stored assessments.`,
    indHits
      ? `Industry (${profile?.industry || "profile"}) matches some library entries — reuse with light tailoring.`
      : `Library entries may span industries — validate examples against this client context.`,
    typeHits
      ? `Assessment type (“${assessmentType}”) appears in memory — good narrative consistency anchor.`
      : `Blend language across assessment types where controls are materially similar.`,
  ].join(" ");

  return {
    confidence: panelConfidence,
    findingText,
    recommendationText,
    evidenceExamples,
    roadmapAction,
    similarPatternNote,
  };
}

function normalizeRecPatternKey(text) {
  const s = String(text || "").trim().replace(/\s+/g, " ");
  if (s.length < 28) return "";
  return s.slice(0, 140);
}

function computeLibraryReuseSummary(profile, assessmentType, domainRows, liveAssessment) {
  const strong = [];
  const moderate = [];
  const uncovered = [];
  DOMAINS.forEach((d) => {
    const row = domainRows[d.id] || {};
    const sug = buildLibrarySuggestionForDomain(d, row, profile, assessmentType, liveAssessment);
    if (!sug) {
      uncovered.push(d.label);
    } else if (sug.confidence === "Strong match") strong.push(d.label);
    else if (sug.confidence === "Moderate match") moderate.push(d.label);
  });
  const recItems = readJsonLibrary(LS_RECOMMENDATIONS_LIBRARY);
  const counts = {};
  recItems.forEach((it) => {
    const k = normalizeRecPatternKey(it.recommendationText);
    if (!k) return;
    counts[k] = (counts[k] || 0) + 1;
  });
  const mostReused = Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 6)
    .map(([text, count]) => ({ text, count }));
  return { strong, moderate, uncovered, mostReused };
}

const PAGE_HEADER = {
  assessment: {
    eyebrow: "Third-party risk · Workspace",
    title: "TPRM Maturity Assessment Engine",
  },
  commandCenter: {
    eyebrow: "Leadership · Engagement",
    title: "Engagement Command Center",
  },
  results: {
    eyebrow: "Outputs · Leadership pack",
    title: "Assessment results",
  },
  analytics: {
    eyebrow: "Analytics · Portfolio visuals",
    title: "Visual Analytics",
  },
  library: {
    eyebrow: "Knowledge · Plays & patterns",
    title: "TPRM content library",
  },
  roadmap: {
    eyebrow: "Execution · Horizons",
    title: "Remediation roadmap",
  },
  reviewQueue: {
    eyebrow: "Governance · Manager review",
    title: "Manager review queue",
  },
  settings: {
    eyebrow: "Configuration",
    title: "Assessment settings",
  },
};

function PillBadge({ children, bg, fg, bd }) {
  return (
    <span
      style={{
        display: "inline-block",
        padding: "4px 10px",
        borderRadius: 999,
        fontSize: 11,
        fontWeight: 800,
        letterSpacing: "0.02em",
        background: bg,
        color: fg,
        border: bd ? `1px solid ${bd}` : "none",
      }}
    >
      {children}
    </span>
  );
}

function ExecutiveScorecard({ liveAssessment, maturityBadgeColor }) {
  const kpi = computeExecutiveKPIs(liveAssessment.entries, liveAssessment.overallScore);
  const overallBenchGap = liveAssessment.overallScore - BENCHMARK_SCORE;

  const cards = [
    {
      label: "Overall maturity score",
      value: String(liveAssessment.overallScore),
      sub: "/ 5 mean",
      badge:
        overallBenchGap >= 0
          ? { text: "Above benchmark", bg: "#ecfdf5", fg: "#15803d" }
          : { text: "Below benchmark", bg: "#fef2f2", fg: "#b91c1c" },
    },
    {
      label: "Domains below benchmark",
      value: String(kpi.belowBench),
      sub: `of ${liveAssessment.entries.length} domains`,
      badge:
        kpi.belowBench >= 5
          ? { text: "Elevated exposure", bg: "#fef2f2", fg: "#b91c1c" }
          : kpi.belowBench >= 2
            ? { text: "Watch list", bg: "#fffbeb", fg: "#b45309" }
            : { text: "Controlled", bg: "#ecfdf5", fg: "#15803d" },
    },
    {
      label: "Critical gaps",
      value: String(kpi.criticalGaps),
      sub: "Score 1–2 + Weak evidence",
      badge:
        kpi.criticalGaps > 0
          ? { text: "Action required", bg: "#fef2f2", fg: "#991b1b" }
          : { text: "Clear", bg: "#ecfdf5", fg: "#15803d" },
    },
    {
      label: "Evidence confidence",
      value: `${kpi.confidencePct}%`,
      sub: "Evidence-weighted assurance",
      badge:
        kpi.confidencePct >= 67
          ? { text: "Strong", bg: "#ecfdf5", fg: "#15803d" }
          : kpi.confidencePct >= 40
            ? { text: "Mixed", bg: "#fffbeb", fg: "#b45309" }
            : { text: "Limited", bg: "#fef2f2", fg: "#b91c1c" },
    },
  ];

  return (
    <section style={{ marginBottom: 22 }}>
      <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: "0.1em", color: C.muted, marginBottom: 12 }}>
        EXECUTIVE SCORECARD
      </div>
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))",
          gap: 14,
        }}
      >
        {cards.map((c) => (
          <div
            key={c.label}
            style={{
              borderRadius: 16,
              padding: "20px 18px",
              background: "linear-gradient(165deg, #ffffff 0%, #f8fafc 100%)",
              border: `1px solid ${C.border}`,
              boxShadow: "0 8px 28px rgba(15,23,42,0.08)",
              position: "relative",
              overflow: "hidden",
            }}
          >
            <div
              style={{
                position: "absolute",
                top: 0,
                left: 0,
                right: 0,
                height: 3,
                background: `linear-gradient(90deg, ${C.accent}, #60a5fa)`,
                opacity: 0.85,
              }}
            />
            <div style={{ fontSize: 11, color: C.muted, fontWeight: 700, marginBottom: 8, lineHeight: 1.3 }}>{c.label}</div>
            <div style={{ fontSize: 36, fontWeight: 900, letterSpacing: "-0.03em", lineHeight: 1.05 }}>{c.value}</div>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{c.sub}</div>
            <div style={{ marginTop: 12 }}>
              <span
                style={{
                  display: "inline-block",
                  padding: "4px 10px",
                  borderRadius: 999,
                  fontSize: 10,
                  fontWeight: 800,
                  letterSpacing: "0.04em",
                  background: c.badge.bg,
                  color: c.badge.fg,
                }}
              >
                {c.badge.text}
              </span>
            </div>
          </div>
        ))}
      </div>
      <div style={{ marginTop: 12, fontSize: 12, color: C.muted }}>
        Maturity band:{" "}
        <strong style={{ color: maturityBadgeColor }}>{liveAssessment.maturityBand}</strong> ·{" "}
        {liveAssessment.maturityTier} tier
      </div>
    </section>
  );
}

function CommandCenterPageContent({
  liveAssessment,
  completionPct,
  model,
  managerReviewStatus,
  onManagerReviewStatus,
  analysisLoading,
  analysisStatus,
  reportQualityWarningCount,
  onNavigate,
  onAnalyzeEvidence,
  onAcceptAllHighConfidence,
  onGenerateReport,
}) {
  const stageDot = (done) => ({
    width: 22,
    height: 22,
    borderRadius: "50%",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: 12,
    fontWeight: 900,
    flexShrink: 0,
    background: done ? "#16a34a" : "#e2e8f0",
    color: done ? "#fff" : "#64748b",
    border: done ? "none" : "1px solid #cbd5e1",
  });

  const aiStage = model.stages.find((s) => s.key === "ai");
  const qaStage = model.stages.find((s) => s.key === "qa");
  const aiReviewed = Boolean(aiStage?.done);
  const qaComplete = Boolean(qaStage?.done);

  const riskVisual =
    liveAssessment.maturityTier === "Low"
      ? {
          headline: "Elevated exposure",
          sub: "Low program maturity — prioritize weakest domains and evidence depth.",
          accent: "#b91c1c",
          rail: "#fecaca",
          fill: "#fef2f2",
          chip: "High priority",
          chipBg: "#fee2e2",
        }
      : liveAssessment.maturityTier === "Moderate"
        ? {
            headline: "Moderate exposure",
            sub: "Mixed maturity — tighten monitoring and close gaps in Tier-1 vendors.",
            accent: "#c2410c",
            rail: "#fed7aa",
            fill: "#fffbeb",
            chip: "Watch",
            chipBg: "#ffedd5",
          }
        : {
            headline: "Controlled posture",
            sub: "Stronger mean maturity — sustain automation and peer benchmarking.",
            accent: "#047857",
            rail: "#a7f3d0",
            fill: "#ecfdf5",
            chip: "Stable",
            chipBg: "#d1fae5",
          };

  const glanceCard = {
    borderRadius: 16,
    padding: "18px 16px",
    border: `1px solid ${C.border}`,
    background: C.surface,
    boxShadow: "0 4px 18px rgba(15,23,42,0.06)",
    minHeight: 148,
    display: "flex",
    flexDirection: "column",
  };

  return (
    <>
      <SectionCard
        title="Engagement Command Center"
        subtitle="Leadership view of assessment progress, risk posture, AI review, QA, and export readiness."
      >
        <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: "0.12em", color: C.muted, marginBottom: 12 }}>
          AT A GLANCE
        </div>
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
            gap: 16,
            marginBottom: 24,
          }}
        >
          <div style={{ ...glanceCard, borderTop: `4px solid ${C.accent}` }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: C.muted }}>Overall maturity score</div>
            <div style={{ display: "flex", alignItems: "baseline", gap: 6, marginTop: 8 }}>
              <span style={{ fontSize: 40, fontWeight: 900, letterSpacing: "-0.03em", color: C.text }}>
                {liveAssessment.overallScore}
              </span>
              <span style={{ fontSize: 18, fontWeight: 700, color: C.muted }}>/ 5</span>
            </div>
            <div style={{ fontSize: 13, color: C.text, marginTop: 6, fontWeight: 600 }}>
              {liveAssessment.maturityBand}
            </div>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{liveAssessment.maturityTier} tier mean</div>
          </div>

          <div style={{ ...glanceCard, borderTop: "4px solid #16a34a" }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: C.muted }}>Workspace completion</div>
            <div style={{ fontSize: 40, fontWeight: 900, marginTop: 8, color: "#15803d" }}>{completionPct}%</div>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 6 }}>Score + evidence strength per domain</div>
            <div
              style={{
                marginTop: 12,
                height: 8,
                borderRadius: 999,
                background: "#e2e8f0",
                overflow: "hidden",
              }}
            >
              <div
                style={{
                  width: `${completionPct}%`,
                  height: "100%",
                  borderRadius: 999,
                  background: "linear-gradient(90deg, #22c55e, #16a34a)",
                  transition: "width 0.25s ease",
                }}
              />
            </div>
          </div>

          <div
            style={{
              ...glanceCard,
              borderTop: `4px solid ${riskVisual.accent}`,
              background: riskVisual.fill,
              borderColor: riskVisual.rail,
            }}
          >
            <div style={{ fontSize: 12, fontWeight: 700, color: C.muted }}>Risk level</div>
            <div style={{ fontSize: 18, fontWeight: 900, marginTop: 10, color: riskVisual.accent }}>
              {riskVisual.headline}
            </div>
            <p style={{ margin: "10px 0 0", fontSize: 12, lineHeight: 1.5, color: C.text, flex: 1 }}>{riskVisual.sub}</p>
            <div style={{ marginTop: 10 }}>
              <span
                style={{
                  display: "inline-block",
                  padding: "4px 10px",
                  borderRadius: 999,
                  fontSize: 10,
                  fontWeight: 800,
                  letterSpacing: "0.04em",
                  background: riskVisual.chipBg,
                  color: riskVisual.accent,
                }}
              >
                {riskVisual.chip}
              </span>
              {model.criticalGaps > 0 && (
                <span style={{ fontSize: 11, color: C.muted, marginLeft: 8 }}>
                  {model.criticalGaps} critical gap{model.criticalGaps === 1 ? "" : "s"}
                </span>
              )}
            </div>
          </div>

          <div style={{ ...glanceCard, borderTop: "4px solid #6366f1" }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: C.muted }}>Review status</div>
            <div style={{ marginTop: 14, display: "flex", flexDirection: "column", gap: 10 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                <span style={{ fontSize: 11, fontWeight: 700, color: C.muted, width: 72 }}>AI</span>
                <span
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    padding: "6px 12px",
                    borderRadius: 999,
                    fontSize: 12,
                    fontWeight: 800,
                    background: aiReviewed ? "#dcfce7" : "#f1f5f9",
                    color: aiReviewed ? "#166534" : "#64748b",
                    border: aiReviewed ? "1px solid #86efac" : "1px solid #cbd5e1",
                  }}
                >
                  {aiReviewed ? "AI reviewed" : "AI pending"}
                </span>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                <span style={{ fontSize: 11, fontWeight: 700, color: C.muted, width: 72 }}>QA</span>
                <span
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    padding: "6px 12px",
                    borderRadius: 999,
                    fontSize: 12,
                    fontWeight: 800,
                    background: qaComplete ? "#dcfce7" : "#fff7ed",
                    color: qaComplete ? "#166534" : "#c2410c",
                    border: qaComplete ? "1px solid #86efac" : "1px solid #fdba74",
                  }}
                >
                  {qaComplete ? "QA complete" : "QA attention"}
                </span>
              </div>
            </div>
            <div style={{ fontSize: 11, color: C.muted, marginTop: 12, lineHeight: 1.45 }}>
              AI = evidence suggestions run for every domain. QA = no unresolved critical pre-report checks.
            </div>
          </div>
        </div>

        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))",
            gap: 14,
            marginBottom: 22,
          }}
        >
          {[
            {
              label: "Domains needing review",
              value: String(model.domainsNeedingReview),
              sub: "Queue, QA, AI divergence, or priority hotspots",
              badge: { text: "Queue", bg: "#fff7ed", fg: "#9a3412" },
            },
            {
              label: "Critical gaps",
              value: String(model.criticalGaps),
              sub: "Critical priority or score ≤2 with Weak evidence",
              badge: { text: "Risk", bg: "#fef2f2", fg: "#b91c1c" },
            },
            {
              label: "Report readiness",
              value: model.reportReady ? "Ready" : "Not ready",
              sub: model.reportReady
                ? "All gates cleared for export"
                : model.reportReadinessReasons[0] || "Complete remaining gates",
              badge: {
                text: model.reportReady ? "Go" : "Hold",
                bg: model.reportReady ? "#dcfce7" : "#fef3c7",
                fg: model.reportReady ? "#166534" : "#b45309",
              },
            },
          ].map((c) => (
            <div
              key={c.label}
              style={{
                borderRadius: 14,
                padding: "16px 14px",
                border: `1px solid ${C.border}`,
                background: "linear-gradient(165deg, #ffffff 0%, #f8fafc 100%)",
              }}
            >
              <div style={{ fontSize: 11, color: C.muted, fontWeight: 700 }}>{c.label}</div>
              <div style={{ fontSize: 28, fontWeight: 900, marginTop: 6, lineHeight: 1.1 }}>{c.value}</div>
              <div style={{ fontSize: 12, color: C.muted, marginTop: 6, lineHeight: 1.45 }}>{c.sub}</div>
              <div style={{ marginTop: 10 }}>
                <span
                  style={{
                    display: "inline-block",
                    padding: "3px 8px",
                    borderRadius: 999,
                    fontSize: 10,
                    fontWeight: 800,
                    background: c.badge.bg,
                    color: c.badge.fg,
                  }}
                >
                  {c.badge.text}
                </span>
              </div>
            </div>
          ))}
        </div>

        <div style={{ fontSize: 12, color: C.muted, marginBottom: 10 }}>
          Pre-export QA messages (all domains): <strong style={{ color: C.text }}>{reportQualityWarningCount}</strong>
        </div>

        <div
          style={{
            padding: "16px 18px",
            borderRadius: 14,
            border: `1px solid ${C.border}`,
            background: "#f8fafc",
            marginBottom: 22,
          }}
        >
          <div style={{ fontSize: 12, fontWeight: 800, color: C.text, marginBottom: 12 }}>Progress tracker</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {model.stages.map((s) => (
              <div key={s.key} style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={stageDot(s.done)}>{s.done ? "✓" : ""}</span>
                <span style={{ fontSize: 13, color: s.done ? C.text : C.muted, fontWeight: s.done ? 700 : 500 }}>
                  {s.label}
                </span>
              </div>
            ))}
          </div>
        </div>

        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fill, minmax(260px, 1fr))",
            gap: 16,
            marginBottom: 22,
          }}
        >
          <div style={{ border: `1px solid ${C.border}`, borderRadius: 14, padding: 16, background: C.surface }}>
            <div style={{ fontSize: 12, fontWeight: 800, color: C.text, marginBottom: 8 }}>Top 5 highest-risk domains</div>
            <ol style={{ margin: 0, paddingLeft: 18, fontSize: 13, color: C.text, lineHeight: 1.55 }}>
              {model.topRiskDomains.map((d) => (
                <li key={d.id}>
                  {d.label} — {d.score}/5 ({d.evidence}) · <span style={{ fontWeight: 700 }}>{d.priority}</span>
                </li>
              ))}
            </ol>
          </div>
          <div style={{ border: `1px solid ${C.border}`, borderRadius: 14, padding: 16, background: C.surface }}>
            <div style={{ fontSize: 12, fontWeight: 800, color: C.text, marginBottom: 8 }}>Evidence gaps (AI)</div>
            <div style={{ fontSize: 26, fontWeight: 900 }}>{model.evidenceGapDomains}</div>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>
              Domains with AI-flagged gaps · {model.evidenceGapBullets} total gap items
            </div>
          </div>
          <div style={{ border: `1px solid ${C.border}`, borderRadius: 14, padding: 16, background: C.surface }}>
            <div style={{ fontSize: 12, fontWeight: 800, color: C.text, marginBottom: 8 }}>Domains below benchmark</div>
            {model.belowBenchmark.length ? (
              <ul style={{ margin: 0, paddingLeft: 18, fontSize: 13, color: C.text, lineHeight: 1.55 }}>
                {model.belowBenchmark.map((name) => (
                  <li key={name}>{name}</li>
                ))}
              </ul>
            ) : (
              <div style={{ fontSize: 13, color: C.muted }}>None — all domains meet or exceed benchmark {BENCHMARK_SCORE}.</div>
            )}
          </div>
        </div>

        <div
          style={{
            border: `1px solid ${C.border}`,
            borderRadius: 14,
            padding: 16,
            background: C.surface,
            marginBottom: 22,
          }}
        >
          <div style={{ fontSize: 12, fontWeight: 800, color: C.text, marginBottom: 10 }}>Manager review</div>
          <p style={{ margin: "0 0 12px", fontSize: 13, color: C.muted, lineHeight: 1.55 }}>
            Report readiness requires manager review to be marked <strong>complete</strong> or <strong>not required</strong>.
            Current: <strong style={{ color: C.text }}>{managerReviewStatus}</strong> · AI run status:{" "}
            <strong style={{ color: C.text }}>{analysisStatus}</strong>
          </p>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {[
              { id: "pending", label: "Pending" },
              { id: "complete", label: "Review complete" },
              { id: "not_required", label: "Not required" },
            ].map((opt) => (
              <button
                key={opt.id}
                type="button"
                onClick={() => onManagerReviewStatus(opt.id)}
                style={{
                  padding: "8px 14px",
                  borderRadius: 10,
                  border:
                    managerReviewStatus === opt.id ? `2px solid ${C.accent}` : `1px solid ${C.border}`,
                  background: managerReviewStatus === opt.id ? C.accentSoft : "#fff",
                  fontWeight: 700,
                  fontSize: 12,
                  cursor: "pointer",
                  color: C.text,
                }}
              >
                {opt.label}
              </button>
            ))}
          </div>
        </div>

        <div style={{ border: `1px solid ${C.border}`, borderRadius: 14, padding: 16, background: C.surface }}>
          <div style={{ fontSize: 12, fontWeight: 800, color: C.text, marginBottom: 10 }}>Action list</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <button
              type="button"
              onClick={() => onNavigate("assessment")}
              style={{
                textAlign: "left",
                padding: "10px 12px",
                borderRadius: 10,
                border: `1px solid ${C.border}`,
                background: "#fff",
                fontWeight: 600,
                fontSize: 13,
                cursor: "pointer",
              }}
            >
              Complete missing notes — open Assessment workspace
            </button>
            <button
              type="button"
              onClick={() => onNavigate("reviewQueue")}
              style={{
                textAlign: "left",
                padding: "10px 12px",
                borderRadius: 10,
                border: `1px solid ${C.border}`,
                background: "#fff",
                fontWeight: 600,
                fontSize: 13,
                cursor: "pointer",
              }}
            >
              Review low-confidence AI suggestions — open Review Queue
            </button>
            <button
              type="button"
              onClick={() => onNavigate("reviewQueue")}
              style={{
                textAlign: "left",
                padding: "10px 12px",
                borderRadius: 10,
                border: `1px solid ${C.border}`,
                background: "#fff",
                fontWeight: 600,
                fontSize: 13,
                cursor: "pointer",
              }}
            >
              Resolve score mismatches — open Review Queue
            </button>
            <button
              type="button"
              disabled={analysisLoading}
              onClick={() => {
                onAcceptAllHighConfidence();
              }}
              style={{
                textAlign: "left",
                padding: "10px 12px",
                borderRadius: 10,
                border: `1px solid ${C.border}`,
                background: "#fff",
                fontWeight: 600,
                fontSize: 13,
                cursor: analysisLoading ? "not-allowed" : "pointer",
              }}
            >
              Approve high-confidence suggestions (bulk accept High confidence)
            </button>
            <button
              type="button"
              disabled={analysisLoading}
              onClick={() => {
                onAnalyzeEvidence();
              }}
              style={{
                textAlign: "left",
                padding: "10px 12px",
                borderRadius: 10,
                border: `1px solid ${C.border}`,
                background: "#fff",
                fontWeight: 600,
                fontSize: 13,
                cursor: analysisLoading ? "not-allowed" : "pointer",
              }}
            >
              Run / refresh AI evidence analysis
            </button>
            <button
              type="button"
              disabled={!model.reportReady || analysisLoading}
              title={!model.reportReady ? model.reportReadinessReasons.join(" ") : "Download PPTX"}
              onClick={() => onGenerateReport()}
              style={{
                textAlign: "left",
                padding: "10px 12px",
                borderRadius: 10,
                border: "none",
                background: model.reportReady ? "#1d4ed8" : "#cbd5e1",
                color: "#fff",
                fontWeight: 800,
                fontSize: 13,
                cursor: !model.reportReady || analysisLoading ? "not-allowed" : "pointer",
              }}
            >
              Generate final report (PPTX download when readiness allows)
            </button>
          </div>
        </div>

        {!model.reportReady && model.reportReadinessReasons.length > 0 && (
          <div
            style={{
              marginTop: 16,
              padding: "12px 14px",
              borderRadius: 12,
              background: "#fffbeb",
              border: "1px solid #fcd34d",
              fontSize: 13,
              color: "#92400e",
            }}
          >
            <strong>Report readiness blockers</strong>
            <ul style={{ margin: "8px 0 0", paddingLeft: 20 }}>
              {model.reportReadinessReasons.map((r, i) => (
                <li key={i} style={{ marginBottom: 4 }}>
                  {r}
                </li>
              ))}
            </ul>
          </div>
        )}
      </SectionCard>
    </>
  );
}

function ReviewQueuePageContent({
  profile,
  liveAssessment,
  aiSuggestionsByDomain,
  reviewFollowUpByDomain,
  onMarkFollowUp,
  onClearFollowUp,
  reviewerNotesByDomain,
  onReviewerNoteChange,
  acceptSuggestionForDomain,
  keepManualScoreForDomain,
}) {
  const sortedRows = useMemo(() => {
    const priOrder = { Critical: 0, High: 1, Medium: 2, Low: 3 };
    const built = liveAssessment.entries.map((e) => {
      const sig = computeDomainReviewSignals(e, profile, aiSuggestionsByDomain, reviewFollowUpByDomain);
      const displayReviewStatus = sig.follow ? `Follow-up (${sig.aiStatus})` : sig.aiStatus;
      return {
        e,
        ...sig,
        displayReviewStatus,
      };
    });
    built.sort((a, b) => {
      const d = priOrder[a.pri.label] - priOrder[b.pri.label];
      if (d !== 0) return d;
      if (a.priHot !== b.priHot) return a.priHot ? -1 : 1;
      if (a.bigDiff !== b.bigDiff) return a.bigDiff ? -1 : 1;
      if (a.lowConf !== b.lowConf) return a.lowConf ? -1 : 1;
      if (a.weakHigh !== b.weakHigh) return a.weakHigh ? -1 : 1;
      if (a.follow !== b.follow) return a.follow ? -1 : 1;
      return a.e.label.localeCompare(b.e.label);
    });
    return built;
  }, [liveAssessment, profile, aiSuggestionsByDomain, reviewFollowUpByDomain]);

  const summary = useMemo(() => {
    let domainsNeedingReview = 0;
    let highConfidenceAccepted = 0;
    let overrides = 0;
    let gapDomains = 0;
    let gapBullets = 0;
    sortedRows.forEach((r) => {
      if (r.needsReview) domainsNeedingReview += 1;
      if (r.aiStatus === "Accepted" && r.aiSug?.confidence === "High") highConfidenceAccepted += 1;
      if (r.aiStatus === "Overridden") overrides += 1;
      if (r.aiSug) {
        const merged = [
          ...(r.aiSug.evidenceGaps || []),
          ...(r.aiSug.missingEvidence || []),
        ].filter((x) => {
          const t = String(x || "").trim();
          return t && t !== "No direct evidence provided.";
        });
        if (merged.length) {
          gapDomains += 1;
          gapBullets += merged.length;
        }
      }
    });
    return { domainsNeedingReview, highConfidenceAccepted, overrides, gapDomains, gapBullets };
  }, [sortedRows]);

  const btnSm = {
    padding: "6px 10px",
    borderRadius: 8,
    fontSize: 11,
    fontWeight: 700,
    cursor: "pointer",
    border: "none",
  };

  return (
    <>
      <SectionCard
        title="Manager review queue"
        subtitle="One place to reconcile AI suggestions, manual scores, evidence gaps, and pre-export QA warnings before sign-off."
      >
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))",
            gap: 14,
            marginBottom: 18,
          }}
        >
          {[
            {
              label: "Domains needing review",
              value: summary.domainsNeedingReview,
              sub: "Follow-up, QA warnings, or material divergence",
              badge: { text: "Queue", bg: "#fff7ed", fg: "#9a3412" },
            },
            {
              label: "High-confidence accepted",
              value: summary.highConfidenceAccepted,
              sub: "AI score accepted with High confidence",
              badge: { text: "Cleared", bg: "#ecfdf5", fg: "#166534" },
            },
            {
              label: "Overrides",
              value: summary.overrides,
              sub: "Analyst kept manual vs. AI",
              badge: { text: "Manual", bg: "#f1f5f9", fg: "#475569" },
            },
            {
              label: "Evidence gaps (AI)",
              value: summary.gapDomains,
              sub: `${summary.gapBullets} gap items across domains`,
              badge: { text: "Gaps", bg: "#eff6ff", fg: "#1d4ed8" },
            },
          ].map((c) => (
            <div
              key={c.label}
              style={{
                borderRadius: 14,
                padding: "16px 14px",
                border: `1px solid ${C.border}`,
                background: "linear-gradient(165deg, #ffffff 0%, #f8fafc 100%)",
              }}
            >
              <div style={{ fontSize: 11, color: C.muted, fontWeight: 700 }}>{c.label}</div>
              <div style={{ fontSize: 30, fontWeight: 900, marginTop: 6 }}>{c.value}</div>
              <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{c.sub}</div>
              <div style={{ marginTop: 10 }}>
                <span
                  style={{
                    display: "inline-block",
                    padding: "3px 8px",
                    borderRadius: 999,
                    fontSize: 10,
                    fontWeight: 800,
                    background: c.badge.bg,
                    color: c.badge.fg,
                  }}
                >
                  {c.badge.text}
                </span>
              </div>
            </div>
          ))}
        </div>

        <div
          style={{
            marginBottom: 14,
            padding: "10px 12px",
            borderRadius: 10,
            background: "#f8fafc",
            border: `1px solid ${C.border}`,
            fontSize: 12,
            color: C.muted,
            lineHeight: 1.5,
          }}
        >
          <strong style={{ color: C.text }}>Highlights:</strong> shaded rows for Critical/High priority; amber for score
          gap ≥2, Low AI confidence, or Weak evidence with high score (4–5). Run{" "}
          <strong>Analyze Evidence &amp; Suggest Scores</strong> from the header to populate AI columns.
        </div>

        <div style={{ overflowX: "auto", borderRadius: 12, border: `1px solid ${C.border}` }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, minWidth: 920 }}>
            <thead>
              <tr style={{ background: "#f1f5f9", textAlign: "left", color: C.muted, fontSize: 11 }}>
                <th style={{ padding: 10, borderBottom: `1px solid ${C.border}` }}>Domain</th>
                <th style={{ padding: 10, borderBottom: `1px solid ${C.border}` }}>Manual</th>
                <th style={{ padding: 10, borderBottom: `1px solid ${C.border}` }}>AI</th>
                <th style={{ padding: 10, borderBottom: `1px solid ${C.border}` }}>Diff</th>
                <th style={{ padding: 10, borderBottom: `1px solid ${C.border}` }}>Confidence</th>
                <th style={{ padding: 10, borderBottom: `1px solid ${C.border}` }}>Review status</th>
                <th style={{ padding: 10, borderBottom: `1px solid ${C.border}` }}>QA #</th>
                <th style={{ padding: 10, borderBottom: `1px solid ${C.border}` }}>Priority</th>
                <th style={{ padding: 10, borderBottom: `1px solid ${C.border}` }}>Actions</th>
                <th style={{ padding: 10, borderBottom: `1px solid ${C.border}` }}>Reviewer note</th>
              </tr>
            </thead>
            <tbody>
              {sortedRows.map((r) => {
                const e = r.e;
                const rowBg = r.priHot
                  ? "rgba(254,202,202,0.45)"
                  : r.bigDiff || r.lowConf || r.weakHigh
                    ? "rgba(254,243,199,0.55)"
                    : "#ffffff";
                const leftBd =
                  r.pri.label === "Critical"
                    ? "#991b1b"
                    : r.pri.label === "High"
                      ? "#ea580c"
                      : "transparent";
                return (
                  <tr
                    key={e.id}
                    style={{
                      background: rowBg,
                      borderLeft: leftBd !== "transparent" ? `4px solid ${leftBd}` : "4px solid transparent",
                    }}
                  >
                    <td style={{ padding: 10, borderBottom: `1px solid ${C.border}`, fontWeight: 700, color: C.text }}>
                      {e.label}
                    </td>
                    <td style={{ padding: 10, borderBottom: `1px solid ${C.border}` }}>{e.score}</td>
                    <td style={{ padding: 10, borderBottom: `1px solid ${C.border}` }}>
                      {r.suggested !== null ? r.suggested : "—"}
                    </td>
                    <td
                      style={{
                        padding: 10,
                        borderBottom: `1px solid ${C.border}`,
                        fontWeight: r.bigDiff ? 800 : 600,
                        color: r.bigDiff ? "#b45309" : C.text,
                      }}
                    >
                      {r.diff === null ? "—" : r.diff > 0 ? `+${r.diff}` : String(r.diff)}
                    </td>
                    <td style={{ padding: 10, borderBottom: `1px solid ${C.border}` }}>
                      {r.aiSug?.confidence ? (
                        <span
                          style={{
                            padding: "2px 8px",
                            borderRadius: 999,
                            fontWeight: 800,
                            fontSize: 10,
                            background:
                              r.aiSug.confidence === "High"
                                ? "#d1fae5"
                                : r.aiSug.confidence === "Low"
                                  ? "#fee2e2"
                                  : "#fef3c7",
                            color:
                              r.aiSug.confidence === "High"
                                ? "#065f46"
                                : r.aiSug.confidence === "Low"
                                  ? "#991b1b"
                                  : "#92400e",
                          }}
                        >
                          {r.aiSug.confidence}
                        </span>
                      ) : (
                        "—"
                      )}
                    </td>
                    <td style={{ padding: 10, borderBottom: `1px solid ${C.border}` }}>{r.displayReviewStatus}</td>
                    <td style={{ padding: 10, borderBottom: `1px solid ${C.border}`, fontVariantNumeric: "tabular-nums" }}>
                      {r.qaCount}
                    </td>
                    <td style={{ padding: 10, borderBottom: `1px solid ${C.border}`, fontWeight: 700 }}>
                      <span style={{ color: r.pri.color }}>{r.pri.label}</span>
                    </td>
                    <td style={{ padding: 10, borderBottom: `1px solid ${C.border}`, verticalAlign: "top" }}>
                      <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                        <button
                          type="button"
                          disabled={!r.aiSug}
                          onClick={() => acceptSuggestionForDomain(e.id)}
                          style={{
                            ...btnSm,
                            background: r.aiSug ? "#1d4ed8" : "#cbd5e1",
                            color: "#fff",
                          }}
                        >
                          Accept AI score
                        </button>
                        <button
                          type="button"
                          disabled={!r.aiSug}
                          onClick={() => keepManualScoreForDomain(e.id)}
                          style={{
                            ...btnSm,
                            background: "#fff",
                            color: "#334155",
                            border: "1px solid #94a3b8",
                          }}
                        >
                          Keep manual score
                        </button>
                        {!r.follow ? (
                          <button
                            type="button"
                            onClick={() => onMarkFollowUp(e.id)}
                            style={{ ...btnSm, background: "#f59e0b", color: "#fff" }}
                          >
                            Mark for follow-up
                          </button>
                        ) : (
                          <button
                            type="button"
                            onClick={() => onClearFollowUp(e.id)}
                            style={{ ...btnSm, background: "#e2e8f0", color: "#334155" }}
                          >
                            Clear follow-up
                          </button>
                        )}
                      </div>
                    </td>
                    <td style={{ padding: 10, borderBottom: `1px solid ${C.border}`, verticalAlign: "top", minWidth: 200 }}>
                      <textarea
                        value={reviewerNotesByDomain[e.id] ?? ""}
                        onChange={(ev) => onReviewerNoteChange(e.id, ev.target.value)}
                        placeholder="Reviewer comment"
                        style={{ ...selectStyle, minHeight: 64, width: "100%", fontSize: 12 }}
                      />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </SectionCard>
    </>
  );
}

function ResultsPageContent({
  liveAssessment,
  hasGenerated,
  maturityBadgeColor,
  aiNarratives,
  narrativeLoading,
  generateTPRMReport,
  reportStatus,
  reportQualityWarnings,
  clientDraft,
}) {
  return (
    <>
      {!hasGenerated && (
        <div
          style={{
            marginBottom: 22,
            padding: "16px 20px",
            borderRadius: 12,
            background: "linear-gradient(90deg, #fffbeb 0%, #fef3c7 100%)",
            border: "1px solid #fcd34d",
            fontSize: 13,
            color: "#92400e",
            boxShadow: "0 2px 12px rgba(245,158,11,0.12)",
          }}
        >
          <strong>Draft view.</strong> Select <strong>Generate assessment</strong> in the header to finalize narratives for
          management reporting. All metrics reflect your current workspace.
        </div>
      )}

      <div
        style={{
          marginBottom: 22,
          padding: "16px 20px",
          borderRadius: 14,
          background: reportQualityWarnings.length
            ? "linear-gradient(90deg, #fff7ed 0%, #ffedd5 100%)"
            : "linear-gradient(90deg, #ecfdf5 0%, #f0fdf4 100%)",
          border: `1px solid ${reportQualityWarnings.length ? "#fdba74" : "#86efac"}`,
          boxShadow: "0 2px 12px rgba(15,23,42,0.06)",
        }}
      >
        <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: "0.08em", color: C.muted, marginBottom: 8 }}>
          Pre-report quality check
        </div>
        {reportQualityWarnings.length === 0 ? (
          <p style={{ margin: 0, fontSize: 13, color: "#166534", lineHeight: 1.5 }}>
            No pre-export issues flagged — review still recommended, but low-score notes, evidence vs. score, AI
            confidence, and manual vs. AI alignment look acceptable for export.
          </p>
        ) : (
          <ul style={{ margin: 0, paddingLeft: 20, fontSize: 13, color: "#9a3412", lineHeight: 1.55 }}>
            {reportQualityWarnings.map((w, i) => (
              <li key={i} style={{ marginBottom: 6 }}>
                {w}
              </li>
            ))}
          </ul>
        )}
      </div>

      <div
        style={{
          marginBottom: 22,
          display: "flex",
          flexWrap: "wrap",
          alignItems: "center",
          justifyContent: "space-between",
          gap: 14,
          padding: "16px 20px",
          borderRadius: 14,
          background: C.surface,
          border: `1px solid ${C.border}`,
          boxShadow: "0 4px 18px rgba(15,23,42,0.06)",
        }}
      >
        <div style={{ flex: "1 1 220px", minWidth: 0 }}>
          <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: "0.08em", color: C.muted, marginBottom: 4 }}>
            Executive export
          </div>
          <p style={{ margin: 0, fontSize: 13, color: C.text, lineHeight: 1.55 }}>
            Download a PowerPoint report (.pptx) with title, executive summary, maturity table, gaps, and roadmap. Review
            the <strong>pre-report quality check</strong> above before exporting.
          </p>
          <p
            style={{
              margin: "10px 0 0",
              fontSize: 12,
              fontWeight: 600,
              color: reportStatus.startsWith("Report failed") || reportStatus.includes("failed") ? "#b91c1c" : C.accent,
              lineHeight: 1.45,
              wordBreak: "break-word",
            }}
          >
            {reportStatus}
          </p>
        </div>
        <button
          type="button"
          onClick={generateTPRMReport}
          style={{
            flexShrink: 0,
            padding: "12px 22px",
            borderRadius: 10,
            border: "none",
            background: `linear-gradient(180deg, #0f172a 0%, #1e3a5f 100%)`,
            color: "#fff",
            fontWeight: 800,
            fontSize: 13,
            cursor: "pointer",
            boxShadow: "0 6px 18px rgba(15,23,42,0.22)",
          }}
        >
          Download Report
        </button>
      </div>

      {hasGenerated && clientDraft && (
        <ClientAssessmentDraftSection draft={clientDraft} narrativeLoading={narrativeLoading} />
      )}

      {/* Executive summary — narrative after Generate */}
      <section
        style={{
          borderRadius: 16,
          padding: "26px 28px",
          marginBottom: 22,
          background: "linear-gradient(135deg, #0c1324 0%, #1e3a5f 55%, #1e40af 100%)",
          color: "#fff",
          boxShadow: "0 12px 40px rgba(15,23,42,0.25)",
        }}
      >
        <div
          style={{
            display: "flex",
            flexWrap: "wrap",
            alignItems: "center",
            justifyContent: "space-between",
            gap: 10,
          }}
        >
          <div style={{ fontSize: 11, opacity: 0.8, textTransform: "uppercase", letterSpacing: "0.12em", fontWeight: 700 }}>
            Executive summary
          </div>
          {hasGenerated && aiNarratives?.source && (
            <span
              style={{
                fontSize: 10,
                fontWeight: 800,
                letterSpacing: "0.06em",
                textTransform: "uppercase",
                padding: "4px 10px",
                borderRadius: 999,
                background: aiNarratives.source === "openai" ? "rgba(34,197,94,0.25)" : "rgba(251,191,36,0.25)",
                border: `1px solid ${aiNarratives.source === "openai" ? "rgba(34,197,94,0.5)" : "rgba(251,191,36,0.45)"}`,
              }}
            >
              {aiNarratives.source === "openai" ? "AI narrative" : "Rule-based narrative"}
            </span>
          )}
        </div>
        {hasGenerated && narrativeLoading ? (
          <p style={{ margin: "14px 0 0", fontSize: 15, lineHeight: 1.7, opacity: 0.85, maxWidth: 900 }}>
            Generating consulting narrative…
          </p>
        ) : (
          <p style={{ margin: "14px 0 0", fontSize: 15, lineHeight: 1.7, opacity: 0.96, maxWidth: 900 }}>
            {hasGenerated && aiNarratives?.executiveSummary
              ? aiNarratives.executiveSummary
              : liveAssessment.executiveSummary}
          </p>
        )}
      </section>

      {hasGenerated && (narrativeLoading || aiNarratives) && (
        <SectionCard
          title="Risk narrative"
          subtitle="Exposure, control weaknesses, and implications — suitable for risk committee discussion."
        >
          {narrativeLoading ? (
            <p style={{ margin: 0, fontSize: 14, color: C.muted }}>Generating risk perspective…</p>
          ) : (
            <p style={{ margin: 0, fontSize: 14, lineHeight: 1.75, color: C.text }}>{aiNarratives?.riskNarrative}</p>
          )}
        </SectionCard>
      )}

      {hasGenerated && (narrativeLoading || aiNarratives) && (
        <SectionCard title="Recommendations" subtitle="Prioritized, actionable moves aligned to this assessment.">
          {narrativeLoading ? (
            <p style={{ margin: 0, fontSize: 14, color: C.muted }}>Generating recommendations…</p>
          ) : (
            <ul style={{ margin: 0, paddingLeft: 22, fontSize: 14, lineHeight: 1.7, color: C.text }}>
              {splitRecommendationLines(aiNarratives.recommendationsText).map((line, i) => (
                <li key={i} style={{ marginBottom: 12 }}>
                  {line}
                </li>
              ))}
            </ul>
          )}
        </SectionCard>
      )}

      <ExecutiveScorecard liveAssessment={liveAssessment} maturityBadgeColor={maturityBadgeColor} />

      <SectionCard title="Top 5 gaps" subtitle="Ranked by score and evidence weighting — same ordering as the engine.">
        <ol style={{ margin: 0, paddingLeft: 22, fontSize: 14, lineHeight: 1.6 }}>
          {liveAssessment.topGapsAll.slice(0, 5).map((g, i) => (
            <li key={i} style={{ marginBottom: 14 }}>
              <strong>{g.domain}</strong>
              <span style={{ color: C.muted }}>
                {" "}
                · {g.score}/5 · {g.evidence}
              </span>
              <div style={{ color: C.muted, fontSize: 13, marginTop: 4 }}>{g.rationale}</div>
            </li>
          ))}
        </ol>
      </SectionCard>

      {/* Maturity Matrix */}
      <SectionCard
        title="Maturity Matrix"
        subtitle="Per-domain score, maturity band (Low / Moderate / Advanced), evidence, and priority rating."
      >
        <div
          style={{
            borderRadius: 14,
            border: `1px solid ${C.border}`,
            overflow: "hidden",
            background: C.surface,
            boxShadow: "0 4px 20px rgba(15,23,42,0.06)",
          }}
        >
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, minWidth: 720 }}>
              <thead>
                <tr style={{ background: "#f1f5f9", borderBottom: `1px solid ${C.border}` }}>
                  {["Domain", "Score", "Maturity level", "Evidence strength", "Priority"].map((h) => (
                    <th
                      key={h}
                      style={{
                        textAlign: "left",
                        padding: "14px 16px",
                        fontSize: 10,
                        fontWeight: 800,
                        letterSpacing: "0.08em",
                        textTransform: "uppercase",
                        color: C.muted,
                      }}
                    >
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {liveAssessment.entries.map((e, idx) => {
                  const mat = matrixMaturityBand(e.score);
                  const pri = matrixPriority(e.score, e.evidence);
                  return (
                    <tr
                      key={e.id}
                      style={{
                        background: idx % 2 === 0 ? "#fff" : "#fafafa",
                        borderBottom: `1px solid ${C.border}`,
                      }}
                    >
                      <td style={{ padding: "12px 16px", fontWeight: 700 }}>{e.label}</td>
                      <td style={{ padding: "12px 16px", fontVariantNumeric: "tabular-nums", fontWeight: 800 }}>{e.score}</td>
                      <td style={{ padding: "12px 16px" }}>
                        <PillBadge bg={mat.badgeBg} fg={mat.badgeFg} bd={mat.badgeBd}>
                          {mat.label}
                        </PillBadge>
                      </td>
                      <td style={{ padding: "12px 16px", color: C.text }}>{e.evidence}</td>
                      <td style={{ padding: "12px 16px", fontWeight: pri.fontWeight, color: pri.color }}>{pri.label}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </SectionCard>

      {/* Roadmap summary */}
      <SectionCard title="Roadmap summary" subtitle="30 / 60 / 90 execution horizons from the scoring engine.">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 16 }}>
          {[
            ["30 days", liveAssessment.roadmap.d30],
            ["60 days", liveAssessment.roadmap.d60],
            ["90 days", liveAssessment.roadmap.d90],
          ].map(([label, items]) => (
            <div
              key={label}
              style={{
                borderRadius: 14,
                padding: 18,
                border: `1px solid ${C.border}`,
                background: "linear-gradient(180deg,#ffffff 0%, #f8fafc 100%)",
                boxShadow: "0 4px 18px rgba(15,23,42,0.06)",
              }}
            >
              <div style={{ fontSize: 11, fontWeight: 900, color: C.accent, letterSpacing: "0.08em", marginBottom: 12 }}>
                {label}
              </div>
              <ul style={{ margin: 0, paddingLeft: 18, fontSize: 13, color: C.text, lineHeight: 1.55 }}>
                {items.map((item, i) => (
                  <li key={i} style={{ marginBottom: 8 }}>
                    {item}
                  </li>
                ))}
              </ul>
            </div>
          ))}
        </div>
      </SectionCard>
    </>
  );
}

function VisualAnalyticsPage({ liveAssessment }) {
  const entries = liveAssessment.entries;
  const n = entries.length;
  const cx = 240;
  const cy = 250;
  const R = 128;
  const clientPoly = polygonPointsFromEntries(entries, cx, cy, R, (e) => e.score);
  const benchPoly = polygonPointsFromEntries(entries, cx, cy, R, () => BENCHMARK_SCORE);
  const shortLabel = (s) => (s.length > 16 ? `${s.slice(0, 14)}…` : s);

  const dist = {
    low: entries.filter((e) => e.score <= 2).length,
    mod: entries.filter((e) => e.score === 3).length,
    adv: entries.filter((e) => e.score >= 4).length,
  };
  const distMax = Math.max(dist.low, dist.mod, dist.adv, 1);

  const insights = buildExecutiveInsights(entries);

  return (
    <>
      <section
        style={{
          marginBottom: 22,
          padding: "20px 24px",
          borderRadius: 16,
          background: "linear-gradient(120deg, #1e3a5f 0%, #0f172a 100%)",
          color: "#e2e8f0",
          boxShadow: "0 12px 40px rgba(15,23,42,0.2)",
        }}
      >
        <div style={{ fontSize: 11, fontWeight: 800, letterSpacing: "0.12em", opacity: 0.8 }}>VISUAL ANALYTICS</div>
        <p style={{ margin: "10px 0 0", fontSize: 15, lineHeight: 1.65, maxWidth: 800 }}>
          Portfolio diagnostic views: radar overlay, domain bars, risk–evidence scatter, and distribution. No external
          chart libraries—SVG and layout-only rendering.
        </p>
      </section>

      {/* Radar */}
      <SectionCard
        title="Maturity radar"
        subtitle="Client scores (blue) vs benchmark ring at 3.5 (slate). Scale 1–5 on radial axis."
      >
        <div style={{ display: "flex", justifyContent: "center", padding: "8px 0 24px" }}>
          <svg width="500" height="500" viewBox="0 0 500 500" style={{ maxWidth: "100%", height: "auto" }}>
            <rect width="500" height="500" fill="#f8fafc" rx="12" />
            {[1, 2, 3, 4, 5].map((k) => (
              <circle
                key={k}
                cx={cx}
                cy={cy}
                r={(k / 5) * R}
                fill="none"
                stroke="#e2e8f0"
                strokeWidth="1"
              />
            ))}
            {entries.map((e, i) => {
              const a = (2 * Math.PI * i) / n - Math.PI / 2;
              const x2 = cx + R * Math.cos(a);
              const y2 = cy + R * Math.sin(a);
              return (
                <line key={e.id} x1={cx} y1={cy} x2={x2} y2={y2} stroke="#e2e8f0" strokeWidth="1" />
              );
            })}
            <polygon
              points={benchPoly}
              fill="rgba(148, 163, 184, 0.22)"
              stroke="#94a3b8"
              strokeWidth="2"
              strokeDasharray="6 4"
            />
            <polygon
              points={clientPoly}
              fill="rgba(37, 99, 235, 0.2)"
              stroke="#2563eb"
              strokeWidth="2.5"
            />
            {entries.map((e, i) => {
              const { lx, ly } = radarPoint(cx, cy, R, 5, i, n);
              return (
                <text
                  key={e.id}
                  x={lx}
                  y={ly}
                  textAnchor="middle"
                  dominantBaseline="middle"
                  fontSize="9"
                  fill="#475569"
                  fontWeight="600"
                >
                  {shortLabel(e.label)}
                </text>
              );
            })}
            <text x={cx} y={36} textAnchor="middle" fontSize="12" fill="#64748b" fontWeight="700">
              1 – 5 maturity scale · rings at each integer
            </text>
          </svg>
        </div>
      </SectionCard>

      {/* Domain bars + benchmark tick */}
      <SectionCard
        title="Domain maturity vs benchmark"
        subtitle="Bar = client score; tick marks sector benchmark 3.5. Gap shown numerically."
      >
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {entries.map((e) => {
            const mat = matrixMaturityBand(e.score);
            const gap = e.score - BENCHMARK_SCORE;
            const benchPct = (BENCHMARK_SCORE / 5) * 100;
            const fillPct = (e.score / 5) * 100;
            return (
              <div key={e.id} style={{ display: "grid", gridTemplateColumns: "minmax(160px, 1fr) 2fr 80px", gap: 12, alignItems: "center" }}>
                <div style={{ fontSize: 12, fontWeight: 700 }}>{e.label}</div>
                <div style={{ position: "relative" }}>
                  <div style={{ height: 18, borderRadius: 8, background: "#e2e8f0", overflow: "visible", position: "relative" }}>
                    <div
                      style={{
                        width: `${fillPct}%`,
                        height: "100%",
                        borderRadius: 8,
                        background: `linear-gradient(90deg, ${mat.badgeFg}99, ${mat.badgeFg})`,
                      }}
                    />
                    <div
                      style={{
                        position: "absolute",
                        left: `${benchPct}%`,
                        top: -4,
                        width: 3,
                        height: 26,
                        background: "#64748b",
                        borderRadius: 2,
                        transform: "translateX(-50%)",
                        zIndex: 2,
                      }}
                      title="Benchmark 3.5"
                    />
                  </div>
                </div>
                <div style={{ fontSize: 12, fontVariantNumeric: "tabular-nums", color: gap < 0 ? "#b91c1c" : "#15803d", fontWeight: 700 }}>
                  {gap >= 0 ? "+" : ""}
                  {gap.toFixed(1)} gap
                </div>
              </div>
            );
          })}
        </div>
        <div style={{ marginTop: 16, fontSize: 12, color: C.muted }}>
          Gray vertical tick = benchmark ({BENCHMARK_SCORE}). Bar color reflects Low / Moderate / Advanced band.
        </div>
      </SectionCard>

      {/* Bubble / scatter */}
      <SectionCard
        title="Risk priority map"
        subtitle="X = domain score (1–5). Y = evidence (Weak=1, Moderate=3, Strong=5). Bubble size: lower scores = larger. Jitter and ring spread separate points; full detail in tooltips."
      >
        <RiskPriorityMapSvg entries={entries} />
      </SectionCard>

      {/* Distribution */}
      <SectionCard title="Maturity distribution" subtitle="Count of domains by matrix band (Low / Moderate / Advanced).">
        <div style={{ display: "flex", gap: 28, alignItems: "flex-end", justifyContent: "center", padding: "20px 12px 8px" }}>
          {[
            { key: "Low", count: dist.low, color: "#fecaca", fg: "#991b1b" },
            { key: "Moderate", count: dist.mod, color: "#fef08a", fg: "#b45309" },
            { key: "Advanced", count: dist.adv, color: "#86efac", fg: "#15803d" },
          ].map((b) => (
            <div key={b.key} style={{ textAlign: "center", width: 120 }}>
              <div style={{ fontSize: 28, fontWeight: 900, color: b.fg }}>{b.count}</div>
              <div style={{ height: 140, display: "flex", alignItems: "flex-end", justifyContent: "center", marginTop: 8 }}>
                <div
                  style={{
                    width: 56,
                    height: `${(b.count / distMax) * 120}px`,
                    minHeight: b.count === 0 ? 4 : 12,
                    borderRadius: "12px 12px 4px 4px",
                    background: `linear-gradient(180deg, ${b.color}, ${b.fg}44)`,
                    border: `1px solid ${b.fg}44`,
                    boxShadow: "0 8px 20px rgba(15,23,42,0.08)",
                  }}
                />
              </div>
              <div style={{ marginTop: 10, fontSize: 12, fontWeight: 800, color: C.text }}>{b.key}</div>
            </div>
          ))}
        </div>
      </SectionCard>

      {/* Insights */}
      <SectionCard title='What this means' subtitle="Interpretive signals driven by your scored inputs—same thresholds as matrix and priorities.">
        <ul style={{ margin: 0, paddingLeft: 22, fontSize: 14, lineHeight: 1.65, color: C.text }}>
          {insights.map((line, i) => (
            <li key={i} style={{ marginBottom: 14 }}>
              {line}
            </li>
          ))}
        </ul>
      </SectionCard>
    </>
  );
}

function LibraryPageContent({ libraryRefreshTrigger = 0 }) {
  const [memTab, setMemTab] = useState("findings");
  const [search, setSearch] = useState("");
  const [domainFilter, setDomainFilter] = useState("all");
  const [findingsLib, setFindingsLib] = useState([]);
  const [recsLib, setRecsLib] = useState([]);
  const [narsLib, setNarsLib] = useState([]);

  useEffect(() => {
    setFindingsLib([...readJsonLibrary(LS_FINDINGS_LIBRARY)].reverse());
    setRecsLib([...readJsonLibrary(LS_RECOMMENDATIONS_LIBRARY)].reverse());
    setNarsLib([...readJsonLibrary(LS_NARRATIVES_LIBRARY)].reverse());
  }, [libraryRefreshTrigger]);

  const filterMemoryList = (list) => {
    let out = list;
    if (domainFilter !== "all") {
      out = out.filter((it) => it.domainId === domainFilter);
    }
    const q = search.trim().toLowerCase();
    if (q) {
      out = out.filter((it) =>
        [it.domain, it.findingText, it.recommendationText, it.evidencePatterns, it.maturityLevel]
          .filter(Boolean)
          .join(" ")
          .toLowerCase()
          .includes(q)
      );
    }
    return out;
  };

  const tabList =
    memTab === "findings"
      ? filterMemoryList(findingsLib)
      : memTab === "recommendations"
        ? filterMemoryList(recsLib)
        : filterMemoryList(narsLib);

  return (
    <>
      <section
        style={{
          marginBottom: 24,
          padding: "20px 22px",
          borderRadius: 14,
          background: "linear-gradient(115deg, #ecfdf5 0%, #ffffff 55%)",
          border: `1px solid ${C.border}`,
          boxShadow: "0 4px 20px rgba(15,23,42,0.06)",
        }}
      >
        <p style={{ margin: 0, fontSize: 14, lineHeight: 1.65, color: C.text }}>
          <strong>Assessment memory.</strong> Each completed run (Generate assessment) appends domain findings, recommendations,
          and narrative fragments to your browser library.{" "}
          <strong>This system improves over time as more assessments are completed</strong> — similar maturity and
          evidence patterns are surfaced automatically when generating executive summaries, risk narratives, and
          recommendations.
        </p>
      </section>

      <SectionCard
        title="Reusable memory (this browser)"
        subtitle="Search and filter items captured from past assessment completions. Keys: findingsLibrary, recommendationsLibrary, narrativesLibrary."
      >
        <div style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 14, alignItems: "center" }}>
          {[
            { id: "findings", label: "Findings" },
            { id: "recommendations", label: "Recommendations" },
            { id: "narratives", label: "Narratives" },
          ].map((t) => (
            <button
              key={t.id}
              type="button"
              onClick={() => setMemTab(t.id)}
              style={{
                padding: "8px 14px",
                borderRadius: 10,
                border: memTab === t.id ? `2px solid ${C.accent}` : `1px solid ${C.border}`,
                background: memTab === t.id ? C.accentSoft : "#fff",
                fontWeight: 700,
                fontSize: 12,
                cursor: "pointer",
                color: C.text,
              }}
            >
              {t.label}{" "}
              <span style={{ color: C.muted, fontWeight: 600 }}>
                (
                {t.id === "findings" ? findingsLib.length : t.id === "recommendations" ? recsLib.length : narsLib.length})
              </span>
            </button>
          ))}
        </div>
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))",
            gap: 12,
            marginBottom: 14,
          }}
        >
          <Field label="Search">
            <input
              type="search"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Search library"
              style={selectStyle}
            />
          </Field>
          <Field label="Domain filter">
            <select
              value={domainFilter}
              onChange={(e) => setDomainFilter(e.target.value)}
              style={selectStyle}
            >
              <option value="all">All domains</option>
              {DOMAINS.map((d) => (
                <option key={d.id} value={d.id}>
                  {d.label}
                </option>
              ))}
              <option value="portfolio">Portfolio / executive</option>
            </select>
          </Field>
        </div>
        {tabList.length === 0 ? (
          <p style={{ margin: 0, fontSize: 13, color: C.muted }}>
            No items yet — complete an assessment (Generate assessment) to populate memory, then return here.
          </p>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 12, maxHeight: 520, overflowY: "auto" }}>
            {tabList.map((it) => (
              <div
                key={it.id}
                style={{
                  borderRadius: 12,
                  padding: 14,
                  border: `1px solid ${C.border}`,
                  background: "#fafafa",
                }}
              >
                <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center", marginBottom: 8 }}>
                  <span style={{ fontSize: 11, fontWeight: 800, color: C.accent }}>{it.domain}</span>
                  <span
                    style={{
                      fontSize: 10,
                      fontWeight: 800,
                      padding: "2px 8px",
                      borderRadius: 999,
                      background: "#e0e7ff",
                      color: "#3730a3",
                    }}
                  >
                    {it.maturityLevel}
                  </span>
                  <span
                    style={{
                      fontSize: 10,
                      fontWeight: 800,
                      padding: "2px 8px",
                      borderRadius: 999,
                      background: "#f1f5f9",
                      color: "#475569",
                    }}
                  >
                    {it.evidenceStrength}
                  </span>
                  <span style={{ fontSize: 10, color: C.muted, marginLeft: "auto" }}>
                    {it.savedAt ? String(it.savedAt).slice(0, 10) : ""}
                  </span>
                </div>
                {memTab === "findings" && it.findingText ? (
                  <p style={{ margin: "0 0 8px", fontSize: 13, color: C.text, lineHeight: 1.55, whiteSpace: "pre-wrap" }}>
                    <strong>Finding.</strong> {it.findingText}
                  </p>
                ) : null}
                {memTab === "findings" && it.recommendationText ? (
                  <p style={{ margin: "0 0 8px", fontSize: 13, color: C.text, lineHeight: 1.55, whiteSpace: "pre-wrap" }}>
                    <strong>Recommendation.</strong> {it.recommendationText}
                  </p>
                ) : null}
                {memTab === "recommendations" && it.recommendationText ? (
                  <p style={{ margin: 0, fontSize: 13, color: C.text, lineHeight: 1.55, whiteSpace: "pre-wrap" }}>
                    {it.recommendationText}
                  </p>
                ) : null}
                {memTab === "narratives" && it.findingText ? (
                  <p style={{ margin: 0, fontSize: 13, color: C.text, lineHeight: 1.55, whiteSpace: "pre-wrap" }}>
                    {it.findingText}
                  </p>
                ) : null}
                {it.evidencePatterns ? (
                  <p style={{ margin: "8px 0 0", fontSize: 11, color: C.muted }}>
                    <strong>Evidence patterns:</strong> {it.evidencePatterns}
                  </p>
                ) : null}
              </div>
            ))}
          </div>
        )}
      </SectionCard>

      <SectionCard title="Common gaps (reference patterns)" subtitle="Signal phrases frequently echoed in QA, audits, and exams.">
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          {[
            {
              title: "Golden record fragmentation",
              body: "Vendor IDs diverge between procurement, legal, and risk systems — residual concentration is opaque.",
              meta: "Finding pattern",
            },
            {
              title: "Attestation without artifact",
              body: "Monitoring ‘green’ status lacks contemporaneous evidence packs suitable for challenge.",
              meta: "Evidence gap",
            },
            {
              title: "Delegated oversight drift",
              body: "Business-led TP ownership without independent risk challenge on Tier-1 exits.",
              meta: "Governance",
            },
          ].map((c) => (
            <div
              key={c.title}
              style={{
                borderRadius: 12,
                padding: 16,
                border: `1px solid ${C.border}`,
                background: "#fafafa",
              }}
            >
              <div style={{ fontSize: 10, fontWeight: 800, color: C.accent, letterSpacing: "0.06em", marginBottom: 6 }}>{c.meta}</div>
              <strong style={{ fontSize: 14 }}>{c.title}</strong>
              <p style={{ margin: "8px 0 0", fontSize: 13, color: C.muted, lineHeight: 1.55 }}>{c.body}</p>
            </div>
          ))}
        </div>
      </SectionCard>

      <SectionCard title="Standard recommendation language" subtitle="Committee-ready phrasing — tailor jurisdiction and materiality.">
        <div style={{ display: "grid", gap: 12 }}>
          {[
            "Establish a Tier-0 governance lane with named sponsors, time-bound control uplift targets, and quarterly attestations to the risk committee.",
            "Embed consolidated concentration metrics (single-counterparty, geography, fourth-party) into monthly MI with breach narratives.",
            "Where residual risk remains elevated, document compensating controls with expiry and independent validation hooks.",
          ].map((t, i) => (
            <div
              key={i}
              style={{
                padding: 16,
                borderRadius: 12,
                borderLeft: `4px solid ${C.accent}`,
                background: C.surface,
                boxShadow: "0 2px 12px rgba(15,23,42,0.05)",
                fontSize: 13,
                lineHeight: 1.6,
              }}
            >
              {t}
            </div>
          ))}
        </div>
      </SectionCard>

      <SectionCard title="Remediation plays" subtitle="Reusable work packages for digital risk delivery.">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 14 }}>
          {[
            {
              title: "Critical vendor diligence reset",
              body: "30/60/90 re-baselining: data room, control testing, contract alignment, and exit readiness for Tier-0.",
              tag: "Play",
            },
            {
              title: "Monitoring & MI uplift",
              body: "KRI library, owner matrix, automated triggers, and board-ready MI one-pager template.",
              tag: "Play",
            },
            {
              title: "Regulatory mapping sprint",
              body: "Map supervision expectations to control inventory; produce evidence index and gap register.",
              tag: "Play",
            },
          ].map((p) => (
            <div
              key={p.title}
              style={{
                borderRadius: 14,
                padding: 18,
                background: "linear-gradient(180deg,#fff,#f8fafc)",
                border: `1px solid ${C.border}`,
                minHeight: 160,
                display: "flex",
                flexDirection: "column",
              }}
            >
              <span
                style={{
                  alignSelf: "flex-start",
                  fontSize: 9,
                  fontWeight: 900,
                  letterSpacing: "0.1em",
                  color: C.accent,
                  background: C.accentSoft,
                  padding: "4px 8px",
                  borderRadius: 6,
                }}
              >
                {p.tag}
              </span>
              <strong style={{ fontSize: 15, marginTop: 10 }}>{p.title}</strong>
              <p style={{ margin: "10px 0 0", fontSize: 13, color: C.muted, lineHeight: 1.55, flex: 1 }}>{p.body}</p>
            </div>
          ))}
        </div>
      </SectionCard>

      <SectionCard title="Evidence examples" subtitle="What “strong” evidence often includes in practice.">
        <ul style={{ margin: 0, paddingLeft: 20, fontSize: 14, lineHeight: 1.65, color: C.text }}>
          <li style={{ marginBottom: 10 }}>Diligence pack with third-line review sign-off and issue closure evidence.</li>
          <li style={{ marginBottom: 10 }}>Control test results mapped to contract clauses and compensating control register.</li>
          <li style={{ marginBottom: 10 }}>Monitoring logs / tickets showing threshold breach, escalation, and management response.</li>
        </ul>
      </SectionCard>

      <SectionCard title="Assessment questions" subtitle="Use in workshops or self-assessments to probe depth.">
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {[
            "How is criticality tiering challenged when business disputes materiality classification?",
            "What happens within 24 hours of a KRI breach for a Tier-0 counterparty?",
            "Can you produce a full fourth-party map for your top 5 vendors in 5 business days?",
            "Where is the single source of truth for contractual vs operational resilience commitments?",
          ].map((q, i) => (
            <div
              key={i}
              style={{
                padding: "14px 16px",
                borderRadius: 10,
                background: "#f8fafc",
                border: `1px solid ${C.border}`,
                fontSize: 13,
                fontWeight: 500,
              }}
            >
              {q}
            </div>
          ))}
        </div>
      </SectionCard>
    </>
  );
}

function RoadmapPageContent({ liveAssessment }) {
  const sorted = [...liveAssessment.entries].sort((a, b) =>
    a.score !== b.score ? a.score - b.score : a.weighted - b.weighted
  );
  const focusDomains = sorted.slice(0, 3).map((e) => e.label).join(", ") || "priority domains";

  const columns = [
    {
      title: "30 days",
      focus: `Stabilize weakest lanes: ${focusDomains}`,
      actions: liveAssessment.roadmap.d30,
      outcome: "Owners named; baseline artifacts and interim policy clarity published.",
      tint: "#eff6ff",
    },
    {
      title: "60 days",
      focus: "Industrialize execution across diligence and monitoring playbooks",
      actions: liveAssessment.roadmap.d60,
      outcome: "Standardized packs live; MI cadence and issue SLAs visible to governance.",
      tint: "#f0fdf4",
    },
    {
      title: "90 days",
      focus: "Assurance, funding lock, and sustained telemetry",
      actions: liveAssessment.roadmap.d90,
      outcome: "Independent QA sample complete; leadership scorecard and FY roadmap funded.",
      tint: "#fefce8",
    },
  ];

  return (
    <>
      <SectionCard
        title="Execution roadmap"
        subtitle={`Prioritized using lowest-scoring domains (current workspace): ${focusDomains}.`}
      >
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 18 }}>
          {columns.map((col) => (
            <div
              key={col.title}
              style={{
                borderRadius: 16,
                padding: 22,
                background: `linear-gradient(180deg, ${col.tint} 0%, #ffffff 65%)`,
                border: `1px solid ${C.border}`,
                boxShadow: "0 8px 28px rgba(15,23,42,0.08)",
                display: "flex",
                flexDirection: "column",
                minHeight: 420,
              }}
            >
              <div style={{ fontSize: 13, fontWeight: 900, color: C.accent, letterSpacing: "0.06em", marginBottom: 14 }}>
                {col.title}
              </div>
              <div style={{ fontSize: 11, fontWeight: 800, color: C.muted, textTransform: "uppercase", letterSpacing: "0.06em" }}>
                Focus area
              </div>
              <p style={{ margin: "8px 0 16px", fontSize: 13, lineHeight: 1.55, fontWeight: 600 }}>{col.focus}</p>
              <div style={{ fontSize: 11, fontWeight: 800, color: C.muted, textTransform: "uppercase", letterSpacing: "0.06em" }}>
                Actions
              </div>
              <ol style={{ margin: "8px 0 16px", paddingLeft: 18, fontSize: 13, lineHeight: 1.55, flex: 1 }}>
                {col.actions.map((a, i) => (
                  <li key={i} style={{ marginBottom: 8 }}>
                    {a}
                  </li>
                ))}
              </ol>
              <div style={{ fontSize: 11, fontWeight: 800, color: C.muted, textTransform: "uppercase", letterSpacing: "0.06em" }}>
                Expected outcome
              </div>
              <p style={{ margin: "8px 0 0", fontSize: 13, lineHeight: 1.55, color: C.text }}>{col.outcome}</p>
            </div>
          ))}
        </div>
      </SectionCard>

      <SectionCard title="Implementation priorities" subtitle="Gap order from the assessment engine (with evidence weighting).">
        <ol style={{ margin: 0, paddingLeft: 22, fontSize: 14, lineHeight: 1.6 }}>
          {liveAssessment.topGapsAll.slice(0, 5).map((g, i) => (
            <li key={i} style={{ marginBottom: 14 }}>
              <strong>{g.domain}</strong> — {g.score}/5 ({g.evidence})
              <div style={{ fontSize: 13, color: C.muted, marginTop: 4 }}>{g.rationale}</div>
            </li>
          ))}
        </ol>
      </SectionCard>
    </>
  );
}

function SettingsPageContent({ profileExtended, domainRows, liveAssessment, assessmentType, aiNarratives }) {
  return (
    <>
      <SectionCard title="Assessment framework" subtitle="Pinned configuration for repeatable runs (static controls in this build).">
        <Field label="Active framework">
          <select style={selectStyle} defaultValue="TPRM-CORE-2025-Q2">
            <option value="TPRM-CORE-2025-Q2">TPRM Core · 2025 Q2 (10-domain lifecycle)</option>
            <option value="TPRM-CORE-2024-Q4">TPRM Core · 2024 Q4 (legacy)</option>
          </select>
        </Field>
        <Field label="Assessment type (workspace)">
          <input readOnly value={assessmentType} style={{ ...selectStyle, cursor: "default", color: C.muted }} />
        </Field>
        <p style={{ margin: "14px 0 0", fontSize: 13, color: C.muted, lineHeight: 1.55 }}>
          The active framework governs domain definitions and scoring rubric. Version history supports alignment with your
          governance calendar.
        </p>
      </SectionCard>

      <SectionCard title="Benchmark target" subtitle="Overlay target used on Results for gap analysis.">
        <div style={{ display: "flex", flexWrap: "wrap", gap: 20, alignItems: "center" }}>
          <div>
            <div style={{ fontSize: 11, fontWeight: 700, color: C.muted }}>Sector benchmark (mean)</div>
            <div style={{ fontSize: 28, fontWeight: 900, marginTop: 4 }}>{BENCHMARK_SCORE}</div>
          </div>
          <p style={{ margin: 0, flex: "1 1 240px", fontSize: 13, color: C.muted, lineHeight: 1.55 }}>
            Reference mean used for relative gap analysis. Configure organization-specific benchmarks when integrating your
            peer or supervisory data sources.
          </p>
        </div>
      </SectionCard>

      <SectionCard title="Evidence weighting" subtitle="How evidence strength breaks ties in gap ordering.">
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#f1f5f9" }}>
              <th style={{ textAlign: "left", padding: 10, borderBottom: `1px solid ${C.border}` }}>Evidence</th>
              <th style={{ textAlign: "left", padding: 10, borderBottom: `1px solid ${C.border}` }}>Weight (subtracted in gap rank)</th>
            </tr>
          </thead>
          <tbody>
            {Object.entries(EVIDENCE_WEIGHT).map(([k, v]) => (
              <tr key={k} style={{ borderBottom: `1px solid ${C.border}` }}>
                <td style={{ padding: 10, fontWeight: 600 }}>{k}</td>
                <td style={{ padding: 10, fontVariantNumeric: "tabular-nums" }}>{v}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </SectionCard>

      <SectionCard title="Scoring definitions" subtitle="Numeric rubric labels for the 1–5 scale.">
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ textAlign: "left", color: C.muted, fontSize: 11 }}>
              <th style={{ padding: "8px 8px 8px 0", borderBottom: `1px solid ${C.border}` }}>Score</th>
              <th style={{ padding: 8, borderBottom: `1px solid ${C.border}` }}>Label</th>
              <th style={{ padding: 8, borderBottom: `1px solid ${C.border}` }}>In model</th>
            </tr>
          </thead>
          <tbody>
            {[1, 2, 3, 4, 5].map((n) => (
              <tr key={n}>
                <td style={{ padding: "10px 8px 10px 0", borderBottom: `1px solid ${C.border}`, fontWeight: 700 }}>{n}</td>
                <td style={{ padding: 10, borderBottom: `1px solid ${C.border}` }}>{SCORE_LABELS[n]}</td>
                <td style={{ padding: 10, borderBottom: `1px solid ${C.border}`, color: C.muted }}>
                  {n <= 2
                    ? "Priority gap band; weak-domain counts"
                    : n === 3
                      ? "Defined baseline; tie-breaker uses evidence"
                      : "Strength anchor for adjacent uplift"}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </SectionCard>

      <SectionCard
        title="Export & handoff"
        subtitle="Package the current workspace snapshot. For GPT-authored narratives on Results, set VITE_OPENAI_API_KEY in a local .env file (optional)."
      >
        <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
          <button
            type="button"
            style={{
              padding: "11px 18px",
              borderRadius: 10,
              border: `1px solid ${C.border}`,
              background: C.surface,
              fontWeight: 700,
              cursor: "pointer",
              boxShadow: "0 2px 8px rgba(15,23,42,0.06)",
            }}
            onClick={() => {
              const blob = new Blob(
                [JSON.stringify({ profile: profileExtended, domainRows, liveAssessment, aiNarratives }, null, 2)],
                { type: "application/json" }
              );
              const a = document.createElement("a");
              a.href = URL.createObjectURL(blob);
              a.download = "tprm-assessment-export.json";
              a.click();
              URL.revokeObjectURL(a.href);
            }}
          >
            Download JSON bundle
          </button>
          <button
            type="button"
            style={{
              padding: "11px 18px",
              borderRadius: 10,
              border: `1px solid ${C.border}`,
              background: C.surface,
              fontWeight: 700,
              cursor: "pointer",
              boxShadow: "0 2px 8px rgba(15,23,42,0.06)",
            }}
            onClick={() => {
              navigator.clipboard
                ?.writeText(aiNarratives?.executiveSummary ?? liveAssessment.executiveSummary)
                .catch(() => {});
            }}
          >
            Copy executive summary
          </button>
        </div>
      </SectionCard>
    </>
  );
}

export default function App() {
  const [activePage, setActivePage] = useState("assessment");
  const [profile, setProfile] = useState(defaultProfile);
  const [assessmentType, setAssessmentType] = useState(ASSESSMENT_TYPES[0]);
  const [scopeNotes, setScopeNotes] = useState("");
  const [supportingEvidenceText, setSupportingEvidenceText] = useState("");
  const [reviewFollowUpByDomain, setReviewFollowUpByDomain] = useState({});
  const [reviewerNotesByDomain, setReviewerNotesByDomain] = useState({});
  /** Command Center: manager sign-off for report readiness (pending | complete | not_required). */
  const [managerReviewStatus, setManagerReviewStatus] = useState("pending");
  const [domainRows, setDomainRows] = useState(initialDomainRows);
  const [aiSuggestionsByDomain, setAiSuggestionsByDomain] = useState({});
  const [analysisStatus, setAnalysisStatus] = useState("Not analyzed");
  const [analysisLoading, setAnalysisLoading] = useState(false);
  /** After first Generate, summary rail unlocks; all metrics use live workspace state. */
  const [hasGenerated, setHasGenerated] = useState(false);
  /** Populated when user runs Generate assessment (OpenAI or placeholder). */
  const [aiNarratives, setAiNarratives] = useState(null);
  const [narrativeLoading, setNarrativeLoading] = useState(false);

  const completionPct = useMemo(() => {
    let filled = 0;
    DOMAINS.forEach((d) => {
      const r = domainRows[d.id];
      if (r && r.score >= 1 && r.score <= 5 && r.evidence) filled += 1;
    });
    return Math.round((filled / DOMAINS.length) * 100);
  }, [domainRows]);

  const profileExtended = useMemo(
    () => ({ ...profile, assessmentType }),
    [profile, assessmentType]
  );

  const [memoryLibraryVersion, setMemoryLibraryVersion] = useState(0);
  /** Auto-suggest from library: per-domain panel payload + dismiss flag (persisted in draft). */
  const [librarySuggestionsByDomain, setLibrarySuggestionsByDomain] = useState({});

  const runGenerate = useCallback(() => {
    setHasGenerated(true);
    setActivePage("results");
    const assessment = buildAssessment(profileExtended, domainRows);
    const data = buildAssessmentDataForNarrative(assessment);
    setNarrativeLoading(true);
    setAiNarratives(null);
    generateNarrative(data, { profile: profileExtended, domainRows })
      .then((narratives) => {
        setAiNarratives(narratives);
        try {
          appendAssessmentMemoryLibraries(profileExtended, assessment, narratives);
        } catch (err) {
          console.warn("Assessment memory append failed:", err);
        }
        setMemoryLibraryVersion((v) => v + 1);
      })
      .finally(() => setNarrativeLoading(false));
  }, [profileExtended, domainRows]);

  const liveAssessment = useMemo(
    () => buildAssessment(profileExtended, domainRows),
    [profileExtended, domainRows]
  );

  const maturityBadgeColor =
    liveAssessment.maturityTier === "Low"
      ? "#b45309"
      : liveAssessment.maturityTier === "Moderate"
        ? "#1d4ed8"
        : "#047857";

  const [reportStatus, setReportStatus] = useState("Report generation ready");
  const [persistenceStatus, setPersistenceStatus] = useState("");
  const persistenceStatusTimerRef = useRef(null);

  const flashPersistenceStatus = useCallback((msg) => {
    if (persistenceStatusTimerRef.current) window.clearTimeout(persistenceStatusTimerRef.current);
    setPersistenceStatus(msg);
    persistenceStatusTimerRef.current = window.setTimeout(() => {
      setPersistenceStatus("");
      persistenceStatusTimerRef.current = null;
    }, 4500);
  }, []);

  const analyzeEvidenceAndSuggestScores = useCallback(async () => {
    setAnalysisLoading(true);
    setAnalysisStatus("Analyzing evidence...");
    try {
      const suggestions = {};
      for (const d of DOMAINS) {
        const row = domainRows[d.id];
        const suggestion = await analyzeDomainEvidenceWithLlm({
          profile,
          assessmentType,
          domain: d,
          row,
          supportingText: supportingEvidenceText,
        });
        suggestions[d.id] = {
          status: "Suggested",
          suggestion,
          source: (import.meta.env.VITE_OPENAI_API_KEY || "").trim() ? "openai" : "mock",
        };
      }
      setAiSuggestionsByDomain(suggestions);
      setAnalysisStatus("Suggested");
      setActivePage("assessment");
    } catch (error) {
      console.error("Evidence analysis failed:", error);
      setAnalysisStatus(`Analysis failed: ${error?.message || String(error)}`);
    } finally {
      setAnalysisLoading(false);
    }
  }, [domainRows, profile, assessmentType, supportingEvidenceText]);

  const acceptSuggestionForDomain = useCallback((domainId) => {
    const review = aiSuggestionsByDomain[domainId];
    if (!review?.suggestion) return;
    const suggestedScore = clampScore(review.suggestion.suggestedScore);
    setDomainRows((prev) => ({
      ...prev,
      [domainId]: {
        ...prev[domainId],
        score: suggestedScore,
      },
    }));
    setAiSuggestionsByDomain((prev) => ({
      ...prev,
      [domainId]: {
        ...prev[domainId],
        status: "Accepted",
      },
    }));
  }, [aiSuggestionsByDomain]);

  const keepManualScoreForDomain = useCallback((domainId) => {
    setAiSuggestionsByDomain((prev) => ({
      ...prev,
      [domainId]: {
        ...prev[domainId],
        status: "Overridden",
      },
    }));
  }, []);

  const markReviewFollowUp = useCallback((domainId) => {
    setReviewFollowUpByDomain((prev) => ({ ...prev, [domainId]: true }));
  }, []);

  const clearReviewFollowUp = useCallback((domainId) => {
    setReviewFollowUpByDomain((prev) => {
      const next = { ...prev };
      delete next[domainId];
      return next;
    });
  }, []);

  const onReviewerNoteChange = useCallback((domainId, text) => {
    setReviewerNotesByDomain((prev) => ({ ...prev, [domainId]: text }));
  }, []);

  const runLibrarySuggestForDomain = useCallback(
    (domainId) => {
      const d = DOMAINS.find((x) => x.id === domainId);
      if (!d) return;
      const row = domainRows[domainId] || {};
      const built = buildLibrarySuggestionForDomain(d, row, profileExtended, assessmentType, liveAssessment);
      setLibrarySuggestionsByDomain((prev) => ({
        ...prev,
        [domainId]: {
          dismissed: false,
          payload: built
            ? { match: true, ...built }
            : { match: false },
        },
      }));
    },
    [domainRows, profileExtended, assessmentType, liveAssessment]
  );

  const runLibrarySuggestAcrossAssessment = useCallback(() => {
    setLibrarySuggestionsByDomain((prev) => {
      const next = { ...prev };
      DOMAINS.forEach((d) => {
        const row = domainRows[d.id] || {};
        const built = buildLibrarySuggestionForDomain(d, row, profileExtended, assessmentType, liveAssessment);
        next[d.id] = {
          dismissed: false,
          payload: built ? { match: true, ...built } : { match: false },
        };
      });
      return next;
    });
    flashPersistenceStatus("Library suggestions populated for all domains.");
  }, [domainRows, profileExtended, assessmentType, liveAssessment, flashPersistenceStatus]);

  const dismissLibrarySuggestion = useCallback((domainId) => {
    setLibrarySuggestionsByDomain((prev) => ({
      ...prev,
      [domainId]: { dismissed: true, payload: null },
    }));
  }, []);

  const applyLibraryFindingToNotes = useCallback((domainId, text) => {
    const t = String(text || "").trim();
    if (!t) return;
    setDomainRows((prev) => {
      const cur = prev[domainId] || {};
      const join = cur.notes ? `${String(cur.notes).trim()}\n\n` : "";
      return {
        ...prev,
        [domainId]: { ...cur, notes: join + t },
      };
    });
  }, []);

  const applyLibraryRecommendationDraft = useCallback((domainId, text) => {
    const t = String(text || "").trim();
    setDomainRows((prev) => {
      const cur = prev[domainId] || {};
      return { ...prev, [domainId]: { ...cur, recommendationDraft: t } };
    });
  }, []);

  const appendLibraryEvidenceToNotes = useCallback((domainId, evidenceLine) => {
    const t = String(evidenceLine || "").trim();
    if (!t) return;
    setDomainRows((prev) => {
      const cur = prev[domainId] || {};
      const join = cur.notes ? `${String(cur.notes).trim()}\n\n` : "";
      return {
        ...prev,
        [domainId]: {
          ...cur,
          notes: `${join}[Evidence example — library]\n${t}`,
        },
      };
    });
  }, []);

  const acceptAllHighConfidenceSuggestions = useCallback(() => {
    const eligibleIds = DOMAINS.filter((d) => {
      const review = aiSuggestionsByDomain[d.id];
      return review?.suggestion?.confidence === "High";
    }).map((d) => d.id);
    if (eligibleIds.length === 0) return;

    setDomainRows((prev) => {
      const next = { ...prev };
      eligibleIds.forEach((domainId) => {
        const review = aiSuggestionsByDomain[domainId];
        next[domainId] = {
          ...next[domainId],
          score: clampScore(review.suggestion.suggestedScore),
        };
      });
      return next;
    });

    setAiSuggestionsByDomain((prev) => {
      const next = { ...prev };
      eligibleIds.forEach((domainId) => {
        next[domainId] = {
          ...next[domainId],
          status: "Accepted",
        };
      });
      return next;
    });
    setAnalysisStatus("Accepted");
  }, [aiSuggestionsByDomain]);

  const highConfidenceSuggestionCount = useMemo(
    () =>
      DOMAINS.filter((d) => {
        const review = aiSuggestionsByDomain[d.id];
        return review?.suggestion?.confidence === "High";
      }).length,
    [aiSuggestionsByDomain]
  );

  const clientDraft = useMemo(() => {
    if (!hasGenerated) return null;
    return buildClientAssessmentDraft(profileExtended, liveAssessment, aiNarratives);
  }, [hasGenerated, profileExtended, liveAssessment, aiNarratives]);

  const reportQualityWarnings = useMemo(
    () => computeReportQualityWarnings(profileExtended, liveAssessment, aiSuggestionsByDomain),
    [profileExtended, liveAssessment, aiSuggestionsByDomain]
  );

  const commandCenterModel = useMemo(
    () =>
      computeEngagementCommandCenterModel({
        profile: profileExtended,
        liveAssessment,
        aiSuggestionsByDomain,
        reviewFollowUpByDomain,
        supportingEvidenceText,
        scopeNotes,
        managerReviewStatus,
        completionPct,
      }),
    [
      profileExtended,
      liveAssessment,
      aiSuggestionsByDomain,
      reviewFollowUpByDomain,
      supportingEvidenceText,
      scopeNotes,
      managerReviewStatus,
      completionPct,
    ]
  );

  const saveAssessmentToLocalStorage = useCallback(() => {
    try {
      const payload = {
        v: PERSISTENCE_DRAFT_VERSION,
        savedAt: new Date().toISOString(),
        profile,
        assessmentType: ASSESSMENT_TYPES.includes(assessmentType) ? assessmentType : ASSESSMENT_TYPES[0],
        scopeNotes,
        domainRows,
        supportingEvidenceText,
        aiSuggestionsByDomain,
        reviewFollowUpByDomain,
        reviewerNotesByDomain,
        managerReviewStatus,
        hasGenerated,
        aiNarratives,
        analysisStatus,
        narrativeLoading: false,
        activePage,
        reportStatus,
        qaWarningsSnapshot: computeReportQualityWarnings(
          profileExtended,
          liveAssessment,
          aiSuggestionsByDomain
        ),
        librarySuggestionsByDomain,
      };
      localStorage.setItem(TPRM_ASSESSMENT_DRAFT_KEY, JSON.stringify(payload));
      flashPersistenceStatus("Saved successfully");
    } catch (err) {
      flashPersistenceStatus("Save failed: " + (err?.message || String(err)));
    }
  }, [
    profile,
    assessmentType,
    scopeNotes,
    domainRows,
    supportingEvidenceText,
    aiSuggestionsByDomain,
    reviewFollowUpByDomain,
    reviewerNotesByDomain,
    managerReviewStatus,
    hasGenerated,
    aiNarratives,
    analysisStatus,
    activePage,
    reportStatus,
    profileExtended,
    liveAssessment,
    flashPersistenceStatus,
    librarySuggestionsByDomain,
  ]);

  const loadAssessmentFromLocalStorage = useCallback(() => {
    try {
      const raw = localStorage.getItem(TPRM_ASSESSMENT_DRAFT_KEY);
      if (!raw) {
        flashPersistenceStatus("No saved draft found in this browser.");
        return;
      }
      const data = JSON.parse(raw);
      if (!data || typeof data.v !== "number" || (data.v !== 1 && data.v !== PERSISTENCE_DRAFT_VERSION)) {
        flashPersistenceStatus("Invalid or unsupported saved draft.");
        return;
      }
      setProfile(mergeLoadedProfile(data.profile));
      setAssessmentType(
        ASSESSMENT_TYPES.includes(data.assessmentType) ? data.assessmentType : ASSESSMENT_TYPES[0]
      );
      setScopeNotes(typeof data.scopeNotes === "string" ? data.scopeNotes : "");
      setSupportingEvidenceText(
        typeof data.supportingEvidenceText === "string" ? data.supportingEvidenceText : ""
      );
      setDomainRows(mergeLoadedDomainRows(data.domainRows));
      setAiSuggestionsByDomain(mergeLoadedAiSuggestionsByDomain(data.aiSuggestionsByDomain));
      setReviewFollowUpByDomain(mergeLoadedReviewFollowUp(data.reviewFollowUpByDomain));
      setReviewerNotesByDomain(mergeLoadedReviewerNotes(data.reviewerNotesByDomain));
      const mr = data.managerReviewStatus;
      setManagerReviewStatus(
        mr === "complete" || mr === "not_required" || mr === "pending" ? mr : "pending"
      );
      setHasGenerated(Boolean(data.hasGenerated));
      setAiNarratives(mergeLoadedAiNarratives(data.aiNarratives));
      setAnalysisStatus(typeof data.analysisStatus === "string" ? data.analysisStatus : "Not analyzed");
      setNarrativeLoading(false);
      setAnalysisLoading(false);
      setActivePage(sanitizeActivePage(data.activePage));
      setReportStatus(
        typeof data.reportStatus === "string" ? data.reportStatus : "Report generation ready"
      );
      setLibrarySuggestionsByDomain(mergeLoadedLibrarySuggestions(data.librarySuggestionsByDomain));
      setMemoryLibraryVersion((v) => v + 1);
      flashPersistenceStatus("Loaded successfully");
    } catch (err) {
      flashPersistenceStatus("Load failed: " + (err?.message || String(err)));
    }
  }, [flashPersistenceStatus]);

  useEffect(() => {
    if (!NAV_PAGE_IDS.has(activePage)) {
      setActivePage("assessment");
    }
  }, [activePage]);

  const startNewAssessment = useCallback(() => {
    if (
      !window.confirm(
        "Start a new assessment? The current workspace will be cleared. Save first if you need to keep this run."
      )
    ) {
      return;
    }
    setProfile({ ...defaultProfile });
    setAssessmentType(ASSESSMENT_TYPES[0]);
    setScopeNotes("");
    setSupportingEvidenceText("");
    setDomainRows(initialDomainRows());
    setAiSuggestionsByDomain({});
    setReviewFollowUpByDomain({});
    setReviewerNotesByDomain({});
    setManagerReviewStatus("pending");
    setHasGenerated(false);
    setAiNarratives(null);
    setAnalysisStatus("Not analyzed");
    setNarrativeLoading(false);
    setAnalysisLoading(false);
    setActivePage("assessment");
    setReportStatus("Report generation ready");
    setLibrarySuggestionsByDomain({});
    flashPersistenceStatus("New assessment started");
  }, [flashPersistenceStatus]);

  const loadDemoAssessment = useCallback(() => {
    const demoRows = buildDemoDomainRows();
    setProfile({ ...DEMO_ASSESSMENT_PROFILE });
    setAssessmentType(DEMO_ASSESSMENT_TYPE);
    setScopeNotes(DEMO_SCOPE_NOTES);
    setSupportingEvidenceText(DEMO_SUPPORTING_TEXT);
    setDomainRows(demoRows);
    setAiSuggestionsByDomain(
      buildDemoAiSuggestionsFromWorkspace(
        DEMO_ASSESSMENT_PROFILE,
        DEMO_ASSESSMENT_TYPE,
        demoRows,
        DEMO_SUPPORTING_TEXT
      )
    );
    setAnalysisStatus("Suggested");
    setHasGenerated(false);
    setAiNarratives(null);
    setNarrativeLoading(false);
    setAnalysisLoading(false);
    setReviewFollowUpByDomain({});
    setReviewerNotesByDomain({});
    setManagerReviewStatus("pending");
    setLibrarySuggestionsByDomain({});
    setReportStatus("Report generation ready");
    setActivePage("assessment");
    flashPersistenceStatus("Demo loaded");
  }, [flashPersistenceStatus]);

  const generateTPRMReport = async () => {
    try {
      setReportStatus("Generating report...");

      const pptx = new pptxgen();
      pptx.layout = "LAYOUT_WIDE";
      pptx.author = "Digital Risk";
      pptx.title = "TPRM Maturity Assessment Report";

      /** PptxGenJS expects 6-char hex without '#' */
      const TC = {
        title: "111827",
        body: "1F2937",
        sub: "374151",
        muted: "6B7280",
        headerBg: "1E3A8A",
        headerFg: "FFFFFF",
        border: "64748B",
      };

      const clip = (s, n) => {
        const t = String(s || "");
        return t.length <= n ? t : `${t.slice(0, n - 1)}…`;
      };

      const execBody =
        (aiNarratives?.executiveSummary && String(aiNarratives.executiveSummary).trim()) ||
        (liveAssessment.executiveSummary && String(liveAssessment.executiveSummary).trim()) ||
        "Use Generate assessment in the application header to produce an executive narrative for this workspace.";

      const gapLines =
        liveAssessment.topGapsAll?.length > 0
          ? liveAssessment.topGapsAll
              .slice(0, 5)
              .map((g) => `• ${g.domain} (${g.score}/5, ${g.evidence}): ${g.rationale}`)
              .join("\n")
          : "Enter domain scores and evidence to populate prioritized gaps for this export.";

      const roadmap = liveAssessment.roadmap;
      const roadmapFallback =
        "  • Assign remediation owners and dates in the Assessment workspace to populate this horizon.";
      const roadmapText = roadmap
        ? [
            "30 days — stabilize",
            ...(roadmap.d30?.length ? roadmap.d30.map((x) => `  • ${x}`) : [roadmapFallback]),
            "",
            "60 days — standardize",
            ...(roadmap.d60?.length ? roadmap.d60.map((x) => `  • ${x}`) : [roadmapFallback]),
            "",
            "90 days — optimize",
            ...(roadmap.d90?.length ? roadmap.d90.map((x) => `  • ${x}`) : [roadmapFallback]),
          ].join("\n")
        : "Remediation horizons populate when roadmap actions are available from the scored assessment.";

      const headerCell = (text) => ({
        text,
        options: {
          bold: true,
          fontFace: "Arial",
          fontSize: 11,
          color: TC.headerFg,
          fill: { color: TC.headerBg },
          valign: "middle",
          align: "center",
        },
      });

      const bodyCell = (text, align = "left") => ({
        text: clip(text, 160),
        options: {
          fontFace: "Arial",
          fontSize: 10,
          color: TC.title,
          valign: "middle",
          align,
        },
      });

      const maturityTableRows = [
        [
          headerCell("Domain"),
          headerCell("Score"),
          headerCell("Maturity band"),
          headerCell("Evidence"),
          headerCell("Priority"),
        ],
      ];

      const rawEntries = Array.isArray(liveAssessment.entries) ? liveAssessment.entries : [];
      if (rawEntries.length === 0) {
        maturityTableRows.push([
          bodyCell("No domain scores yet — complete the assessment worksheet.", "left"),
          bodyCell("—", "center"),
          bodyCell("—", "center"),
          bodyCell("—", "center"),
          bodyCell("—", "center"),
        ]);
      } else {
        rawEntries.forEach((e) => {
          const band = matrixMaturityBand(e.score).label;
          const pri = matrixPriority(e.score, e.evidence).label;
          maturityTableRows.push([
            bodyCell(e.label, "left"),
            bodyCell(String(e.score), "center"),
            bodyCell(band, "center"),
            bodyCell(String(e.evidence || "—"), "center"),
            bodyCell(pri, "center"),
          ]);
        });
      }

      let slideCount = 0;

      // Slide 1 — Title
      const slide1 = pptx.addSlide();
      slideCount += 1;
      slide1.background = { color: "FFFFFF" };
      slide1.addText("TPRM Maturity Assessment Report", {
        x: 0.5,
        y: 0.55,
        w: 12.3,
        h: 0.9,
        fontSize: 28,
        bold: true,
        color: TC.title,
        fontFace: "Arial",
      });
      slide1.addText("Generated from the TPRM Maturity Assessment Engine", {
        x: 0.5,
        y: 1.45,
        w: 12.3,
        h: 0.5,
        fontSize: 14,
        color: TC.sub,
        fontFace: "Arial",
      });
      slide1.addText(clip(reportClientLine(profileExtended), 120), {
        x: 0.5,
        y: 2.05,
        w: 12.3,
        h: 0.45,
        fontSize: 13,
        color: TC.body,
        fontFace: "Arial",
      });
      slide1.addText(formatReportDate(), {
        x: 0.5,
        y: 2.55,
        w: 12.3,
        h: 0.35,
        fontSize: 12,
        color: TC.muted,
        fontFace: "Arial",
      });

      // Slide 2 — Executive Summary
      const slide2 = pptx.addSlide();
      slideCount += 1;
      slide2.background = { color: "FFFFFF" };
      slide2.addText("Executive Summary", {
        x: 0.5,
        y: 0.45,
        w: 12.3,
        h: 0.6,
        fontSize: 22,
        bold: true,
        color: TC.title,
        fontFace: "Arial",
      });
      slide2.addText(clip(execBody, 2800), {
        x: 0.5,
        y: 1.15,
        w: 12.3,
        h: 5.45,
        fontSize: 12,
        color: TC.body,
        fontFace: "Arial",
        valign: "top",
        wrap: true,
      });

      // Slide 3 — Maturity table
      const slide3 = pptx.addSlide();
      slideCount += 1;
      slide3.background = { color: "FFFFFF" };
      slide3.addText("Maturity assessment table", {
        x: 0.5,
        y: 0.45,
        w: 12.3,
        h: 0.55,
        fontSize: 22,
        bold: true,
        color: TC.title,
        fontFace: "Arial",
      });
      slide3.addText(
        `Overall: ${liveAssessment.overallScore ?? "—"} / 5 · ${liveAssessment.maturityBand ?? "—"} (${liveAssessment.maturityTier ?? "—"} tier)`,
        {
          x: 0.5,
          y: 1.05,
          w: 12.3,
          h: 0.38,
          fontSize: 13,
          color: TC.body,
          fontFace: "Arial",
        }
      );
      slide3.addTable(maturityTableRows, {
        x: 0.5,
        y: 1.55,
        w: 12.3,
        colW: [4.15, 0.85, 1.45, 1.35, 1.55],
        border: { type: "solid", color: TC.border, pt: 1 },
        fontFace: "Arial",
        rowH: 0.34,
      });

      // Slide 4 — Top Gaps
      const slide4 = pptx.addSlide();
      slideCount += 1;
      slide4.background = { color: "FFFFFF" };
      slide4.addText("Top Gaps", {
        x: 0.5,
        y: 0.45,
        w: 12.3,
        h: 0.55,
        fontSize: 22,
        bold: true,
        color: TC.title,
        fontFace: "Arial",
      });
      slide4.addText(clip(gapLines, 4500), {
        x: 0.5,
        y: 1.1,
        w: 12.3,
        h: 5.25,
        fontSize: 12,
        color: TC.body,
        fontFace: "Arial",
        valign: "top",
        wrap: true,
      });

      // Slide 5 — 30 / 60 / 90 Roadmap
      const slide5 = pptx.addSlide();
      slideCount += 1;
      slide5.background = { color: "FFFFFF" };
      slide5.addText("30 / 60 / 90 Day Roadmap", {
        x: 0.5,
        y: 0.45,
        w: 12.3,
        h: 0.55,
        fontSize: 22,
        bold: true,
        color: TC.title,
        fontFace: "Arial",
      });
      slide5.addText(clip(roadmapText, 4500), {
        x: 0.5,
        y: 1.1,
        w: 12.3,
        h: 5.25,
        fontSize: 12,
        color: TC.body,
        fontFace: "Arial",
        valign: "top",
        wrap: true,
      });

      console.log(`Slides created: ${slideCount}`);

      const arrayBuffer = await pptx.write({ outputType: "arraybuffer" });

      const blob = new Blob([arrayBuffer], {
        type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      });

      const url = window.URL.createObjectURL(blob);

      const link = document.createElement("a");
      link.href = url;
      link.download = "TPRM_Maturity_Assessment_Report.pptx";

      document.body.appendChild(link);
      link.click();

      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

      const qw = computeReportQualityWarnings(profileExtended, liveAssessment, aiSuggestionsByDomain);
      setReportStatus(
        qw.length
          ? `Report downloaded · ${qw.length} pre-report warning(s) — see Pre-report quality check above`
          : "Report downloaded"
      );
    } catch (error) {
      console.error(error);
      setReportStatus("Report failed: " + error.message);
    }
  };

  const header = PAGE_HEADER[activePage] ?? PAGE_HEADER.assessment;

  const domainsLoadedCount = Object.keys(domainRows).length;
  const domainsCatalogCount = DOMAINS.length;

  let persistedDraftPresent = false;
  try {
    if (typeof localStorage !== "undefined") {
      persistedDraftPresent = Boolean(localStorage.getItem(TPRM_ASSESSMENT_DRAFT_KEY));
    }
  } catch {
    persistedDraftPresent = false;
  }

  const workspaceMatchesDemoSeed = useMemo(() => {
    if (JSON.stringify(profile) !== JSON.stringify(defaultProfile)) return false;
    const base = initialDomainRows();
    return DOMAINS.every((d) => {
      const r = domainRows[d.id];
      const b = base[d.id];
      if (!r || !b) return false;
      return (
        r.score === b.score &&
        r.evidence === b.evidence &&
        String(r.notes || "").trim() === "" &&
        String(r.recommendationDraft || "").trim() === ""
      );
    });
  }, [profile, domainRows]);

  const demoDataLoadedLabel = persistedDraftPresent
    ? "Yes (saved draft in browser)"
    : workspaceMatchesDemoSeed
      ? "No (in-memory default seed only)"
      : "No (customized workspace, no saved draft)";

  return (
    <div
      className="tprm-app-root"
      style={{
        minHeight: "100vh",
        width: "100%",
        display: "flex",
        flexDirection: "column",
      }}
    >
      <style>{`
        html, body, #root { height: 100%; margin: 0; }
        * { box-sizing: border-box; }
        body {
          font-family: "IBM Plex Sans", system-ui, -apple-system, "Segoe UI", sans-serif;
          -webkit-font-smoothing: antialiased;
        }
        select, textarea, button { font: inherit; }
        textarea { resize: vertical; min-height: 56px; line-height: 1.5; }
      `}</style>

      {import.meta.env.DEV ? (
        <div
          role="status"
          aria-label="Debug panel"
          style={{
            flexShrink: 0,
            width: "100%",
            padding: "10px 14px",
            background: "#fef9c3",
            borderBottom: "3px solid #ca8a04",
            fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", monospace',
            fontSize: 12,
            lineHeight: 1.45,
            color: "#422006",
            boxShadow: "0 2px 8px rgba(180, 83, 9, 0.15)",
            zIndex: 50,
          }}
        >
          <div style={{ fontWeight: 800, letterSpacing: "0.06em", marginBottom: 6, fontSize: 11, color: "#713f12" }}>
            DEBUG
          </div>
          <div
            style={{
              display: "flex",
              flexWrap: "wrap",
              gap: "6px 28px",
              alignItems: "baseline",
            }}
          >
            <span>
              <strong>Page:</strong> {activePage}
            </span>
            <span style={{ maxWidth: "100%", wordBreak: "break-word" }}>
              <strong>Report status:</strong> {reportStatus}
            </span>
            <span>
              <strong>Domains loaded:</strong> {domainsLoadedCount} / {domainsCatalogCount}
            </span>
            <span style={{ maxWidth: "100%", wordBreak: "break-word" }}>
              <strong>Demo data loaded:</strong> {demoDataLoadedLabel}
            </span>
          </div>
        </div>
      ) : null}

      <div
        style={{
          display: "flex",
          flex: 1,
          minHeight: 0,
          width: "100%",
          background: "linear-gradient(165deg, #e2e8f0 0%, #f1f5f9 45%, #f8fafc 100%)",
          color: C.text,
        }}
      >
        {/* Sidebar */}
        <aside
          aria-label="Primary navigation"
          style={{
            width: 240,
            flexShrink: 0,
            background: C.sidebar,
            color: "#fff",
            display: "flex",
            flexDirection: "column",
            padding: "22px 0",
            borderRight: "1px solid #1e293b",
          }}
        >
          <div style={{ padding: "0 20px 20px", borderBottom: "1px solid #334155" }}>
            <div
              style={{
                fontSize: 10,
                letterSpacing: "0.14em",
                textTransform: "uppercase",
                color: C.sidebarMuted,
                fontWeight: 700,
              }}
            >
              Digital Risk
            </div>
            <div
              style={{
                fontSize: 14,
                fontWeight: 800,
                marginTop: 8,
                lineHeight: 1.3,
                wordBreak: "break-word",
              }}
            >
              TPRM Maturity Assessment Engine
            </div>
          </div>
          <nav style={{ padding: "16px 12px", flex: 1 }}>
            {NAV_ITEMS.map(({ id, label }) => {
              const active = activePage === id;
              return (
                <button
                  key={id}
                  type="button"
                  onClick={() => setActivePage(id)}
                  aria-current={active ? "page" : undefined}
                  style={{
                    display: "block",
                    width: "100%",
                    textAlign: "left",
                    padding: "11px 14px",
                    borderRadius: 8,
                    marginBottom: 6,
                    fontSize: 13,
                    fontWeight: active ? 700 : 500,
                    background: active ? "rgba(37,99,235,0.25)" : "transparent",
                    color: active ? "#fff" : C.sidebarMuted,
                    cursor: "pointer",
                    border: active ? "1px solid rgba(59,130,246,0.35)" : "1px solid transparent",
                  }}
                >
                  {label}
                </button>
              );
            })}
          </nav>
          <div
            style={{
              padding: "12px 20px",
              fontSize: 11,
              color: C.sidebarMuted,
              borderTop: "1px solid #334155",
            }}
          >
            Digital Risk Platform · v1
          </div>
        </aside>

        {/* Main column */}
        <div style={{ flex: 1, display: "flex", flexDirection: "column", minWidth: 0 }}>
          {/* Header */}
          <header
            style={{
              minHeight: 62,
              flexShrink: 0,
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              gap: 14,
              padding: "14px 32px",
              background: "linear-gradient(180deg, #ffffff 0%, #f8fafc 100%)",
              borderBottom: `1px solid ${C.border}`,
              boxShadow: "0 4px 20px rgba(15,23,42,0.06)",
            }}
          >
            <div>
              <div style={{ fontSize: 11, color: C.muted, fontWeight: 600 }}>
                {header.eyebrow}
              </div>
              <h1
                style={{
                  margin: "4px 0 2px",
                  fontSize: 20,
                  fontWeight: 800,
                  letterSpacing: "-0.02em",
                  color: C.text,
                }}
              >
                TPRM Maturity Assessment Engine
              </h1>
              {header.title !== "TPRM Maturity Assessment Engine" && (
                <div
                  style={{
                    fontSize: 15,
                    fontWeight: 700,
                    color: C.text,
                    letterSpacing: "-0.01em",
                  }}
                >
                  {header.title}
                </div>
              )}
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 10,
                flexWrap: "wrap",
                justifyContent: "flex-end",
              }}
            >
              <div style={{ textAlign: "right" }}>
                <div style={{ fontSize: 10, color: C.muted, fontWeight: 600 }}>
                  Workspace completion
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 4 }}>
                  <div
                    style={{
                      width: 120,
                      height: 6,
                      borderRadius: 999,
                      background: "#e2e8f0",
                      overflow: "hidden",
                    }}
                  >
                    <div
                      style={{
                        width: `${completionPct}%`,
                        height: "100%",
                        borderRadius: 999,
                        background: `linear-gradient(90deg, ${C.accent}, #3b82f6)`,
                        transition: "width 0.25s ease",
                      }}
                    />
                  </div>
                  <span style={{ fontSize: 13, fontWeight: 700, color: C.accent }}>
                    {completionPct}%
                  </span>
                </div>
              </div>
              <button
                type="button"
                onClick={runGenerate}
                title="Runs scoring against current workspace and opens Results"
                style={{
                  background: `linear-gradient(180deg, #2563eb 0%, #1d4ed8 100%)`,
                  color: "#fff",
                  border: "none",
                  borderRadius: 10,
                  padding: "10px 20px",
                  fontWeight: 700,
                  fontSize: 13,
                  cursor: "pointer",
                  boxShadow: "0 4px 14px rgba(37,99,235,0.35)",
                }}
              >
                Generate assessment
              </button>
              <button
                type="button"
                onClick={analyzeEvidenceAndSuggestScores}
                disabled={analysisLoading}
                style={{
                  background: analysisLoading
                    ? "#94a3b8"
                    : "linear-gradient(180deg, #0f766e 0%, #0f766e 100%)",
                  color: "#fff",
                  border: "none",
                  borderRadius: 10,
                  padding: "10px 16px",
                  fontWeight: 700,
                  fontSize: 13,
                  cursor: analysisLoading ? "not-allowed" : "pointer",
                  boxShadow: "0 4px 14px rgba(15,118,110,0.3)",
                }}
              >
                {analysisLoading ? "Analyzing evidence..." : "Analyze Evidence & Suggest Scores"}
              </button>
              <button
                type="button"
                onClick={acceptAllHighConfidenceSuggestions}
                disabled={analysisLoading || highConfidenceSuggestionCount === 0}
                style={{
                  background:
                    analysisLoading || highConfidenceSuggestionCount === 0
                      ? "#94a3b8"
                      : "linear-gradient(180deg, #14532d 0%, #166534 100%)",
                  color: "#fff",
                  border: "none",
                  borderRadius: 10,
                  padding: "10px 16px",
                  fontWeight: 700,
                  fontSize: 13,
                  cursor:
                    analysisLoading || highConfidenceSuggestionCount === 0
                      ? "not-allowed"
                      : "pointer",
                  boxShadow: "0 4px 14px rgba(22,101,52,0.3)",
                }}
              >
                Accept All High-Confidence Suggestions
              </button>
              <button
                type="button"
                onClick={saveAssessmentToLocalStorage}
                style={{
                  padding: "9px 14px",
                  borderRadius: 10,
                  border: `1px solid ${C.border}`,
                  background: C.surface,
                  fontWeight: 700,
                  fontSize: 12,
                  cursor: "pointer",
                  color: C.text,
                }}
              >
                Save Assessment
              </button>
              <button
                type="button"
                onClick={loadDemoAssessment}
                style={{
                  padding: "10px 16px",
                  borderRadius: 10,
                  border: "2px solid #d97706",
                  background: "linear-gradient(180deg, #fef3c7 0%, #fde68a 100%)",
                  fontWeight: 800,
                  fontSize: 13,
                  cursor: "pointer",
                  color: "#78350f",
                  boxShadow: "0 2px 12px rgba(217,119,6,0.35)",
                }}
              >
                Load Demo Assessment
              </button>
              <button
                type="button"
                onClick={loadAssessmentFromLocalStorage}
                style={{
                  padding: "9px 14px",
                  borderRadius: 10,
                  border: `1px solid ${C.border}`,
                  background: C.surface,
                  fontWeight: 700,
                  fontSize: 12,
                  cursor: "pointer",
                  color: C.text,
                }}
              >
                Load Assessment
              </button>
              <button
                type="button"
                onClick={startNewAssessment}
                style={{
                  padding: "9px 14px",
                  borderRadius: 10,
                  border: "1px solid #fecaca",
                  background: "#fef2f2",
                  fontWeight: 700,
                  fontSize: 12,
                  cursor: "pointer",
                  color: "#991b1b",
                }}
              >
                New Assessment
              </button>
              {persistenceStatus ? (
                <span
                  style={{
                    fontSize: 12,
                    fontWeight: 700,
                    color:
                      persistenceStatus === "Saved successfully" ||
                      persistenceStatus === "Loaded successfully" ||
                      persistenceStatus === "New assessment started" ||
                      persistenceStatus === "Demo loaded"
                        ? "#15803d"
                        : "#b45309",
                    whiteSpace: "nowrap",
                  }}
                >
                  {persistenceStatus}
                </span>
              ) : null}
            </div>
          </header>

          {/* Body: workspace + right rail */}
          <div style={{ display: "flex", flex: 1, minHeight: 0 }}>
            {/* Scrollable workspace */}
            <main
              style={{
                flex: 1,
                overflowY: "auto",
                padding: "28px 32px 56px",
              }}
            >
              {activePage === "commandCenter" && (
                <CommandCenterPageContent
                  liveAssessment={liveAssessment}
                  completionPct={completionPct}
                  model={commandCenterModel}
                  managerReviewStatus={managerReviewStatus}
                  onManagerReviewStatus={setManagerReviewStatus}
                  analysisLoading={analysisLoading}
                  analysisStatus={analysisStatus}
                  reportQualityWarningCount={reportQualityWarnings.length}
                  onNavigate={setActivePage}
                  onAnalyzeEvidence={analyzeEvidenceAndSuggestScores}
                  onAcceptAllHighConfidence={acceptAllHighConfidenceSuggestions}
                  onGenerateReport={generateTPRMReport}
                />
              )}

              {activePage === "assessment" && (
              <>
              <div
                style={{
                  marginBottom: 24,
                  padding: "18px 20px",
                  borderRadius: 14,
                  border: "2px solid #f59e0b",
                  background: "linear-gradient(135deg, #fffbeb 0%, #fef3c7 55%, #fff 100%)",
                  boxShadow: "0 8px 28px rgba(245,158,11,0.18)",
                  display: "flex",
                  flexWrap: "wrap",
                  alignItems: "center",
                  justifyContent: "space-between",
                  gap: 16,
                }}
              >
                <div style={{ flex: "1 1 280px", minWidth: 0 }}>
                  <div
                    style={{
                      fontSize: 11,
                      fontWeight: 800,
                      letterSpacing: "0.08em",
                      color: "#b45309",
                      textTransform: "uppercase",
                      marginBottom: 6,
                    }}
                  >
                    Sample workspace
                  </div>
                  <div style={{ fontSize: 16, fontWeight: 800, color: C.text, marginBottom: 6 }}>
                    Load sample assessment data
                  </div>
                  <p style={{ margin: 0, fontSize: 14, color: C.muted, lineHeight: 1.55 }}>
                    Applies a full client profile, domain scores, evidence notes, supporting text, and AI suggestions. Your
                    browser-saved draft is unchanged until you choose Save.
                  </p>
                </div>
                <button
                  type="button"
                  onClick={loadDemoAssessment}
                  style={{
                    flexShrink: 0,
                    padding: "14px 22px",
                    borderRadius: 12,
                    border: "2px solid #d97706",
                    background: "linear-gradient(180deg, #fbbf24 0%, #f59e0b 100%)",
                    fontWeight: 800,
                    fontSize: 15,
                    cursor: "pointer",
                    color: "#451a03",
                    boxShadow: "0 6px 20px rgba(217,119,6,0.45)",
                  }}
                >
                  Load Demo Assessment
                </button>
              </div>
              {/* Client Profile */}
              <SectionCard title="Client profile" subtitle="Organizational context shapes interpretation and remediation sequencing.">
                <div
                  style={{
                    display: "grid",
                    gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))",
                    gap: 16,
                  }}
                >
                  {[
                    ["industry", "Industry", PROFILE_OPTIONS.industry],
                    ["companySize", "Company size", PROFILE_OPTIONS.companySize],
                    [
                      "regulatoryIntensity",
                      "Regulatory intensity",
                      PROFILE_OPTIONS.regulatoryIntensity,
                    ],
                    [
                      "thirdPartyVolume",
                      "Third-party volume",
                      PROFILE_OPTIONS.thirdPartyVolume,
                    ],
                    [
                      "geographicFootprint",
                      "Geographic footprint",
                      PROFILE_OPTIONS.geographicFootprint,
                    ],
                  ].map(([key, label, opts]) => (
                    <Field key={key} label={label}>
                      <select
                        value={profile[key]}
                        onChange={(e) =>
                          setProfile((p) => ({ ...p, [key]: e.target.value }))
                        }
                        style={selectStyle}
                      >
                        {opts.map((o) => (
                          <option key={o} value={o}>
                            {o}
                          </option>
                        ))}
                      </select>
                    </Field>
                  ))}
                  <div style={{ gridColumn: "1 / -1" }}>
                    <Field label="Client / organization name (optional, for report cover)">
                      <input
                        type="text"
                        value={profile.clientName ?? ""}
                        onChange={(e) => setProfile((p) => ({ ...p, clientName: e.target.value }))}
                        placeholder="Organization name (optional)"
                        style={selectStyle}
                      />
                    </Field>
                  </div>
                </div>
              </SectionCard>

              {/* Assessment Scope */}
              <SectionCard title="Assessment scope" subtitle="Define the lens for this review cycle.">
                <Field label="Assessment type">
                  <select
                    value={assessmentType}
                    onChange={(e) => setAssessmentType(e.target.value)}
                    style={selectStyle}
                  >
                    {ASSESSMENT_TYPES.map((t) => (
                      <option key={t} value={t}>
                        {t}
                      </option>
                    ))}
                  </select>
                </Field>
                <Field label="Scope notes (optional)">
                  <textarea
                    value={scopeNotes}
                    onChange={(e) => setScopeNotes(e.target.value)}
                    placeholder="Scope, entities, and review period for this cycle"
                    style={{ ...selectStyle, minHeight: 72, lineHeight: 1.5 }}
                  />
                </Field>
                <Field label="Supporting evidence text (paste artifacts, excerpts, summaries)">
                  <textarea
                    value={supportingEvidenceText}
                    onChange={(e) => setSupportingEvidenceText(e.target.value)}
                    placeholder="Cross-domain evidence: audits, policies, test results, committee materials"
                    style={{ ...selectStyle, minHeight: 90, lineHeight: 1.5 }}
                  />
                </Field>
              </SectionCard>

              {/* Domain scoring (includes evidence & notes per domain) */}
              <SectionCard
                title="Domain scoring"
                subtitle="Evidence strength and assessor notes are captured alongside each score—both inform gap prioritization."
              >
                <div
                  style={{
                    fontSize: 11,
                    fontWeight: 800,
                    letterSpacing: "0.08em",
                    textTransform: "uppercase",
                    color: C.muted,
                    marginBottom: 12,
                  }}
                >
                  Evidence & notes · per domain
                </div>
                <div
                  style={{
                    marginBottom: 12,
                    padding: "10px 12px",
                    borderRadius: 10,
                    background: "#f8fafc",
                    border: `1px solid ${C.border}`,
                    fontSize: 12,
                    color: C.muted,
                  }}
                >
                  AI review status: <strong style={{ color: C.text }}>{analysisStatus}</strong> · AI suggests scores only;
                  analyst review and approval remain required.
                  {" "}High-confidence suggestions available: <strong style={{ color: C.text }}>{highConfidenceSuggestionCount}</strong>.
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
                  {DOMAINS.map((d) => {
                    const row = domainRows[d.id];
                    const aiReview = aiSuggestionsByDomain[d.id];
                    const aiSuggestion = aiReview?.suggestion;
                    const libCell = librarySuggestionsByDomain[d.id];
                    const libPayload =
                      libCell && !libCell.dismissed && libCell.payload != null ? libCell.payload : null;
                    const libConfPalette =
                      libPayload?.match === true
                        ? libPayload.confidence === "Strong match"
                          ? { bg: "#ecfdf5", fg: "#047857", bd: "#6ee7b7" }
                          : libPayload.confidence === "Moderate match"
                            ? { bg: "#fffbeb", fg: "#b45309", bd: "#fcd34d" }
                            : { bg: "#f1f5f9", fg: "#475569", bd: "#cbd5e1" }
                        : null;
                    return (
                      <div
                        key={d.id}
                        style={{
                          border: `1px solid ${C.border}`,
                          borderRadius: 12,
                          padding: 16,
                          background: C.surface,
                          boxShadow: "0 1px 3px rgba(15,23,42,0.05)",
                        }}
                      >
                        <div
                          style={{
                            display: "flex",
                            flexWrap: "wrap",
                            alignItems: "flex-start",
                            justifyContent: "space-between",
                            gap: 12,
                            marginBottom: 10,
                          }}
                        >
                          <div style={{ flex: "1 1 220px" }}>
                            <div style={{ fontWeight: 700, fontSize: 14 }}>{d.label}</div>
                            <p style={{ margin: "6px 0 0", fontSize: 12, color: C.muted, lineHeight: 1.5 }}>
                              {d.description}
                            </p>
                          </div>
                          <div style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 10 }}>
                            <ScoreBadge score={row.score} />
                            <button
                              type="button"
                              onClick={() => runLibrarySuggestForDomain(d.id)}
                              style={{
                                padding: "8px 12px",
                                borderRadius: 8,
                                border: `1px solid ${C.accent}`,
                                background: C.accentSoft,
                                fontWeight: 700,
                                fontSize: 11,
                                cursor: "pointer",
                                color: C.accent,
                                whiteSpace: "nowrap",
                              }}
                            >
                              Suggest from Library
                            </button>
                          </div>
                        </div>
                        <div
                          style={{
                            display: "grid",
                            gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))",
                            gap: 12,
                            alignItems: "start",
                          }}
                        >
                          <Field label="Maturity score">
                            <select
                              value={row.score}
                              onChange={(e) =>
                                {
                                  const nextScore = Number(e.target.value);
                                  setDomainRows((prev) => ({
                                    ...prev,
                                    [d.id]: {
                                      ...prev[d.id],
                                      score: nextScore,
                                    },
                                  }));
                                  if (
                                    aiSuggestion &&
                                    nextScore !== clampScore(aiSuggestion.suggestedScore)
                                  ) {
                                    setAiSuggestionsByDomain((prev) => ({
                                      ...prev,
                                      [d.id]: {
                                        ...prev[d.id],
                                        status: "Overridden",
                                      },
                                    }));
                                  }
                                }
                              }
                              style={selectStyle}
                            >
                              {[1, 2, 3, 4, 5].map((n) => (
                                <option key={n} value={n}>
                                  {n} — {SCORE_LABELS[n]}
                                </option>
                              ))}
                            </select>
                          </Field>
                          <Field label="Evidence strength">
                            <select
                              value={row.evidence}
                              onChange={(e) =>
                                setDomainRows((prev) => ({
                                  ...prev,
                                  [d.id]: {
                                    ...prev[d.id],
                                    evidence: e.target.value,
                                  },
                                }))
                              }
                              style={selectStyle}
                            >
                              {["Weak", "Moderate", "Strong"].map((ev) => (
                                <option key={ev} value={ev}>
                                  {ev}
                                </option>
                              ))}
                            </select>
                          </Field>
                          <div style={{ gridColumn: "1 / -1" }}>
                            <Field label="Assessor notes">
                              <textarea
                                value={row.notes}
                                onChange={(e) =>
                                  setDomainRows((prev) => ({
                                    ...prev,
                                    [d.id]: {
                                      ...prev[d.id],
                                      notes: e.target.value,
                                    },
                                  }))
                                }
                                placeholder="Control tests, interviews, and documentation for this domain"
                                style={{ ...selectStyle, minHeight: 64 }}
                              />
                            </Field>
                          </div>
                          <div style={{ gridColumn: "1 / -1" }}>
                            <Field label="Domain recommendation (draft — library or analyst)">
                              <textarea
                                value={row.recommendationDraft ?? ""}
                                onChange={(e) =>
                                  setDomainRows((prev) => ({
                                    ...prev,
                                    [d.id]: {
                                      ...prev[d.id],
                                      recommendationDraft: e.target.value,
                                    },
                                  }))
                                }
                                placeholder="Draft recommendation language (optional)"
                                style={{ ...selectStyle, minHeight: 56 }}
                              />
                            </Field>
                          </div>
                          {libPayload && (
                            <div style={{ gridColumn: "1 / -1" }}>
                              {libPayload.match === false ? (
                                <div
                                  style={{
                                    borderRadius: 10,
                                    border: `1px solid ${C.border}`,
                                    background: "#f8fafc",
                                    padding: 12,
                                    fontSize: 13,
                                    color: C.muted,
                                    lineHeight: 1.55,
                                  }}
                                >
                                  No close library match found. Use AI analysis or create a new library item after
                                  review.
                                </div>
                              ) : (
                                <div
                                  style={{
                                    borderRadius: 10,
                                    border: `1px solid ${libConfPalette?.bd || C.border}`,
                                    background: libConfPalette?.bg || "#f8fafc",
                                    padding: 12,
                                  }}
                                >
                                  <div
                                    style={{
                                      display: "flex",
                                      flexWrap: "wrap",
                                      alignItems: "center",
                                      justifyContent: "space-between",
                                      gap: 8,
                                      marginBottom: 10,
                                    }}
                                  >
                                    <div style={{ fontSize: 12, fontWeight: 800, color: C.text }}>
                                      Library suggestions
                                    </div>
                                    <span
                                      style={{
                                        display: "inline-block",
                                        padding: "3px 10px",
                                        borderRadius: 999,
                                        fontSize: 10,
                                        fontWeight: 800,
                                        color: libConfPalette?.fg || C.text,
                                        background: "#fff",
                                        border: `1px solid ${libConfPalette?.bd || C.border}`,
                                      }}
                                    >
                                      {libPayload.confidence}
                                    </span>
                                  </div>
                                  {libPayload.findingText ? (
                                    <p style={{ margin: "0 0 8px", fontSize: 12, color: C.text, lineHeight: 1.55 }}>
                                      <strong>Suggested finding language.</strong> {libPayload.findingText}
                                    </p>
                                  ) : null}
                                  {libPayload.recommendationText ? (
                                    <p style={{ margin: "0 0 8px", fontSize: 12, color: C.text, lineHeight: 1.55 }}>
                                      <strong>Suggested recommendation language.</strong> {libPayload.recommendationText}
                                    </p>
                                  ) : null}
                                  {libPayload.evidenceExamples?.length > 0 ? (
                                    <div style={{ margin: "0 0 8px", fontSize: 12, color: C.text, lineHeight: 1.55 }}>
                                      <strong>Common evidence examples.</strong>
                                      <ul style={{ margin: "6px 0 0", paddingLeft: 18 }}>
                                        {libPayload.evidenceExamples.map((ex, i) => (
                                          <li key={i} style={{ marginBottom: 4 }}>
                                            {ex}
                                          </li>
                                        ))}
                                      </ul>
                                    </div>
                                  ) : null}
                                  {libPayload.roadmapAction ? (
                                    <p style={{ margin: "0 0 8px", fontSize: 12, color: C.text, lineHeight: 1.55 }}>
                                      <strong>Suggested roadmap action.</strong> {libPayload.roadmapAction}
                                    </p>
                                  ) : null}
                                  {libPayload.similarPatternNote ? (
                                    <p
                                      style={{
                                        margin: "0 0 10px",
                                        fontSize: 11,
                                        color: C.muted,
                                        lineHeight: 1.5,
                                        fontStyle: "italic",
                                      }}
                                    >
                                      <strong>Similar pattern note.</strong> {libPayload.similarPatternNote}
                                    </p>
                                  ) : null}
                                  <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                                    <button
                                      type="button"
                                      onClick={() =>
                                        applyLibraryFindingToNotes(d.id, libPayload.findingText)
                                      }
                                      disabled={!libPayload.findingText}
                                      style={{
                                        padding: "6px 10px",
                                        borderRadius: 8,
                                        border: `1px solid ${C.border}`,
                                        background: "#fff",
                                        fontWeight: 700,
                                        fontSize: 11,
                                        cursor: libPayload.findingText ? "pointer" : "not-allowed",
                                        opacity: libPayload.findingText ? 1 : 0.5,
                                      }}
                                    >
                                      Apply finding to notes
                                    </button>
                                    <button
                                      type="button"
                                      onClick={() =>
                                        applyLibraryRecommendationDraft(d.id, libPayload.recommendationText)
                                      }
                                      disabled={!libPayload.recommendationText}
                                      style={{
                                        padding: "6px 10px",
                                        borderRadius: 8,
                                        border: `1px solid ${C.accent}`,
                                        background: C.accentSoft,
                                        fontWeight: 700,
                                        fontSize: 11,
                                        cursor: libPayload.recommendationText ? "pointer" : "not-allowed",
                                        opacity: libPayload.recommendationText ? 1 : 0.5,
                                        color: C.accent,
                                      }}
                                    >
                                      Apply recommendation
                                    </button>
                                    {libPayload.evidenceExamples?.map((ex, i) => (
                                      <button
                                        key={i}
                                        type="button"
                                        onClick={() => appendLibraryEvidenceToNotes(d.id, ex)}
                                        style={{
                                          padding: "6px 10px",
                                          borderRadius: 8,
                                          border: `1px solid ${C.border}`,
                                          background: "#fff",
                                          fontWeight: 600,
                                          fontSize: 11,
                                          cursor: "pointer",
                                        }}
                                      >
                                        Add evidence example {i + 1}
                                      </button>
                                    ))}
                                    <button
                                      type="button"
                                      onClick={() => dismissLibrarySuggestion(d.id)}
                                      style={{
                                        padding: "6px 10px",
                                        borderRadius: 8,
                                        border: "none",
                                        background: "#e2e8f0",
                                        fontWeight: 700,
                                        fontSize: 11,
                                        cursor: "pointer",
                                        color: "#334155",
                                      }}
                                    >
                                      Dismiss suggestion
                                    </button>
                                  </div>
                                </div>
                              )}
                            </div>
                          )}
                          {aiSuggestion && (
                            <div style={{ gridColumn: "1 / -1" }}>
                              <div
                                style={{
                                  border: "1px solid #bfdbfe",
                                  borderRadius: 10,
                                  background: "#eff6ff",
                                  padding: 12,
                                }}
                              >
                                <div
                                  style={{
                                    display: "flex",
                                    flexWrap: "wrap",
                                    alignItems: "center",
                                    justifyContent: "space-between",
                                    gap: 10,
                                    marginBottom: 8,
                                  }}
                                >
                                  <div style={{ fontSize: 12, fontWeight: 700, color: "#1e3a8a" }}>
                                    AI suggestion review
                                  </div>
                                  <div style={{ fontSize: 11, color: "#334155" }}>
                                    Status: <strong>{aiReview?.status || "Suggested"}</strong>
                                  </div>
                                </div>
                                <div style={{ fontSize: 12, color: "#1f2937", marginBottom: 6 }}>
                                  Current manual score: <strong>{row.score}</strong> · Suggested score:{" "}
                                  <strong>{aiSuggestion.suggestedScore}</strong>{" "}
                                  <span
                                    style={{
                                      display: "inline-block",
                                      marginLeft: 8,
                                      padding: "2px 8px",
                                      borderRadius: 999,
                                      fontSize: 10,
                                      fontWeight: 800,
                                      letterSpacing: "0.04em",
                                      color:
                                        aiSuggestion.confidence === "High"
                                          ? "#065f46"
                                          : aiSuggestion.confidence === "Low"
                                            ? "#991b1b"
                                            : "#92400e",
                                      background:
                                        aiSuggestion.confidence === "High"
                                          ? "#d1fae5"
                                          : aiSuggestion.confidence === "Low"
                                            ? "#fee2e2"
                                            : "#fef3c7",
                                    }}
                                  >
                                    {aiSuggestion.confidence} confidence
                                  </span>
                                </div>
                                <div style={{ fontSize: 12, color: "#334155", lineHeight: 1.55 }}>
                                  <strong>Evidence summary:</strong> {aiSuggestion.evidenceSummary}
                                </div>
                                <div style={{ marginTop: 6, fontSize: 12, color: "#334155", lineHeight: 1.55 }}>
                                  <strong>Scoring rationale:</strong> {aiSuggestion.scoringRationale}
                                </div>
                                <div style={{ marginTop: 8, display: "grid", gap: 6 }}>
                                  <div style={{ fontSize: 12, color: "#334155" }}>
                                    <strong>Evidence used:</strong> {aiSuggestion.evidenceUsed.join("; ")}
                                  </div>
                                  <div style={{ fontSize: 12, color: "#334155" }}>
                                    <strong>Confidence reason:</strong> {aiSuggestion.confidenceReason}
                                  </div>
                                  <div style={{ fontSize: 12, color: "#334155" }}>
                                    <strong>Evidence gaps:</strong> {aiSuggestion.evidenceGaps.join("; ")}
                                  </div>
                                  <div style={{ fontSize: 12, color: "#334155" }}>
                                    <strong>Key gaps:</strong> {aiSuggestion.keyGaps.join("; ")}
                                  </div>
                                  <div style={{ fontSize: 12, color: "#334155" }}>
                                    <strong>Recommended actions:</strong>{" "}
                                    {aiSuggestion.recommendedActions.join("; ")}
                                  </div>
                                  <div style={{ fontSize: 12, color: "#334155" }}>
                                    <strong>Missing evidence:</strong>{" "}
                                    {aiSuggestion.missingEvidence.join("; ")}
                                  </div>
                                </div>
                                <div style={{ marginTop: 10, display: "flex", gap: 8, flexWrap: "wrap" }}>
                                  <button
                                    type="button"
                                    onClick={() => acceptSuggestionForDomain(d.id)}
                                    style={{
                                      padding: "8px 12px",
                                      borderRadius: 8,
                                      border: "none",
                                      background: "#1d4ed8",
                                      color: "#fff",
                                      fontWeight: 700,
                                      cursor: "pointer",
                                    }}
                                  >
                                    Accept suggestion
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => keepManualScoreForDomain(d.id)}
                                    style={{
                                      padding: "8px 12px",
                                      borderRadius: 8,
                                      border: "1px solid #94a3b8",
                                      background: "#fff",
                                      color: "#334155",
                                      fontWeight: 700,
                                      cursor: "pointer",
                                    }}
                                  >
                                    Keep manual score
                                  </button>
                                </div>
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>

                <div style={{ marginTop: 14, fontSize: 12, color: C.muted }}>
                  Evidence strength affects gap prioritization when scores tie—Weak evidence increases remediation urgency.
                </div>
              </SectionCard>
              </>
              )}

              {activePage === "results" && (
                <ResultsPageContent
                  liveAssessment={liveAssessment}
                  hasGenerated={hasGenerated}
                  maturityBadgeColor={maturityBadgeColor}
                  aiNarratives={aiNarratives}
                  narrativeLoading={narrativeLoading}
                  generateTPRMReport={generateTPRMReport}
                  reportStatus={reportStatus}
                  reportQualityWarnings={reportQualityWarnings}
                  clientDraft={clientDraft}
                />
              )}

              {activePage === "analytics" && <VisualAnalyticsPage liveAssessment={liveAssessment} />}

              {activePage === "library" && (
                <LibraryPageContent libraryRefreshTrigger={memoryLibraryVersion} />
              )}

              {activePage === "roadmap" && <RoadmapPageContent liveAssessment={liveAssessment} />}

              {activePage === "reviewQueue" && (
                <ReviewQueuePageContent
                  profile={profileExtended}
                  liveAssessment={liveAssessment}
                  aiSuggestionsByDomain={aiSuggestionsByDomain}
                  reviewFollowUpByDomain={reviewFollowUpByDomain}
                  onMarkFollowUp={markReviewFollowUp}
                  onClearFollowUp={clearReviewFollowUp}
                  reviewerNotesByDomain={reviewerNotesByDomain}
                  onReviewerNoteChange={onReviewerNoteChange}
                  acceptSuggestionForDomain={acceptSuggestionForDomain}
                  keepManualScoreForDomain={keepManualScoreForDomain}
                />
              )}

              {activePage === "settings" && (
                <SettingsPageContent
                  profileExtended={profileExtended}
                  domainRows={domainRows}
                  liveAssessment={liveAssessment}
                  assessmentType={assessmentType}
                  aiNarratives={aiNarratives}
                />
              )}

              {!NAV_PAGE_IDS.has(activePage) && (
                <div
                  style={{
                    padding: 32,
                    maxWidth: 520,
                    borderRadius: 12,
                    border: `1px solid ${C.border}`,
                    background: C.surface,
                  }}
                >
                  <div style={{ fontSize: 12, fontWeight: 800, color: C.muted, marginBottom: 8 }}>
                    Unrecognized view
                  </div>
                  <p style={{ margin: "0 0 16px", fontSize: 14, lineHeight: 1.5, color: C.text }}>
                    Choose a section from the left navigation: Command Center, Assessment, Results, and more.
                  </p>
                  <button
                    type="button"
                    onClick={() => setActivePage("assessment")}
                    style={{
                      background: C.accent,
                      color: "#fff",
                      border: "none",
                      borderRadius: 8,
                      padding: "10px 16px",
                      fontWeight: 700,
                      fontSize: 13,
                      cursor: "pointer",
                    }}
                  >
                    Open Assessment
                  </button>
                </div>
              )}
            </main>

            {/* Right summary rail — while configuring assessment */}
            {activePage === "assessment" && (
            <aside
              style={{
                width: 318,
                flexShrink: 0,
                borderLeft: `1px solid ${C.border}`,
                background: C.surface,
                overflowY: "auto",
                padding: "20px 20px 32px",
                boxShadow: "-4px 0 24px rgba(15,23,42,0.04)",
              }}
            >
              <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, letterSpacing: "0.08em", marginBottom: 16 }}>
                Results summary
              </div>

              {!hasGenerated ? (
                <div
                  style={{
                    padding: 24,
                    borderRadius: 12,
                    border: `2px dashed ${C.border}`,
                    textAlign: "center",
                    color: C.muted,
                    fontSize: 13,
                    lineHeight: 1.55,
                  }}
                >
                  Run <strong style={{ color: C.text }}>Generate assessment</strong> to unlock this summary and jump to{" "}
                  <strong style={{ color: C.text }}>Results</strong> for detailed metrics and narratives at any time.
                </div>
              ) : (
                <>
                  <div
                    style={{
                      borderRadius: 14,
                      padding: 20,
                      background: `linear-gradient(160deg, #f8fafc 0%, #fff 100%)`,
                      border: `1px solid ${C.border}`,
                      marginBottom: 16,
                    }}
                  >
                    <div style={{ fontSize: 11, color: C.muted, fontWeight: 600 }}>Overall score</div>
                    <div style={{ fontSize: 36, fontWeight: 800, letterSpacing: "-0.03em", marginTop: 4 }}>
                      {liveAssessment.overallScore}
                      <span style={{ fontSize: 16, fontWeight: 600, color: C.muted }}> / 5</span>
                    </div>
                    <div style={{ marginTop: 12, display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                      <span style={{ fontSize: 12, color: C.muted }}>Maturity level</span>
                      <span
                        style={{
                          display: "inline-block",
                          padding: "4px 10px",
                          borderRadius: 999,
                          fontSize: 12,
                          fontWeight: 700,
                          background: `${maturityBadgeColor}18`,
                          color: maturityBadgeColor,
                          border: `1px solid ${maturityBadgeColor}44`,
                        }}
                      >
                        {liveAssessment.maturityBand}
                      </span>
                      <span
                        style={{
                          fontSize: 11,
                          padding: "4px 8px",
                          borderRadius: 6,
                          background: "#f1f5f9",
                          color: C.muted,
                          fontWeight: 600,
                        }}
                      >
                        {liveAssessment.maturityTier} tier
                      </span>
                    </div>
                  </div>

                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns: "1fr 1fr",
                      gap: 10,
                      marginBottom: 16,
                    }}
                  >
                    <MiniStat label="Weak domains" value={String(liveAssessment.weakCount)} hint="Score ≤2" />
                    <MiniStat
                      label="Domains scored"
                      value={`${DOMAINS.length}/${DOMAINS.length}`}
                      hint="Complete"
                    />
                  </div>

                  <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, letterSpacing: "0.06em", marginBottom: 10 }}>
                    Top 3 priority gaps
                  </div>
                  <ul style={{ margin: 0, paddingLeft: 18, fontSize: 13, lineHeight: 1.5 }}>
                    {liveAssessment.top3.map((g) => (
                      <li key={g.domain} style={{ marginBottom: 10 }}>
                        <strong>{g.domain}</strong>
                        <div style={{ color: C.muted, fontSize: 12 }}>
                          {g.score}/5 · {g.evidence}
                        </div>
                      </li>
                    ))}
                  </ul>

                  <div
                    style={{
                      marginTop: 20,
                      paddingTop: 16,
                      borderTop: `1px solid ${C.border}`,
                      fontSize: 12,
                      color: C.muted,
                      lineHeight: 1.5,
                    }}
                  >
                    Open <strong style={{ color: C.text }}>Results</strong> for full narrative and roadmap.
                  </div>
                </>
              )}
            </aside>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

const selectStyle = {
  width: "100%",
  padding: "11px 14px",
  borderRadius: 8,
  border: "1px solid #cbd5e1",
  background: "#fff",
  fontSize: 14,
  lineHeight: 1.45,
};

function SectionCard({ title, subtitle, children }) {
  return (
    <section
      style={{
        background: C.surface,
        borderRadius: 16,
        border: `1px solid ${C.border}`,
        padding: "26px 28px",
        marginBottom: 24,
        boxShadow: "0 4px 24px rgba(15,23,42,0.06)",
      }}
    >
      <h2
        style={{
          margin: 0,
          fontSize: 18,
          fontWeight: 800,
          letterSpacing: "-0.02em",
          color: C.text,
          lineHeight: 1.25,
        }}
      >
        {title}
      </h2>
      {subtitle && (
        <p style={{ margin: "8px 0 20px", fontSize: 14, color: C.muted, lineHeight: 1.55 }}>{subtitle}</p>
      )}
      {!subtitle && <div style={{ height: 6 }} />}
      {children}
    </section>
  );
}

function Field({ label, children }) {
  return (
    <label style={{ display: "block", fontSize: 13, fontWeight: 600, color: "#334155" }}>
      <span style={{ display: "block", marginBottom: 8 }}>{label}</span>
      {children}
    </label>
  );
}

function ScoreBadge({ score }) {
  const tone =
    score <= 2 ? { bg: "#fff7ed", fg: "#9a3412", bd: "#fed7aa" } : score === 3
      ? { bg: "#fffbeb", fg: "#a16207", bd: "#fde68a" }
      : score === 4
        ? { bg: "#eff6ff", fg: "#1d4ed8", bd: "#bfdbfe" }
        : { bg: "#ecfdf5", fg: "#047857", bd: "#a7f3d0" };
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        padding: "6px 12px",
        borderRadius: 999,
        fontSize: 12,
        fontWeight: 800,
        background: tone.bg,
        color: tone.fg,
        border: `1px solid ${tone.bd}`,
        whiteSpace: "nowrap",
      }}
    >
      {score} · {SCORE_LABELS[score]}
    </span>
  );
}

function MiniStat({ label, value, hint }) {
  return (
    <div
      style={{
        borderRadius: 10,
        padding: "12px 14px",
        border: `1px solid ${C.border}`,
        background: "#f8fafc",
      }}
    >
      <div style={{ fontSize: 10, fontWeight: 700, color: C.muted, textTransform: "uppercase", letterSpacing: "0.05em" }}>
        {label}
      </div>
      <div style={{ fontSize: 22, fontWeight: 800, marginTop: 4 }}>{value}</div>
      <div style={{ fontSize: 10, color: C.muted, marginTop: 2 }}>{hint}</div>
    </div>
  );
}
