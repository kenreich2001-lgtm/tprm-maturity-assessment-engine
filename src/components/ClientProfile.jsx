import SectionCard from "./SectionCard.jsx";
import SelectField from "./SelectField.jsx";

const OPTIONS = {
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
  assessmentType: [
    "Full maturity assessment",
    "Rapid diagnostic",
    "Targeted domain review",
  ],
};

export default function ClientProfile({ profile, onChange }) {
  const set = (key) => (v) => onChange({ ...profile, [key]: v });

  return (
    <SectionCard
      title="Client profile"
      subtitle="Context drives interpretation of maturity scores and remediation sequencing."
    >
      <div className="grid gap-6 sm:grid-cols-2 lg:grid-cols-3">
        <SelectField
          label="Industry"
          id="industry"
          value={profile.industry}
          onChange={set("industry")}
          options={OPTIONS.industry}
        />
        <SelectField
          label="Company size"
          id="companySize"
          value={profile.companySize}
          onChange={set("companySize")}
          options={OPTIONS.companySize}
        />
        <SelectField
          label="Regulatory intensity"
          id="regulatoryIntensity"
          value={profile.regulatoryIntensity}
          onChange={set("regulatoryIntensity")}
          options={OPTIONS.regulatoryIntensity}
        />
        <SelectField
          label="Third-party volume"
          id="thirdPartyVolume"
          value={profile.thirdPartyVolume}
          onChange={set("thirdPartyVolume")}
          options={OPTIONS.thirdPartyVolume}
        />
        <SelectField
          label="Geographic footprint"
          id="geographicFootprint"
          value={profile.geographicFootprint}
          onChange={set("geographicFootprint")}
          options={OPTIONS.geographicFootprint}
        />
        <SelectField
          label="Assessment type"
          id="assessmentType"
          value={profile.assessmentType}
          onChange={set("assessmentType")}
          options={OPTIONS.assessmentType}
          hint="Determines depth of recommendations and roadmap emphasis."
        />
      </div>
    </SectionCard>
  );
}
