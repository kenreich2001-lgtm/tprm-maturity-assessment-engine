export default function ScoreBadge({ score }) {
  const n = Number(score);
  const palette =
    n <= 2
      ? "bg-amber-50 text-amber-900 ring-amber-200"
      : n === 3
        ? "bg-blue-50 text-blue-900 ring-blue-200"
        : "bg-emerald-50 text-emerald-900 ring-emerald-200";
  return (
    <span
      className={`inline-flex items-center rounded-md px-2.5 py-0.5 text-sm font-semibold ring-1 ring-inset ${palette}`}
    >
      {n}/5
    </span>
  );
}
