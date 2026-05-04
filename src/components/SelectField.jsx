export default function SelectField({
  label,
  id,
  value,
  onChange,
  options,
  hint,
}) {
  return (
    <div className="flex flex-col gap-1.5">
      <label
        htmlFor={id}
        className="text-xs font-medium uppercase tracking-wide text-slate-500"
      >
        {label}
      </label>
      <select
        id={id}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className="rounded-lg border border-slate-200 bg-white px-3 py-2.5 text-sm text-slate-800 shadow-sm outline-none ring-accent/10 transition focus:border-accent focus:ring-2 focus:ring-accent/20"
      >
        {options.map((opt) => (
          <option key={opt} value={opt}>
            {opt}
          </option>
        ))}
      </select>
      {hint && <p className="text-xs text-slate-400">{hint}</p>}
    </div>
  );
}
