export default function SectionCard({ title, subtitle, children, className = "" }) {
  return (
    <section
      className={`rounded-xl border border-slate-200/80 bg-white shadow-sm shadow-slate-900/5 ${className}`}
    >
      {(title || subtitle) && (
        <div className="border-b border-slate-100 px-6 py-4">
          {title && (
            <h2 className="text-lg font-semibold tracking-tight text-navy-900">
              {title}
            </h2>
          )}
          {subtitle && (
            <p className="mt-1 text-sm text-slate-500">{subtitle}</p>
          )}
        </div>
      )}
      <div className="p-6">{children}</div>
    </section>
  );
}
