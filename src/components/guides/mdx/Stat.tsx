interface StatProps {
  value: string;
  label: string;
}

export function Stat({ value, label }: StatProps) {
  return (
    <div className="my-8 text-center py-6 border-y border-slate-200">
      <div className="text-4xl md:text-5xl font-bold text-[var(--primary-600)] mb-2">
        {value}
      </div>
      <div className="text-slate-600 text-sm md:text-base">{label}</div>
    </div>
  );
}
