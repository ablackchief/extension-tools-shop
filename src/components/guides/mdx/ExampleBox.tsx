interface ExampleBoxProps {
  title: string;
  children: React.ReactNode;
}

export function ExampleBox({ title, children }: ExampleBoxProps) {
  return (
    <div className="my-6 bg-slate-50 rounded-xl p-6 border border-slate-200">
      <h4 className="font-semibold text-slate-900 mb-3 flex items-center gap-2">
        <span className="w-2 h-2 bg-[var(--primary-500)] rounded-full" />
        {title}
      </h4>
      <div className="text-slate-600 text-sm [&>p]:m-0 [&>ul]:m-0 [&>ul]:pl-4">
        {children}
      </div>
    </div>
  );
}
