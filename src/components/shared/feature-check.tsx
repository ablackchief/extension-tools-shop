import { Check } from "lucide-react";
import { cn } from "@/lib/utils";

interface FeatureCheckProps {
  children: React.ReactNode;
  format?: 'XLSX' | 'DOCX';
  description?: string;
  included?: boolean;
  badge?: string;
  className?: string;
}

export function FeatureCheck({
  children,
  format,
  description,
  included,
  badge,
  className,
}: FeatureCheckProps) {
  return (
    <div className={cn("flex items-start gap-3", className)}>
      <div className="mt-0.5 flex-shrink-0">
        <div className="w-5 h-5 rounded-full bg-green-100 flex items-center justify-center">
          <Check className="w-3 h-3 text-green-600" />
        </div>
      </div>
      <div className="flex-1">
        <div className="flex items-center gap-2">
          <span className={cn(
            "font-medium text-slate-900",
            included && "text-slate-600"
          )}>
            {children}
          </span>
          {format && (
            <span className="px-2 py-0.5 text-xs font-medium rounded bg-slate-100 text-slate-600">
              {format}
            </span>
          )}
          {badge && (
            <span className="px-2 py-0.5 text-xs font-semibold rounded bg-gradient-to-r from-green-500 to-emerald-500 text-white">
              {badge}
            </span>
          )}
        </div>
        {description && (
          <p className="text-sm text-slate-500 mt-0.5">{description}</p>
        )}
      </div>
    </div>
  );
}
