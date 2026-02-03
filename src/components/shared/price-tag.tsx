import { cn } from "@/lib/utils";

interface PriceTagProps {
  price: number;
  updatePrice?: number;
  className?: string;
}

export function PriceTag({ price, updatePrice, className }: PriceTagProps) {
  return (
    <div className={cn("flex items-baseline gap-1", className)}>
      <span className="text-4xl font-bold text-slate-900">
        £{price.toFixed(2)}
      </span>
      {updatePrice && (
        <span className="text-slate-500 text-sm">
          + £{updatePrice.toFixed(2)}/year
        </span>
      )}
    </div>
  );
}
