import * as React from "react";
import { cva, type VariantProps } from "class-variance-authority";
import { cn } from "@/lib/utils";

const badgeVariants = cva(
  "inline-flex items-center rounded-full px-3 py-1 text-xs font-semibold transition-colors",
  {
    variants: {
      variant: {
        default:
          "bg-blue-100 text-blue-700",
        secondary:
          "bg-slate-100 text-slate-700",
        success:
          "bg-green-100 text-green-700",
        warning:
          "bg-amber-100 text-amber-700",
        destructive:
          "bg-red-100 text-red-700",
        outline:
          "border border-slate-300 text-slate-700 bg-white",
        popular:
          "bg-gradient-to-r from-orange-500 to-orange-600 text-white shadow-sm",
        new:
          "bg-gradient-to-r from-green-500 to-emerald-500 text-white shadow-sm",
      },
    },
    defaultVariants: {
      variant: "default",
    },
  }
);

export interface BadgeProps
  extends React.HTMLAttributes<HTMLDivElement>,
    VariantProps<typeof badgeVariants> {}

function Badge({ className, variant, ...props }: BadgeProps) {
  return (
    <div className={cn(badgeVariants({ variant }), className)} {...props} />
  );
}

export { Badge, badgeVariants };
