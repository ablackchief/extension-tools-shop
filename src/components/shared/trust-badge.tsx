import { Shield, Award, Building2, LucideIcon } from "lucide-react";
import { cn } from "@/lib/utils";

interface TrustBadgeProps {
  icon: "shield" | "award" | "building";
  text: string;
  className?: string;
}

const icons: Record<string, LucideIcon> = {
  shield: Shield,
  award: Award,
  building: Building2,
};

export function TrustBadge({ icon, text, className }: TrustBadgeProps) {
  const Icon = icons[icon];

  return (
    <div
      className={cn(
        "inline-flex items-center gap-2 px-4 py-2 bg-white/10 backdrop-blur-sm rounded-full text-blue-200 text-sm",
        className
      )}
    >
      <Icon className="w-4 h-4" />
      <span>{text}</span>
    </div>
  );
}
