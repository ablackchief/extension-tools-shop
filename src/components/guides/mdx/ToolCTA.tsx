import Link from 'next/link';
import {
  Calculator,
  ClipboardCheck,
  FileText,
  Wallet,
  CheckSquare,
  FileEdit,
  MessageSquare,
  Calendar,
  GitBranch,
  Scale,
  Wrench
} from 'lucide-react';
import { getTool } from '@/lib/tool-data';
import { Button } from '@/components/ui/button';

const iconMap: Record<string, React.ElementType> = {
  Calculator,
  ClipboardCheck,
  FileText,
  Wallet,
  CheckSquare,
  FileEdit,
  MessageSquare,
  Calendar,
  GitBranch,
  Scale,
  Wrench,
};

interface ToolCTAProps {
  tool: string;
}

export function ToolCTA({ tool }: ToolCTAProps) {
  const toolData = getTool(tool);

  if (!toolData) {
    return null;
  }

  const Icon = iconMap[toolData.icon] || Wrench;

  return (
    <div className="my-8 bg-gradient-to-r from-[var(--primary-50)] to-blue-50 border-l-4 border-[var(--primary-500)] rounded-r-xl p-6">
      <div className="flex flex-col sm:flex-row sm:items-center gap-4">
        <div className="flex items-center gap-3 flex-1">
          <div className="w-12 h-12 rounded-xl bg-white shadow-sm flex items-center justify-center flex-shrink-0">
            <Icon className="w-6 h-6 text-[var(--primary-600)]" />
          </div>
          <div>
            <h4 className="font-bold text-slate-900">{toolData.name}</h4>
            <p className="text-sm text-slate-600">{toolData.description}</p>
          </div>
        </div>
        <Button asChild variant="orange" size="sm" className="flex-shrink-0">
          <Link href={toolData.link}>
            Get it — {toolData.price}
          </Link>
        </Button>
      </div>
    </div>
  );
}
