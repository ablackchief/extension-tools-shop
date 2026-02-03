import Link from 'next/link';
import { ArrowRight } from 'lucide-react';
import type { ToolCTA } from '@/lib/guides';

interface SidebarToolCTAProps {
  tool: ToolCTA;
}

export function SidebarToolCTA({ tool }: SidebarToolCTAProps) {
  return (
    <div className="bg-gradient-to-br from-[var(--primary-50)] to-blue-50 rounded-xl p-5 border border-[var(--primary-100)]">
      <h4 className="font-bold text-slate-900 mb-2">{tool.name}</h4>
      <p className="text-sm text-slate-600 mb-4">{tool.description}</p>
      <Link
        href={tool.link}
        className="inline-flex items-center gap-1 text-sm font-semibold text-[var(--primary-600)] hover:text-[var(--primary-700)] transition-colors"
      >
        Get the {tool.bundle} bundle
        <ArrowRight className="w-4 h-4" />
      </Link>
    </div>
  );
}
