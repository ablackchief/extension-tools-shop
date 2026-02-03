import Link from 'next/link';
import { ChevronRight, Home } from 'lucide-react';

interface BreadcrumbItem {
  label: string;
  href?: string;
}

interface BreadcrumbsProps {
  items: BreadcrumbItem[];
}

export function Breadcrumbs({ items }: BreadcrumbsProps) {
  return (
    <nav className="flex items-center gap-2 text-sm text-slate-400 flex-wrap">
      <Link
        href="/"
        className="hover:text-slate-600 transition-colors flex items-center gap-1"
      >
        <Home className="w-3 h-3" />
        <span>Home</span>
      </Link>
      {items.map((item, index) => (
        <span key={index} className="flex items-center gap-2">
          <ChevronRight className="w-3 h-3" />
          {item.href ? (
            <Link href={item.href} className="hover:text-slate-600 transition-colors">
              {item.label}
            </Link>
          ) : (
            <span className="text-slate-600 line-clamp-1">{item.label}</span>
          )}
        </span>
      ))}
    </nav>
  );
}
