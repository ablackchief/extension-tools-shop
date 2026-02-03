'use client';

import { useEffect, useState } from 'react';
import { cn } from '@/lib/utils';

interface TOCItem {
  id: string;
  text: string;
}

interface TableOfContentsProps {
  content: string;
}

export function TableOfContents({ content }: TableOfContentsProps) {
  const [headings, setHeadings] = useState<TOCItem[]>([]);
  const [activeId, setActiveId] = useState<string>('');

  // Extract h2 headings from content
  useEffect(() => {
    const h2Regex = /^## (.+)$/gm;
    const matches = [...content.matchAll(h2Regex)];
    const items = matches.map((match) => {
      const text = match[1];
      const id = text.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
      return { id, text };
    });
    setHeadings(items);
  }, [content]);

  // Set up intersection observer
  useEffect(() => {
    if (headings.length === 0) return;

    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            setActiveId(entry.target.id);
          }
        });
      },
      {
        rootMargin: '-100px 0px -66% 0px',
      }
    );

    headings.forEach(({ id }) => {
      const element = document.getElementById(id);
      if (element) {
        observer.observe(element);
      }
    });

    return () => observer.disconnect();
  }, [headings]);

  if (headings.length === 0) {
    return null;
  }

  const scrollToHeading = (id: string) => {
    const element = document.getElementById(id);
    if (element) {
      element.scrollIntoView({ behavior: 'smooth' });
    }
  };

  return (
    <nav className="bg-white rounded-xl p-6 shadow-sm border border-slate-100">
      <h4 className="text-sm font-semibold text-slate-400 uppercase tracking-wider mb-4">
        In This Guide
      </h4>
      <ul className="space-y-2">
        {headings.map(({ id, text }) => (
          <li key={id}>
            <button
              onClick={() => scrollToHeading(id)}
              className={cn(
                'text-left text-sm w-full py-1 px-3 rounded transition-all duration-200',
                activeId === id
                  ? 'text-[var(--primary-600)] font-medium border-l-2 border-[var(--primary-500)] bg-[var(--primary-50)]'
                  : 'text-slate-600 hover:text-slate-900 border-l-2 border-transparent hover:border-slate-300'
              )}
            >
              {text}
            </button>
          </li>
        ))}
      </ul>
    </nav>
  );
}
