'use client';

import { cn } from '@/lib/utils';

const categories = [
  { label: 'All', value: 'all' },
  { label: 'Extension Costs', value: 'Extension Costs' },
  { label: 'Planning & PD', value: 'Planning & PD' },
  { label: 'Finding Builders', value: 'Finding Builders' },
  { label: 'Contracts & Payment', value: 'Contracts & Payment' },
  { label: 'During the Build', value: 'During the Build' },
  { label: 'Completion & Defects', value: 'Completion & Defects' },
];

interface CategoryFilterProps {
  activeCategory: string;
  onCategoryChange: (category: string) => void;
}

export function CategoryFilter({ activeCategory, onCategoryChange }: CategoryFilterProps) {
  return (
    <div className="sticky top-16 md:top-20 z-40 bg-white border-b border-slate-200 shadow-sm">
      <div className="max-w-7xl mx-auto px-4 py-4">
        <div className="flex items-center gap-2 overflow-x-auto pb-2 -mb-2 scrollbar-hide">
          {categories.map((category) => (
            <button
              key={category.value}
              onClick={() => onCategoryChange(category.value)}
              className={cn(
                'px-4 py-2 rounded-full text-sm font-medium whitespace-nowrap transition-all duration-200',
                activeCategory === category.value
                  ? 'bg-[var(--primary-600)] text-white shadow-md'
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              )}
            >
              {category.label}
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}
