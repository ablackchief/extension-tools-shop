import { GuideCard } from './GuideCard';
import type { GuidePost } from '@/lib/guides';

interface RelatedGuidesProps {
  guides: GuidePost[];
}

export function RelatedGuides({ guides }: RelatedGuidesProps) {
  if (guides.length === 0) return null;

  return (
    <section className="bg-slate-50 py-16">
      <div className="max-w-7xl mx-auto px-4">
        <h2 className="text-2xl md:text-3xl font-bold text-slate-900 mb-8 text-center">
          Related Guides
        </h2>
        <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-8">
          {guides.slice(0, 3).map((guide) => (
            <GuideCard key={guide.slug} post={guide} />
          ))}
        </div>
      </div>
    </section>
  );
}
