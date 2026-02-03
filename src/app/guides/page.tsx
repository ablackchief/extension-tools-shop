import type { Metadata } from 'next';
import { GuidesPageClient } from './GuidesPageClient';
import { getAllGuides } from '@/lib/guides';

export const metadata: Metadata = {
  title: 'Extension Guides | The Extension Survival Guide',
  description:
    'Free guides from an ARB-registered architect covering extension costs, planning permission, builder selection, contracts, and completion. Real data. No fluff.',
  openGraph: {
    title: 'Extension Guides | The Extension Survival Guide',
    description:
      'Free guides from an ARB-registered architect. Real data on costs, planning, builders, and more.',
    type: 'website',
    url: 'https://extensionsurvivalguide.co.uk/guides/',
  },
  alternates: {
    canonical: 'https://extensionsurvivalguide.co.uk/guides/',
  },
};

export default function GuidesPage() {
  const guides = getAllGuides();

  // Schema.org CollectionPage structured data
  const schemaData = {
    '@context': 'https://schema.org',
    '@type': 'CollectionPage',
    name: 'Extension Guides',
    description:
      'Free guides from an ARB-registered architect covering extension costs, planning permission, builder selection, contracts, and completion.',
    url: 'https://extensionsurvivalguide.co.uk/guides/',
    publisher: {
      '@type': 'Organization',
      name: 'The Extension Survival Guide',
    },
  };

  return (
    <>
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{ __html: JSON.stringify(schemaData) }}
      />
      <GuidesPageClient initialGuides={guides} />
    </>
  );
}
