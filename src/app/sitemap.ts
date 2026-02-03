import { MetadataRoute } from 'next';
import { getAllGuides } from '@/lib/guides';

export default function sitemap(): MetadataRoute.Sitemap {
  const guides = getAllGuides();

  const guidePages = guides.map((guide) => ({
    url: `https://extensionsurvivalguide.co.uk/guides/${guide.slug}/`,
    lastModified: new Date(guide.updated),
    changeFrequency: 'monthly' as const,
    priority: 0.8,
  }));

  return [
    {
      url: 'https://extensionsurvivalguide.co.uk/',
      lastModified: new Date(),
      changeFrequency: 'monthly',
      priority: 1.0,
    },
    {
      url: 'https://extensionsurvivalguide.co.uk/guides/',
      lastModified: new Date(),
      changeFrequency: 'weekly',
      priority: 0.9,
    },
    ...guidePages,
  ];
}
