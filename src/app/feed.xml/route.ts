import { getAllGuides } from '@/lib/guides';

export async function GET() {
  const guides = getAllGuides();

  const feedItems = guides
    .map(
      (guide) => `
    <item>
      <title><![CDATA[${guide.title}]]></title>
      <description><![CDATA[${guide.description}]]></description>
      <link>https://extensionsurvivalguide.co.uk/guides/${guide.slug}/</link>
      <guid isPermaLink="true">https://extensionsurvivalguide.co.uk/guides/${guide.slug}/</guid>
      <pubDate>${new Date(guide.date).toUTCString()}</pubDate>
      <author>support@extensionsurvivalguide.co.uk (${guide.author})</author>
      <category>${guide.category}</category>
    </item>`
    )
    .join('');

  const feed = `<?xml version="1.0" encoding="UTF-8"?>
<rss version="2.0" xmlns:atom="http://www.w3.org/2005/Atom">
  <channel>
    <title>The Extension Survival Guide — Guides</title>
    <description>Free extension guides from an ARB-registered architect. Real data on costs, planning, builders, and more.</description>
    <link>https://extensionsurvivalguide.co.uk/guides/</link>
    <atom:link href="https://extensionsurvivalguide.co.uk/feed.xml" rel="self" type="application/rss+xml"/>
    <language>en-GB</language>
    <lastBuildDate>${new Date().toUTCString()}</lastBuildDate>
    <managingEditor>support@extensionsurvivalguide.co.uk (Abre Etteh)</managingEditor>
    <webMaster>support@extensionsurvivalguide.co.uk (Abre Etteh)</webMaster>
    ${feedItems}
  </channel>
</rss>`;

  return new Response(feed, {
    headers: {
      'Content-Type': 'application/xml',
      'Cache-Control': 'max-age=0, s-maxage=3600',
    },
  });
}
