import type { Metadata } from 'next';
import { notFound } from 'next/navigation';
import { MDXRemote } from 'next-mdx-remote/rsc';
import { Header } from '@/components/layout/header';
import { Footer } from '@/components/layout/footer';
import {
  Breadcrumbs,
  AuthorBox,
  RelatedGuides,
  EmailCapture,
  FAQSection,
  TableOfContents,
  SidebarToolCTA,
} from '@/components/guides';
import { mdxComponents } from '@/components/guides/mdx/MDXComponents';
import { getAllGuides, getGuideBySlug, getRelatedGuides } from '@/lib/guides';
import { Clock, Calendar, Award } from 'lucide-react';
import Image from 'next/image';
import { FileText } from 'lucide-react';

interface Props {
  params: Promise<{ slug: string }>;
}

export async function generateStaticParams() {
  const guides = getAllGuides();
  return guides.map((guide) => ({ slug: guide.slug }));
}

export async function generateMetadata({ params }: Props): Promise<Metadata> {
  const { slug } = await params;
  const post = getGuideBySlug(slug);

  if (!post) {
    return { title: 'Guide Not Found' };
  }

  return {
    title: `${post.title} | Extension Survival Guide`,
    description: post.description,
    authors: [{ name: post.author }],
    openGraph: {
      title: post.title,
      description: post.description,
      type: 'article',
      publishedTime: post.date,
      modifiedTime: post.updated,
      authors: [post.author],
      images: post.featuredImage
        ? [{ url: post.featuredImage, alt: post.featuredImageAlt }]
        : undefined,
    },
    alternates: {
      canonical: `https://extensionsurvivalguide.co.uk/guides/${post.slug}/`,
    },
  };
}

function formatDate(dateString: string): string {
  const date = new Date(dateString);
  return date.toLocaleDateString('en-GB', {
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  });
}

export default async function GuidePage({ params }: Props) {
  const { slug } = await params;
  const post = getGuideBySlug(slug);

  if (!post) {
    notFound();
  }

  const relatedGuides = getRelatedGuides(post.relatedPosts);

  // Schema.org Article structured data
  const articleSchema = {
    '@context': 'https://schema.org',
    '@type': 'Article',
    headline: post.title,
    description: post.description,
    image: post.featuredImage || undefined,
    author: {
      '@type': 'Person',
      name: post.author,
      jobTitle: post.authorCredential,
      description: post.authorBio,
      url: 'https://extensionsurvivalguide.co.uk/about/',
    },
    publisher: {
      '@type': 'Organization',
      name: 'The Extension Survival Guide',
      url: 'https://extensionsurvivalguide.co.uk',
    },
    datePublished: post.date,
    dateModified: post.updated,
    mainEntityOfPage: {
      '@type': 'WebPage',
      '@id': `https://extensionsurvivalguide.co.uk/guides/${post.slug}/`,
    },
  };

  // FAQ Schema if FAQs exist
  const faqSchema =
    post.faqs.length > 0
      ? {
          '@context': 'https://schema.org',
          '@type': 'FAQPage',
          mainEntity: post.faqs.map((faq) => ({
            '@type': 'Question',
            name: faq.question,
            acceptedAnswer: {
              '@type': 'Answer',
              text: faq.answer,
            },
          })),
        }
      : null;

  return (
    <div className="min-h-screen bg-white">
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{ __html: JSON.stringify(articleSchema) }}
      />
      {faqSchema && (
        <script
          type="application/ld+json"
          dangerouslySetInnerHTML={{ __html: JSON.stringify(faqSchema) }}
        />
      )}

      <Header />

      <main className="pt-24 md:pt-28">
        {/* Breadcrumbs */}
        <div className="max-w-7xl mx-auto px-4 py-4">
          <Breadcrumbs
            items={[
              { label: 'Guides', href: '/guides' },
              { label: post.category, href: `/guides?category=${encodeURIComponent(post.category)}` },
              { label: post.title },
            ]}
          />
        </div>

        {/* Article Header */}
        <header className="max-w-3xl mx-auto px-4 text-center mb-8">
          <span className="inline-block px-3 py-1 rounded-full text-xs font-medium bg-[var(--primary-50)] text-[var(--primary-700)] mb-4">
            {post.category}
          </span>
          <h1 className="text-3xl md:text-4xl lg:text-5xl font-bold text-slate-900 mb-4">
            {post.title}
          </h1>
          <p className="text-lg md:text-xl text-slate-600 mb-6">
            {post.description}
          </p>

          {/* Author Row */}
          <div className="flex items-center justify-center gap-4 text-sm text-slate-500">
            <div className="flex items-center gap-2">
              <div className="w-8 h-8 rounded-full bg-gradient-to-br from-[var(--primary-500)] to-[var(--primary-700)] flex items-center justify-center">
                <span className="text-xs font-bold text-white">AE</span>
              </div>
              <span className="font-medium text-slate-700">{post.author}</span>
              <Award className="w-4 h-4 text-[var(--gold-500)]" />
            </div>
            <span className="text-slate-300">·</span>
            <span className="flex items-center gap-1">
              <Calendar className="w-4 h-4" />
              {formatDate(post.date)}
            </span>
            <span className="text-slate-300">·</span>
            <span className="flex items-center gap-1">
              <Clock className="w-4 h-4" />
              {post.readingTime} min read
            </span>
          </div>
        </header>

        {/* Featured Image */}
        <div className="max-w-5xl mx-auto px-4 mb-12">
          <div className="relative aspect-[21/9] rounded-xl overflow-hidden">
            {post.featuredImage ? (
              <Image
                src={post.featuredImage}
                alt={post.featuredImageAlt}
                fill
                className="object-cover"
                priority
              />
            ) : (
              <div className="absolute inset-0 bg-gradient-to-br from-[var(--primary-100)] via-[var(--primary-50)] to-blue-100 flex items-center justify-center">
                <FileText className="w-24 h-24 text-[var(--primary-300)]" />
              </div>
            )}
          </div>
        </div>

        {/* Content with Sidebar */}
        <div className="max-w-7xl mx-auto px-4">
          <div className="lg:grid lg:grid-cols-[1fr_280px] lg:gap-12">
            {/* Sidebar - Desktop only */}
            <aside className="hidden lg:block order-2">
              <div className="sticky top-28 space-y-6">
                <TableOfContents content={post.content} />
                {post.toolCTA && post.toolCTA.name && (
                  <SidebarToolCTA tool={post.toolCTA} />
                )}
              </div>
            </aside>

            {/* Article Content */}
            <article className="order-1 max-w-3xl">
              <div className="prose prose-lg prose-slate max-w-none">
                <MDXRemote source={post.content} components={mdxComponents} />
              </div>

              {/* FAQ Section */}
              <FAQSection faqs={post.faqs} />

              {/* Author Box */}
              <div className="mt-12">
                <AuthorBox
                  name={post.author}
                  credential={post.authorCredential}
                  bio={post.authorBio}
                />
              </div>
            </article>
          </div>
        </div>

        {/* Related Guides */}
        <div className="mt-16">
          <RelatedGuides guides={relatedGuides} />
        </div>

        {/* Email Capture */}
        <EmailCapture />
      </main>

      <Footer />
    </div>
  );
}
