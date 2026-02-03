import Link from 'next/link';
import Image from 'next/image';
import { Clock, FileText } from 'lucide-react';
import { cn } from '@/lib/utils';
import type { GuidePost } from '@/lib/guides';

interface GuideCardProps {
  post: GuidePost;
  featured?: boolean;
}

function formatDate(dateString: string): string {
  const date = new Date(dateString);
  return date.toLocaleDateString('en-GB', {
    day: 'numeric',
    month: 'short',
    year: 'numeric',
  });
}

export function GuideCard({ post, featured = false }: GuideCardProps) {
  if (featured) {
    return (
      <Link
        href={`/guides/${post.slug}`}
        className="group block bg-white rounded-2xl shadow-lg overflow-hidden hover:shadow-xl transition-all duration-300"
      >
        <div className="grid md:grid-cols-5 gap-0">
          {/* Image */}
          <div className="md:col-span-2 relative aspect-video md:aspect-auto">
            {post.featuredImage ? (
              <Image
                src={post.featuredImage}
                alt={post.featuredImageAlt}
                fill
                className="object-cover"
              />
            ) : (
              <div className="absolute inset-0 bg-gradient-to-br from-[var(--primary-100)] via-[var(--primary-50)] to-blue-100 flex items-center justify-center">
                <FileText className="w-16 h-16 text-[var(--primary-300)]" />
              </div>
            )}
          </div>

          {/* Content */}
          <div className="md:col-span-3 p-6 md:p-8 flex flex-col justify-center">
            <span className="inline-block px-3 py-1 rounded-full text-xs font-medium bg-[var(--primary-50)] text-[var(--primary-700)] w-fit mb-3">
              {post.category}
            </span>
            <h2 className="text-xl md:text-2xl font-bold text-slate-900 group-hover:text-[var(--primary-600)] transition-colors mb-3 line-clamp-2">
              {post.title}
            </h2>
            <p className="text-slate-600 mb-4 line-clamp-2 md:line-clamp-3">
              {post.description}
            </p>
            <div className="flex items-center gap-4 text-sm text-slate-400">
              <span>{formatDate(post.date)}</span>
              <span className="flex items-center gap-1">
                <Clock className="w-4 h-4" />
                {post.readingTime} min read
              </span>
            </div>
          </div>
        </div>
      </Link>
    );
  }

  return (
    <Link
      href={`/guides/${post.slug}`}
      className="group block bg-white rounded-xl shadow-md hover:shadow-xl transition-all duration-300 hover:-translate-y-1 overflow-hidden"
    >
      {/* Image */}
      <div className="relative aspect-video">
        {post.featuredImage ? (
          <Image
            src={post.featuredImage}
            alt={post.featuredImageAlt}
            fill
            className="object-cover"
          />
        ) : (
          <div className="absolute inset-0 bg-gradient-to-br from-[var(--primary-100)] via-[var(--primary-50)] to-blue-100 flex items-center justify-center">
            <FileText className="w-12 h-12 text-[var(--primary-300)]" />
          </div>
        )}
      </div>

      {/* Content */}
      <div className="p-5">
        <span className="inline-block px-3 py-1 rounded-full text-xs font-medium bg-[var(--primary-50)] text-[var(--primary-700)] mb-3">
          {post.category}
        </span>
        <h3 className="text-lg font-bold text-slate-900 group-hover:text-[var(--primary-600)] transition-colors mb-2 line-clamp-2">
          {post.title}
        </h3>
        <p className="text-sm text-slate-600 mb-4 line-clamp-3">
          {post.description}
        </p>
        <div className="flex items-center justify-between text-xs text-slate-400">
          <span>{formatDate(post.date)}</span>
          <span className="flex items-center gap-1">
            <Clock className="w-3 h-3" />
            {post.readingTime} min read
          </span>
        </div>
      </div>
    </Link>
  );
}
