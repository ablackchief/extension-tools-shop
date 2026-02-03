import Link from 'next/link';
import { ArrowRight, Award } from 'lucide-react';

interface AuthorBoxProps {
  name: string;
  credential: string;
  bio: string;
}

export function AuthorBox({ name, credential, bio }: AuthorBoxProps) {
  return (
    <div className="bg-white rounded-xl shadow-md p-6 md:p-8">
      <div className="flex flex-col sm:flex-row gap-6">
        {/* Avatar */}
        <div className="flex-shrink-0">
          <div className="w-20 h-20 rounded-full bg-gradient-to-br from-[var(--primary-500)] to-[var(--primary-700)] flex items-center justify-center">
            <span className="text-2xl font-bold text-white">
              {name.split(' ').map(n => n[0]).join('')}
            </span>
          </div>
        </div>

        {/* Content */}
        <div className="flex-1">
          <div className="flex items-center gap-2 mb-1">
            <h3 className="font-bold text-lg text-slate-900">{name}</h3>
            <Award className="w-4 h-4 text-[var(--gold-500)]" />
          </div>
          <p className="text-sm font-medium text-[var(--primary-600)] mb-3">
            {credential}
          </p>
          <p className="text-slate-600 text-sm mb-4">{bio}</p>
          <Link
            href="/guides"
            className="inline-flex items-center gap-1 text-sm font-medium text-[var(--primary-600)] hover:text-[var(--primary-700)] transition-colors"
          >
            View All Guides
            <ArrowRight className="w-4 h-4" />
          </Link>
        </div>
      </div>
    </div>
  );
}
