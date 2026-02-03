'use client';

import { useState, useMemo } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { Award } from 'lucide-react';
import { Header } from '@/components/layout/header';
import { Footer } from '@/components/layout/footer';
import {
  GuideCard,
  CategoryFilter,
  Pagination,
  EmailCapture,
} from '@/components/guides';
import type { GuidePost } from '@/lib/guides';

const POSTS_PER_PAGE = 12;

interface GuidesPageClientProps {
  initialGuides: GuidePost[];
}

export function GuidesPageClient({ initialGuides }: GuidesPageClientProps) {
  const [activeCategory, setActiveCategory] = useState('all');
  const [currentPage, setCurrentPage] = useState(1);

  // Filter guides by category
  const filteredGuides = useMemo(() => {
    if (activeCategory === 'all') return initialGuides;
    return initialGuides.filter((guide) => guide.category === activeCategory);
  }, [initialGuides, activeCategory]);

  // Reset to page 1 when category changes
  const handleCategoryChange = (category: string) => {
    setActiveCategory(category);
    setCurrentPage(1);
  };

  // Paginate
  const totalPages = Math.ceil(filteredGuides.length / POSTS_PER_PAGE);
  const startIndex = (currentPage - 1) * POSTS_PER_PAGE;
  const paginatedGuides = filteredGuides.slice(startIndex, startIndex + POSTS_PER_PAGE);

  // Separate featured (first) from rest
  const featuredGuide = currentPage === 1 && activeCategory === 'all' ? paginatedGuides[0] : null;
  const gridGuides = featuredGuide ? paginatedGuides.slice(1) : paginatedGuides;

  return (
    <div className="min-h-screen bg-white">
      <Header />

      {/* Hero Section */}
      <section className="bg-gradient-to-br from-slate-900 via-blue-900 to-slate-900 pt-32 pb-16 md:pt-40 md:pb-20">
        <div className="max-w-7xl mx-auto px-4 text-center">
          <h1 className="text-3xl md:text-5xl font-bold text-white mb-4">
            Extension Guides
          </h1>
          <p className="text-lg md:text-xl text-slate-300 max-w-2xl mx-auto mb-8">
            Free, honest advice from an ARB-registered architect. No fluff. No sales
            pitches. Just the information that prevents the predictable disasters.
          </p>
          <div className="inline-flex items-center gap-2 px-4 py-2 bg-white/10 rounded-full text-sm text-slate-300">
            <Award className="w-4 h-4 text-[var(--gold-400)]" />
            Written by Abre Etteh, ARB Registered Architect | 20 Years Experience
          </div>
        </div>
      </section>

      {/* Category Filter */}
      <CategoryFilter
        activeCategory={activeCategory}
        onCategoryChange={handleCategoryChange}
      />

      {/* Content */}
      <section className="py-12 md:py-16">
        <div className="max-w-7xl mx-auto px-4">
          <AnimatePresence mode="wait">
            <motion.div
              key={`${activeCategory}-${currentPage}`}
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              transition={{ duration: 0.2 }}
            >
              {filteredGuides.length === 0 ? (
                <div className="text-center py-16">
                  <p className="text-slate-500 text-lg">
                    No guides found in this category yet.
                  </p>
                </div>
              ) : (
                <>
                  {/* Featured Post */}
                  {featuredGuide && (
                    <div className="mb-12">
                      <GuideCard post={featuredGuide} featured />
                    </div>
                  )}

                  {/* Grid */}
                  {gridGuides.length > 0 && (
                    <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-8">
                      {gridGuides.map((guide) => (
                        <GuideCard key={guide.slug} post={guide} />
                      ))}
                    </div>
                  )}
                </>
              )}
            </motion.div>
          </AnimatePresence>

          {/* Pagination */}
          <Pagination
            currentPage={currentPage}
            totalPages={totalPages}
            onPageChange={setCurrentPage}
          />
        </div>
      </section>

      {/* Email Capture */}
      <EmailCapture />

      <Footer />
    </div>
  );
}
