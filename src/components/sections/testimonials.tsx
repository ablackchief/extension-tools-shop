"use client";

import { Star } from "lucide-react";
import { Card } from "@/components/ui/card";
import { TESTIMONIALS } from "@/lib/constants";
import { motion } from "framer-motion";

interface TestimonialCardProps {
  quote: string;
  author: string;
  location: string;
  rating: number;
  project: string;
}

function TestimonialCard({
  quote,
  author,
  location,
  rating,
  project,
}: TestimonialCardProps) {
  return (
    <Card className="p-8 h-full flex flex-col">
      <div className="flex gap-1 mb-4">
        {Array.from({ length: rating }).map((_, i) => (
          <Star key={i} className="w-5 h-5 fill-amber-400 text-amber-400" />
        ))}
      </div>

      <blockquote className="text-slate-700 flex-1 mb-6">
        &ldquo;{quote}&rdquo;
      </blockquote>

      <div className="flex items-center gap-4">
        <div className="w-12 h-12 rounded-full bg-gradient-to-br from-blue-500 to-blue-600 flex items-center justify-center text-white font-bold">
          {author.charAt(0)}
        </div>
        <div>
          <div className="font-semibold text-slate-900">{author}</div>
          <div className="text-sm text-slate-500">
            {location} • {project}
          </div>
        </div>
      </div>
    </Card>
  );
}

export function TestimonialsSection() {
  return (
    <section id="testimonials" className="py-24 bg-slate-50">
      <div className="max-w-7xl mx-auto px-4">
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          className="text-center max-w-3xl mx-auto mb-16"
        >
          <h2 className="text-3xl md:text-4xl font-bold text-slate-900">
            Trusted by UK Homeowners
          </h2>
        </motion.div>

        <div className="grid md:grid-cols-3 gap-8">
          {TESTIMONIALS.map((testimonial, index) => (
            <motion.div
              key={index}
              initial={{ opacity: 0, y: 20 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true }}
              transition={{ delay: index * 0.1 }}
            >
              <TestimonialCard {...testimonial} />
            </motion.div>
          ))}
        </div>
      </div>
    </section>
  );
}
