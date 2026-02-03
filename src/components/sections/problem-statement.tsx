"use client";

import { TrendingUp, Clock, AlertTriangle } from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { motion } from "framer-motion";

interface PainPointCardProps {
  icon: React.ReactNode;
  stat: string;
  title: string;
  description: string;
  quote: string;
}

function PainPointCard({ icon, stat, title, description, quote }: PainPointCardProps) {
  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      whileInView={{ opacity: 1, y: 0 }}
      viewport={{ once: true }}
      transition={{ duration: 0.5 }}
      className="bg-white rounded-2xl p-8 shadow-sm border border-slate-200 hover:shadow-lg transition-shadow"
    >
      <div className="flex items-center gap-4 mb-6">
        <div className="w-12 h-12 rounded-xl bg-red-50 flex items-center justify-center">
          {icon}
        </div>
        <div>
          <div className="text-3xl font-bold text-red-600">{stat}</div>
          <div className="text-slate-600 font-medium">{title}</div>
        </div>
      </div>
      <p className="text-slate-600 mb-6">{description}</p>
      <blockquote className="border-l-4 border-red-200 pl-4 text-sm text-slate-500 italic">
        &ldquo;{quote}&rdquo;
      </blockquote>
    </motion.div>
  );
}

export function ProblemStatementSection() {
  return (
    <section className="py-24 bg-slate-50">
      <div className="max-w-7xl mx-auto px-4">
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          className="text-center max-w-3xl mx-auto"
        >
          <Badge variant="outline" className="mb-4">
            The Reality Check
          </Badge>
          <h2 className="text-3xl md:text-4xl font-bold text-slate-900">
            Extensions Go Wrong. A Lot.
          </h2>
          <p className="mt-6 text-lg text-slate-600">
            Based on analysis of thousands of forum posts from UK homeowners...
          </p>
        </motion.div>

        <div className="mt-16 grid md:grid-cols-3 gap-8">
          <PainPointCard
            icon={<TrendingUp className="w-6 h-6 text-red-500" />}
            stat="20-80%"
            title="Over Budget"
            description="Most extensions exceed their budget, often by tens of thousands of pounds."
            quote="We've gone significantly over budget (about £80-90k). We've had to borrow money from family, get personal loans and a credit card."
          />
          <PainPointCard
            icon={<Clock className="w-6 h-6 text-red-500" />}
            stat="2-3x"
            title="Longer Than Quoted"
            description="Timelines routinely double or triple from the original estimate."
            quote="Told 12 weeks, was 28. Builder would disappear for days at a time and nothing happened."
          />
          <PainPointCard
            icon={<AlertTriangle className="w-6 h-6 text-red-500" />}
            stat="50%"
            title="Value Recovery"
            description="Extensions typically add only 50% of their cost to property value."
            quote="I've just had an extension, it cost us £30k and added £5k to the house value."
          />
        </div>
      </div>
    </section>
  );
}
