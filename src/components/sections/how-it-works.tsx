"use client";

import { Package, Download, FileSpreadsheet } from "lucide-react";
import { motion } from "framer-motion";

interface StepProps {
  number: string;
  title: string;
  description: string;
  icon: React.ReactNode;
}

function Step({ number, title, description, icon }: StepProps) {
  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      whileInView={{ opacity: 1, y: 0 }}
      viewport={{ once: true }}
      className="text-center"
    >
      <div className="relative inline-block mb-6">
        <div className="w-20 h-20 rounded-2xl bg-gradient-to-br from-blue-500 to-blue-600 flex items-center justify-center shadow-lg">
          {icon}
        </div>
        <div className="absolute -top-2 -right-2 w-8 h-8 rounded-full bg-orange-500 flex items-center justify-center text-white font-bold text-sm shadow">
          {number}
        </div>
      </div>
      <h3 className="text-xl font-bold text-slate-900 mb-3">{title}</h3>
      <p className="text-slate-600 max-w-xs mx-auto">{description}</p>
    </motion.div>
  );
}

export function HowItWorksSection() {
  return (
    <section className="py-24 bg-white">
      <div className="max-w-7xl mx-auto px-4">
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          className="text-center max-w-3xl mx-auto mb-16"
        >
          <h2 className="text-3xl md:text-4xl font-bold text-slate-900">
            Ready in 60 Seconds
          </h2>
        </motion.div>

        <div className="grid md:grid-cols-3 gap-12">
          <Step
            number="1"
            title="Choose Your Bundle"
            description="Select the toolkit that matches your project stage. Not sure? Start with Complete Toolkit."
            icon={<Package className="w-10 h-10 text-white" />}
          />
          <Step
            number="2"
            title="Instant Download"
            description="Secure payment via Stripe. Download link delivered immediately to your email."
            icon={<Download className="w-10 h-10 text-white" />}
          />
          <Step
            number="3"
            title="Open and Use"
            description="Works with Excel, Google Sheets, Word, and Google Docs. Start planning today."
            icon={<FileSpreadsheet className="w-10 h-10 text-white" />}
          />
        </div>

        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          className="mt-16 text-center"
        >
          <p className="text-slate-500">
            Compatible with: Microsoft Excel 2016+, Google Sheets, Microsoft Word
            2016+, Google Docs
          </p>
        </motion.div>
      </div>
    </section>
  );
}
