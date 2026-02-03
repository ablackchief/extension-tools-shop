'use client';

import {
  Accordion,
  AccordionContent,
  AccordionItem,
  AccordionTrigger,
} from '@/components/ui/accordion';
import type { FAQ } from '@/lib/guides';

interface FAQSectionProps {
  faqs: FAQ[];
}

export function FAQSection({ faqs }: FAQSectionProps) {
  if (faqs.length === 0) return null;

  return (
    <section className="bg-slate-50 rounded-xl p-6 md:p-8 my-12">
      <h2 className="text-xl md:text-2xl font-bold text-slate-900 mb-6">
        Frequently Asked Questions
      </h2>
      <Accordion type="single" collapsible className="space-y-3">
        {faqs.map((faq, index) => (
          <AccordionItem
            key={index}
            value={`faq-${index}`}
            className="bg-white rounded-lg border border-slate-200 px-4"
          >
            <AccordionTrigger className="text-left font-medium text-slate-900 py-4 hover:no-underline">
              {faq.question}
            </AccordionTrigger>
            <AccordionContent className="text-slate-600 pb-4">
              {faq.answer}
            </AccordionContent>
          </AccordionItem>
        ))}
      </Accordion>
    </section>
  );
}
