"use client";

import { Button } from "@/components/ui/button";
import { motion } from "framer-motion";

export function FinalCTASection() {
  const scrollToSection = (id: string) => {
    const element = document.getElementById(id);
    if (element) {
      element.scrollIntoView({ behavior: "smooth" });
    }
  };

  return (
    <section className="py-24 bg-gradient-to-br from-blue-600 via-blue-700 to-purple-800 text-white">
      <div className="max-w-4xl mx-auto px-4 text-center">
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
        >
          <h2 className="text-3xl md:text-4xl lg:text-5xl font-bold">
            Your Extension is a Major Investment
          </h2>
          <p className="mt-6 text-xl text-blue-100">
            The average UK extension costs £50,000-£150,000. These tools cost less
            than one hour of architect fees and could save you thousands.
          </p>

          <div className="mt-12 flex flex-col sm:flex-row gap-6 justify-center">
            <Button
              size="xl"
              className="bg-white text-blue-700 hover:bg-blue-50 px-10 py-6 text-lg font-semibold"
              onClick={() => scrollToSection("pricing")}
            >
              Get Complete Toolkit - £39.99
            </Button>
            <Button
              size="xl"
              variant="outline"
              className="border-white text-white hover:bg-white/10 px-10 py-6 text-lg"
              onClick={() => scrollToSection("pricing")}
            >
              View All Bundles
            </Button>
          </div>

          <p className="mt-8 text-blue-200 text-sm">
            Instant download • 30-day money-back guarantee • Works with Excel &
            Word
          </p>
        </motion.div>
      </div>
    </section>
  );
}
