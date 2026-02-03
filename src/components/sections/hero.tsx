"use client";

import { ArrowRight, ChevronDown, Shield } from "lucide-react";
import { Button } from "@/components/ui/button";
import { StatCounter } from "@/components/shared/stat-counter";
import { motion } from "framer-motion";

export function HeroSection() {
  const scrollToSection = (id: string) => {
    const element = document.getElementById(id);
    if (element) {
      element.scrollIntoView({ behavior: "smooth" });
    }
  };

  return (
    <section className="relative min-h-screen bg-gradient-to-br from-slate-900 via-blue-900 to-slate-900 overflow-hidden">
      {/* Animated grid pattern overlay */}
      <div
        className="absolute inset-0 opacity-10"
        style={{
          backgroundImage: `url("data:image/svg+xml,%3Csvg width='60' height='60' viewBox='0 0 60 60' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='none' fill-rule='evenodd'%3E%3Cg fill='%23ffffff' fill-opacity='0.4'%3E%3Cpath d='M36 34v-4h-2v4h-4v2h4v4h2v-4h4v-2h-4zm0-30V0h-2v4h-4v2h4v4h2V6h4V4h-4zM6 34v-4H4v4H0v2h4v4h2v-4h4v-2H6zM6 4V0H4v4H0v2h4v4h2V6h4V4H6z'/%3E%3C/g%3E%3C/g%3E%3C/svg%3E")`,
        }}
      />

      {/* Floating tool previews */}
      <motion.div
        initial={{ opacity: 0, x: 50, rotate: 6 }}
        animate={{ opacity: 0.8, x: 0, rotate: 6 }}
        transition={{ duration: 1, delay: 0.5 }}
        className="absolute top-32 right-10 w-48 md:w-64 hidden lg:block"
      >
        <div className="bg-white rounded-lg shadow-2xl p-4 transform hover:scale-105 transition-transform">
          <div className="h-32 bg-gradient-to-br from-blue-50 to-blue-100 rounded flex items-center justify-center">
            <span className="text-blue-600 font-medium text-sm">Budget Planner</span>
          </div>
        </div>
      </motion.div>

      <motion.div
        initial={{ opacity: 0, x: 50, rotate: -3 }}
        animate={{ opacity: 0.7, x: 0, rotate: -3 }}
        transition={{ duration: 1, delay: 0.7 }}
        className="absolute bottom-48 right-24 w-40 md:w-56 hidden lg:block"
      >
        <div className="bg-white rounded-lg shadow-2xl p-4 transform hover:scale-105 transition-transform">
          <div className="h-28 bg-gradient-to-br from-orange-50 to-orange-100 rounded flex items-center justify-center">
            <span className="text-orange-600 font-medium text-sm">Decision Calculator</span>
          </div>
        </div>
      </motion.div>

      <div className="relative max-w-7xl mx-auto px-4 pt-32 pb-20">
        <div className="max-w-3xl">
          {/* Trust badge */}
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6 }}
            className="inline-flex items-center gap-2 px-4 py-2 bg-white/10 backdrop-blur-sm rounded-full text-blue-200 text-sm mb-8"
          >
            <Shield className="w-4 h-4" />
            Created by an ARB Registered Architect with 20 Years Experience
          </motion.div>

          <motion.h1
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6, delay: 0.1 }}
            className="text-4xl sm:text-5xl md:text-6xl lg:text-7xl font-bold text-white leading-tight"
          >
            Plan Your Extension
            <span className="block text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-cyan-300">
              Without the Disasters
            </span>
          </motion.h1>

          <motion.p
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6, delay: 0.2 }}
            className="mt-8 text-lg md:text-xl text-slate-300 leading-relaxed max-w-2xl"
          >
            Professional Excel and Word templates used by architects to protect
            homeowners from budget blowouts, timeline disasters, and builder nightmares.
            Download instantly. Use forever.
          </motion.p>

{/* Stats - Temporarily disabled
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6, delay: 0.3 }}
            className="mt-12 flex flex-wrap gap-8 md:gap-12"
          >
            <StatCounter number="2,500+" label="Templates Downloaded" />
            <StatCounter number="£47M+" label="Projects Planned" />
            <StatCounter number="20" label="Years Experience" />
          </motion.div>
*/}

          {/* CTA */}
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6, delay: 0.4 }}
            className="mt-12 flex flex-col sm:flex-row gap-4"
          >
            <Button
              size="xl"
              variant="orange"
              onClick={() => scrollToSection("pricing")}
            >
              Get the Complete Toolkit - £39.99
              <ArrowRight className="ml-2 w-5 h-5" />
            </Button>
            <Button
              size="xl"
              variant="outline"
              className="border-white/30 text-white hover:bg-white/10"
              onClick={() => scrollToSection("tools")}
            >
              See What&apos;s Included
            </Button>
          </motion.div>
        </div>
      </div>

      {/* Scroll indicator */}
      <motion.div
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        transition={{ duration: 0.6, delay: 1 }}
        className="absolute bottom-8 left-1/2 -translate-x-1/2"
      >
        <motion.div
          animate={{ y: [0, 10, 0] }}
          transition={{ repeat: Infinity, duration: 2 }}
        >
          <ChevronDown className="w-8 h-8 text-white/50" />
        </motion.div>
      </motion.div>
    </section>
  );
}
