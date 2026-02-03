"use client";

import Image from "next/image";
import { Building2, Award } from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { motion } from "framer-motion";

export function AuthorSection() {
  return (
    <section id="author" className="py-24 bg-white">
      <div className="max-w-7xl mx-auto px-4">
        <div className="grid lg:grid-cols-2 gap-16 items-center">
          <motion.div
            initial={{ opacity: 0, x: -20 }}
            whileInView={{ opacity: 1, x: 0 }}
            viewport={{ once: true }}
          >
            <div className="relative">
              <div className="aspect-[4/5] bg-gradient-to-br from-orange-100 to-orange-200 rounded-2xl shadow-xl overflow-hidden">
                <Image
                  src="/abre-etteh.jpg"
                  alt="Abre Etteh - ARB Registered Architect"
                  fill
                  className="object-cover object-top"
                  sizes="(max-width: 768px) 100vw, 50vw"
                  priority
                />
              </div>
              {/* Decorative elements */}
              <div className="absolute -bottom-6 -right-6 w-32 h-32 bg-orange-100 rounded-2xl -z-10" />
              <div className="absolute -top-6 -left-6 w-24 h-24 bg-blue-50 rounded-2xl -z-10" />
            </div>
          </motion.div>

          <motion.div
            initial={{ opacity: 0, x: 20 }}
            whileInView={{ opacity: 1, x: 0 }}
            viewport={{ once: true }}
          >
            <Badge className="mb-4">About the Author</Badge>
            <h2 className="text-3xl md:text-4xl font-bold text-slate-900">
              Created by Abre Etteh
            </h2>
            <p className="mt-2 text-xl text-blue-600 font-medium">
              ARB Registered Architect
            </p>

            <div className="mt-8 space-y-4 text-lg text-slate-600">
              <p>
                20 years of experience across residential, commercial, and heritage
                sectors.
              </p>
              <p>
                Former Design Review Panel member at Merton Council, where I saw
                firsthand why projects succeed and fail.
              </p>
              <p>
                Developer of the London Small Sites Toolkit, working at the
                intersection of policy and practice.
              </p>
              <p>
                These templates contain everything I wish my clients had known
                before starting their projects.
              </p>
            </div>

            <div className="mt-8 flex flex-wrap items-center gap-6">
              <div className="flex items-center gap-2">
                <Building2 className="w-5 h-5 text-blue-600" />
                <span className="text-slate-600">ARB Registered</span>
              </div>
              <div className="flex items-center gap-2">
                <Award className="w-5 h-5 text-blue-600" />
                <span className="text-slate-600">20 Years Experience</span>
              </div>
            </div>

            <div className="mt-8 p-6 bg-slate-50 rounded-xl border border-slate-200">
              <p className="text-slate-600 italic">
                &ldquo;Most guides are written by people who extended their house once.
                These tools come from someone who&apos;s designed hundreds of
                extensions and seen what actually goes wrong.&rdquo;
              </p>
            </div>
          </motion.div>
        </div>
      </div>
    </section>
  );
}
