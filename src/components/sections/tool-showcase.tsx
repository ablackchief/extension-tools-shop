"use client";

import {
  Calculator,
  PiggyBank,
  GitBranch,
  ClipboardList,
  Award,
  Receipt,
  FileSpreadsheet,
  FileText,
  type LucideIcon,
} from "lucide-react";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Badge } from "@/components/ui/badge";
import { TOOLS } from "@/lib/constants";
import { motion } from "framer-motion";

const TOOL_ICONS: Record<string, { icon: LucideIcon; color: string; bg: string }> = {
  calculator: { icon: Calculator, color: "text-emerald-400", bg: "bg-emerald-500/20" },
  budget: { icon: PiggyBank, color: "text-amber-400", bg: "bg-amber-500/20" },
  "pd-flowchart": { icon: GitBranch, color: "text-violet-400", bg: "bg-violet-500/20" },
  scope: { icon: ClipboardList, color: "text-cyan-400", bg: "bg-cyan-500/20" },
  "builder-scorecard": { icon: Award, color: "text-orange-400", bg: "bg-orange-500/20" },
  "payment-schedule": { icon: Receipt, color: "text-rose-400", bg: "bg-rose-500/20" },
};

interface ToolShowcaseProps {
  toolId: string;
  title: string;
  format: string;
  description: string;
  features: string[];
}

function ToolShowcase({ toolId, title, format, description, features }: ToolShowcaseProps) {
  const isExcel = format.includes("XLSX");
  const toolIcon = TOOL_ICONS[toolId];
  const Icon = toolIcon?.icon || (isExcel ? FileSpreadsheet : FileText);
  const iconColor = toolIcon?.color || (isExcel ? "text-green-400" : "text-blue-400");
  const iconBg = toolIcon?.bg || (isExcel ? "bg-green-500/20" : "bg-blue-500/20");

  return (
    <div className="grid lg:grid-cols-2 gap-12 items-center">
      <div>
        <div className="flex items-center gap-3 mb-4">
          <div
            className={`w-10 h-10 rounded-lg flex items-center justify-center ${iconBg}`}
          >
            <Icon className={`w-5 h-5 ${iconColor}`} />
          </div>
          <Badge variant="outline" className="bg-white/10 border-white/20 text-white">
            {format}
          </Badge>
        </div>

        <h3 className="text-2xl md:text-3xl font-bold text-white mb-4">
          {title}
        </h3>

        <p className="text-slate-300 text-lg mb-8">{description}</p>

        <ul className="space-y-3">
          {features.map((feature, i) => (
            <li key={i} className="flex items-start gap-3">
              <div className="w-5 h-5 rounded-full bg-green-500/20 flex items-center justify-center mt-0.5 flex-shrink-0">
                <div className="w-2 h-2 rounded-full bg-green-400" />
              </div>
              <span className="text-slate-300">{feature}</span>
            </li>
          ))}
        </ul>
      </div>

      <div className="relative">
        <div className="bg-gradient-to-br from-slate-800 to-slate-900 rounded-2xl p-8 shadow-2xl">
          <div className="aspect-[4/3] bg-slate-700/50 rounded-lg flex items-center justify-center">
            <div className="text-center">
              <div className={`w-24 h-24 rounded-2xl ${iconBg} flex items-center justify-center mx-auto mb-4`}>
                <Icon className={`w-12 h-12 ${iconColor}`} />
              </div>
              <p className="text-white font-medium text-lg mb-1">{title}</p>
              <p className="text-slate-400 text-sm">{format}</p>
            </div>
          </div>
        </div>
        {/* Decorative glow */}
        <div className="absolute inset-0 -z-10 bg-gradient-to-r from-blue-500/20 to-purple-500/20 blur-3xl rounded-full" />
      </div>
    </div>
  );
}

export function ToolShowcaseSection() {
  const showcaseTools = TOOLS.slice(0, 6);

  return (
    <section id="tools" className="py-24 bg-slate-900 text-white">
      <div className="max-w-7xl mx-auto px-4">
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          className="text-center max-w-3xl mx-auto mb-16"
        >
          <h2 className="text-3xl md:text-4xl font-bold">
            See What You&apos;re Getting
          </h2>
          <p className="mt-6 text-lg text-slate-400">
            Professional-grade templates with built-in formulas, validation, and
            guidance.
          </p>
        </motion.div>

        <Tabs defaultValue={showcaseTools[0].id} className="w-full">
          <TabsList className="flex flex-wrap justify-center gap-2 mb-12 bg-transparent">
            {showcaseTools.map((tool) => (
              <TabsTrigger
                key={tool.id}
                value={tool.id}
                className="data-[state=active]:bg-white data-[state=active]:text-slate-900"
              >
                {tool.name.split(" ").slice(0, 2).join(" ")}
              </TabsTrigger>
            ))}
          </TabsList>

          {showcaseTools.map((tool) => (
            <TabsContent key={tool.id} value={tool.id}>
              <motion.div
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ duration: 0.3 }}
              >
                <ToolShowcase
                  toolId={tool.id}
                  title={tool.name}
                  format={tool.format === "XLSX" ? "Excel (XLSX)" : "Word (DOCX)"}
                  description={tool.description}
                  features={tool.features}
                />
              </motion.div>
            </TabsContent>
          ))}
        </Tabs>
      </div>
    </section>
  );
}
