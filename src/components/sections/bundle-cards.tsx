"use client";

import { useState } from "react";
import { ShieldCheck, Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { FeatureCheck } from "@/components/shared/feature-check";
import { PriceTag } from "@/components/shared/price-tag";
import { BUNDLES } from "@/lib/constants";
import { BundleType } from "@/types";
import { motion } from "framer-motion";
import { cn } from "@/lib/utils";

export function BundleCardsSection() {
  const [loadingBundle, setLoadingBundle] = useState<BundleType | null>(null);
  const [error, setError] = useState<string | null>(null);

  const handleCheckout = async (bundleId: BundleType) => {
    setLoadingBundle(bundleId);
    setError(null);

    try {
      const response = await fetch("/api/checkout", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ bundle: bundleId }),
      });

      const data = await response.json();

      if (data.url) {
        window.location.href = data.url;
      } else {
        setError("Something went wrong. Please try again.");
        setLoadingBundle(null);
      }
    } catch (error) {
      setError("Connection error. Please try again.");
      setLoadingBundle(null);
    }
  };

  return (
    <section id="pricing" className="py-24 bg-white">
      <div className="max-w-7xl mx-auto px-4">
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          className="text-center max-w-3xl mx-auto mb-16"
        >
          <Badge className="bg-blue-100 text-blue-700 mb-4">
            Professional Tools
          </Badge>
          <h2 className="text-3xl md:text-4xl font-bold text-slate-900">
            The Same Templates Architects Use
          </h2>
          <p className="mt-6 text-lg text-slate-600">
            Instant download. Works with Excel and Word. No subscription for core tools.
          </p>
        </motion.div>

        <div className="grid lg:grid-cols-3 gap-8">
          {BUNDLES.map((bundle, index) => (
            <motion.div
              key={bundle.id}
              initial={{ opacity: 0, y: 20 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true }}
              transition={{ delay: index * 0.1 }}
            >
              <Card
                className={cn(
                  "relative h-full flex flex-col",
                  bundle.popular &&
                    "border-2 border-orange-500 shadow-xl scale-105 z-10"
                )}
              >
                {bundle.popular && (
                  <div className="absolute -top-4 left-1/2 -translate-x-1/2">
                    <Badge variant="popular">Most Popular</Badge>
                  </div>
                )}

                <CardHeader className="pb-4">
                  <h3 className="text-xl font-bold text-slate-900">
                    {bundle.name}
                  </h3>
                  <p className="text-slate-500 text-sm">{bundle.description}</p>
                  <div className="pt-4">
                    <PriceTag
                      price={bundle.price}
                      updatePrice={bundle.updatePrice}
                    />
                  </div>
                </CardHeader>

                <CardContent className="flex-1 flex flex-col">
                  <div className="space-y-4 flex-1">
                    {bundle.features.map((feature, i) => (
                      <FeatureCheck
                        key={i}
                        format={feature.format}
                        description={feature.description}
                        included={feature.included}
                        badge={feature.badge}
                      >
                        {feature.name}
                      </FeatureCheck>
                    ))}
                  </div>

                  <Button
                    className="w-full mt-8"
                    size="lg"
                    variant={bundle.popular ? "orange" : "default"}
                    onClick={() => handleCheckout(bundle.id)}
                    disabled={loadingBundle !== null}
                  >
                    {loadingBundle === bundle.id ? (
                      <>
                        <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                        Processing...
                      </>
                    ) : (
                      `Get ${bundle.name}`
                    )}
                  </Button>
                  {error && loadingBundle === null && (
                    <p className="text-red-500 text-sm text-center mt-2">{error}</p>
                  )}
                </CardContent>
              </Card>
            </motion.div>
          ))}
        </div>

        {/* Guarantee */}
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          className="mt-12 text-center"
        >
          <div className="inline-flex items-center gap-3 px-6 py-3 bg-green-50 rounded-full">
            <ShieldCheck className="w-5 h-5 text-green-600" />
            <span className="text-green-800 font-medium">
              30-Day Money Back Guarantee - No Questions Asked
            </span>
          </div>
        </motion.div>

        {/* Individual template note */}
        <motion.p
          initial={{ opacity: 0 }}
          whileInView={{ opacity: 1 }}
          viewport={{ once: true }}
          className="mt-6 text-center text-sm text-slate-500"
        >
          Looking for an individual template?{" "}
          <a
            href="mailto:support@extensionsurvivalguide.co.uk?subject=Individual%20Template%20Enquiry"
            className="text-[var(--primary-600)] hover:underline"
          >
            Send us an email
          </a>{" "}
          and we&apos;d be happy to help.
        </motion.p>
      </div>
    </section>
  );
}
