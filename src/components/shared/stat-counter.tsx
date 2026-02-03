"use client";

import { useEffect, useState, useRef } from "react";
import { cn } from "@/lib/utils";

interface StatCounterProps {
  number: string;
  label: string;
  className?: string;
}

export function StatCounter({ number, label, className }: StatCounterProps) {
  const [displayNumber, setDisplayNumber] = useState("0");
  const ref = useRef<HTMLDivElement>(null);
  const hasAnimated = useRef(false);

  useEffect(() => {
    const observer = new IntersectionObserver(
      (entries) => {
        if (entries[0].isIntersecting && !hasAnimated.current) {
          hasAnimated.current = true;
          animateNumber();
        }
      },
      { threshold: 0.5 }
    );

    if (ref.current) {
      observer.observe(ref.current);
    }

    return () => observer.disconnect();
  }, []);

  const animateNumber = () => {
    const numericPart = number.replace(/[^0-9.]/g, "");
    const prefix = number.match(/^[^0-9]*/)?.[0] || "";
    const suffix = number.match(/[^0-9]*$/)?.[0] || "";
    const targetNumber = parseFloat(numericPart);

    if (isNaN(targetNumber)) {
      setDisplayNumber(number);
      return;
    }

    const duration = 2000;
    const steps = 60;
    const stepDuration = duration / steps;
    let currentStep = 0;

    const interval = setInterval(() => {
      currentStep++;
      const progress = currentStep / steps;
      const easeOut = 1 - Math.pow(1 - progress, 3);
      const currentValue = Math.round(targetNumber * easeOut);

      if (numericPart.includes(".")) {
        setDisplayNumber(`${prefix}${currentValue.toLocaleString()}${suffix}`);
      } else {
        setDisplayNumber(`${prefix}${currentValue.toLocaleString()}${suffix}`);
      }

      if (currentStep >= steps) {
        clearInterval(interval);
        setDisplayNumber(number);
      }
    }, stepDuration);
  };

  return (
    <div ref={ref} className={cn("text-center", className)}>
      <div className="text-3xl md:text-4xl font-bold text-white">
        {displayNumber}
      </div>
      <div className="text-blue-200 text-sm mt-1">{label}</div>
    </div>
  );
}
