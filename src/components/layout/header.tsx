"use client";

import { useState, useEffect } from "react";
import Link from "next/link";
import { usePathname } from "next/navigation";
import { Menu, X } from "lucide-react";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

export function Header() {
  const [isScrolled, setIsScrolled] = useState(false);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const pathname = usePathname();

  const isGuidesPage = pathname?.startsWith("/guides");
  const isHomePage = pathname === "/";

  useEffect(() => {
    const handleScroll = () => {
      setIsScrolled(window.scrollY > 10);
    };

    window.addEventListener("scroll", handleScroll);
    return () => window.removeEventListener("scroll", handleScroll);
  }, []);

  const scrollToSection = (id: string) => {
    setIsMobileMenuOpen(false);
    if (isHomePage) {
      const element = document.getElementById(id);
      if (element) {
        element.scrollIntoView({ behavior: "smooth" });
      }
    } else {
      window.location.href = `/#${id}`;
    }
  };

  const navLinkClasses = (isActive?: boolean) =>
    cn(
      "text-sm font-medium transition-colors hover:text-blue-500",
      isActive
        ? "text-blue-600"
        : isScrolled
        ? "text-slate-600"
        : "text-white/80 hover:text-white"
    );

  return (
    <header
      className={cn(
        "fixed top-0 left-0 right-0 z-50 transition-all duration-300",
        isScrolled
          ? "bg-white/95 backdrop-blur-md shadow-sm"
          : "bg-transparent"
      )}
    >
      <div className="max-w-7xl mx-auto px-4">
        <div className="flex items-center justify-between h-16 md:h-20">
          <Link href="/" className="flex items-center gap-2">
            <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-blue-600 to-blue-700 flex items-center justify-center">
              <span className="text-white font-bold text-sm">ES</span>
            </div>
            <span
              className={cn(
                "font-semibold hidden sm:block transition-colors",
                isScrolled ? "text-slate-900" : "text-white"
              )}
            >
              Extension Survival Guide
            </span>
          </Link>

          <nav className="hidden md:flex items-center gap-8">
            <Link href="/guides" className={navLinkClasses(isGuidesPage)}>
              Guides
            </Link>
            <button
              onClick={() => scrollToSection("tools")}
              className={navLinkClasses()}
            >
              Tools
            </button>
            <button
              onClick={() => scrollToSection("pricing")}
              className={navLinkClasses()}
            >
              Pricing
            </button>
            <button
              onClick={() => scrollToSection("faq")}
              className={navLinkClasses()}
            >
              FAQ
            </button>
          </nav>

          <div className="flex items-center gap-4">
            <Button
              variant="orange"
              size="sm"
              onClick={() => scrollToSection("pricing")}
              className="hidden md:flex"
            >
              Get Started
            </Button>

            <button
              className={cn(
                "md:hidden p-2 rounded-lg transition-colors",
                isScrolled
                  ? "text-slate-600 hover:bg-slate-100"
                  : "text-white hover:bg-white/10"
              )}
              onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)}
            >
              {isMobileMenuOpen ? (
                <X className="w-6 h-6" />
              ) : (
                <Menu className="w-6 h-6" />
              )}
            </button>
          </div>
        </div>
      </div>

      {/* Mobile menu */}
      {isMobileMenuOpen && (
        <div className="md:hidden bg-white border-t border-slate-200 shadow-lg">
          <nav className="flex flex-col p-4 gap-2">
            <Link
              href="/guides"
              onClick={() => setIsMobileMenuOpen(false)}
              className={cn(
                "px-4 py-3 text-left rounded-lg",
                isGuidesPage
                  ? "bg-blue-50 text-blue-600"
                  : "text-slate-600 hover:bg-slate-50"
              )}
            >
              Guides
            </Link>
            <button
              onClick={() => scrollToSection("tools")}
              className="px-4 py-3 text-left text-slate-600 hover:bg-slate-50 rounded-lg"
            >
              Tools
            </button>
            <button
              onClick={() => scrollToSection("pricing")}
              className="px-4 py-3 text-left text-slate-600 hover:bg-slate-50 rounded-lg"
            >
              Pricing
            </button>
            <button
              onClick={() => scrollToSection("faq")}
              className="px-4 py-3 text-left text-slate-600 hover:bg-slate-50 rounded-lg"
            >
              FAQ
            </button>
            <Button
              variant="orange"
              className="mt-2"
              onClick={() => scrollToSection("pricing")}
            >
              Get Started
            </Button>
          </nav>
        </div>
      )}
    </header>
  );
}
