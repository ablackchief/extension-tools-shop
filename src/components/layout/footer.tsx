import Link from "next/link";
import { Mail, BookOpen, Linkedin } from "lucide-react";

export function Footer() {
  return (
    <footer className="bg-slate-900 text-white">
      <div className="max-w-7xl mx-auto px-4 py-16">
        <div className="grid md:grid-cols-4 gap-12">
          {/* Brand */}
          <div className="md:col-span-1">
            <div className="flex items-center gap-2 mb-4">
              <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-blue-500 to-blue-600 flex items-center justify-center">
                <span className="text-white font-bold text-sm">ES</span>
              </div>
              <span className="font-semibold">Extension Survival Guide</span>
            </div>
            <p className="text-slate-400 text-sm mb-6">
              Professional planning tools for UK homeowners. Created by an ARB
              registered architect with 20 years of experience.
            </p>
            <div className="flex items-center gap-4">
              <a
                href="mailto:support@extensionsurvivalguide.co.uk"
                className="flex items-center gap-2 text-slate-400 hover:text-white transition-colors"
              >
                <Mail className="w-4 h-4" />
                <span className="text-sm">Contact</span>
              </a>
            </div>
          </div>

          {/* Guides */}
          <div>
            <h4 className="font-semibold mb-4">Guides</h4>
            <ul className="space-y-3">
              <li>
                <Link
                  href="/guides?category=Extension+Costs"
                  className="text-slate-400 hover:text-white transition-colors text-sm"
                >
                  Extension Costs
                </Link>
              </li>
              <li>
                <Link
                  href="/guides?category=Planning+%26+PD"
                  className="text-slate-400 hover:text-white transition-colors text-sm"
                >
                  Planning Permission
                </Link>
              </li>
              <li>
                <Link
                  href="/guides?category=Finding+Builders"
                  className="text-slate-400 hover:text-white transition-colors text-sm"
                >
                  Finding Builders
                </Link>
              </li>
              <li>
                <Link
                  href="/guides?category=Contracts+%26+Payment"
                  className="text-slate-400 hover:text-white transition-colors text-sm"
                >
                  Contracts & Payment
                </Link>
              </li>
              <li>
                <Link
                  href="/guides?category=During+the+Build"
                  className="text-slate-400 hover:text-white transition-colors text-sm"
                >
                  During the Build
                </Link>
              </li>
              <li>
                <Link
                  href="/guides?category=Completion+%26+Defects"
                  className="text-slate-400 hover:text-white transition-colors text-sm"
                >
                  Completion & Defects
                </Link>
              </li>
            </ul>
          </div>

          {/* Quick Links */}
          <div>
            <h4 className="font-semibold mb-4">Quick Links</h4>
            <ul className="space-y-3">
              <li>
                <Link
                  href="/#pricing"
                  className="text-slate-400 hover:text-white transition-colors text-sm"
                >
                  Pricing
                </Link>
              </li>
              <li>
                <Link
                  href="/#tools"
                  className="text-slate-400 hover:text-white transition-colors text-sm"
                >
                  Tool Previews
                </Link>
              </li>
              <li>
                <Link
                  href="/#faq"
                  className="text-slate-400 hover:text-white transition-colors text-sm"
                >
                  FAQ
                </Link>
              </li>
              <li>
                <Link
                  href="/guides"
                  className="text-slate-400 hover:text-white transition-colors text-sm"
                >
                  All Guides
                </Link>
              </li>
            </ul>
          </div>

          {/* About */}
          <div>
            <h4 className="font-semibold mb-4">About</h4>
            <ul className="space-y-3">
              <li>
                <a
                  href="/#author"
                  className="text-slate-400 hover:text-white transition-colors text-sm flex items-center gap-2"
                >
                  <BookOpen className="w-4 h-4" />
                  About the Author
                </a>
              </li>
              <li>
                <a
                  href="https://linkedin.com"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="text-slate-400 hover:text-white transition-colors text-sm flex items-center gap-2"
                >
                  <Linkedin className="w-4 h-4" />
                  LinkedIn
                </a>
              </li>
            </ul>
          </div>
        </div>

        <div className="mt-12 pt-8 border-t border-slate-800 flex flex-col md:flex-row items-center justify-between gap-4">
          <p className="text-slate-500 text-sm">
            © {new Date().getFullYear()} Extension Survival Guide. All rights
            reserved.
          </p>
          <div className="flex items-center gap-6">
            <Link
              href="/privacy"
              className="text-slate-500 hover:text-slate-300 text-sm"
            >
              Privacy Policy
            </Link>
            <Link
              href="/terms"
              className="text-slate-500 hover:text-slate-300 text-sm"
            >
              Terms of Service
            </Link>
          </div>
        </div>
      </div>
    </footer>
  );
}
