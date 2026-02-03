import { Header } from "@/components/layout/header";
import { Footer } from "@/components/layout/footer";
import { HeroSection } from "@/components/sections/hero";
import { ProblemStatementSection } from "@/components/sections/problem-statement";
import { BundleCardsSection } from "@/components/sections/bundle-cards";
import { ToolShowcaseSection } from "@/components/sections/tool-showcase";
import { HowItWorksSection } from "@/components/sections/how-it-works";
// import { TestimonialsSection } from "@/components/sections/testimonials"; // Temporarily disabled
import { AuthorSection } from "@/components/sections/author-section";
import { FAQSection } from "@/components/sections/faq";
import { FinalCTASection } from "@/components/sections/final-cta";

export default function HomePage() {
  return (
    <main className="min-h-screen">
      <Header />
      <HeroSection />
      <ProblemStatementSection />
      <BundleCardsSection />
      <ToolShowcaseSection />
      <HowItWorksSection />
      {/* <TestimonialsSection /> - Temporarily disabled */}
      <AuthorSection />
      <FAQSection />
      <FinalCTASection />
      <Footer />
    </main>
  );
}
