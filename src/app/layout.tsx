import type { Metadata } from "next";
import { Inter } from "next/font/google";
import "./globals.css";

const inter = Inter({
  subsets: ["latin"],
  variable: "--font-geist-sans",
});

export const metadata: Metadata = {
  metadataBase: new URL("https://extensionsurvivalguide.co.uk"),
  title: "Extension Survival Guide - Professional Planning Tools for UK Homeowners",
  description:
    "Professional Excel and Word templates used by architects to protect homeowners from budget blowouts, timeline disasters, and builder nightmares. Created by an ARB Registered Architect.",
  keywords: [
    "extension planning",
    "home extension",
    "budget planner",
    "builder vetting",
    "UK homeowner",
    "extension calculator",
    "planning permission",
    "snagging checklist",
  ],
  authors: [{ name: "Abre Etteh", url: "https://extensionsurvivalguide.co.uk" }],
  openGraph: {
    type: "website",
    locale: "en_GB",
    url: "https://extensionsurvivalguide.co.uk",
    title: "Extension Survival Guide - Professional Planning Tools",
    description:
      "Professional Excel and Word templates to plan your extension without the disasters.",
    siteName: "Extension Survival Guide",
    images: [
      {
        url: "/og-image.jpg",
        width: 1200,
        height: 630,
        alt: "Extension Survival Guide Tools",
      },
    ],
  },
  twitter: {
    card: "summary_large_image",
    title: "Extension Survival Guide",
    description:
      "Professional Excel and Word templates to plan your extension without the disasters.",
    images: ["/og-image.jpg"],
  },
  robots: {
    index: true,
    follow: true,
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" className="scroll-smooth">
      <body className={`${inter.variable} antialiased`}>{children}</body>
    </html>
  );
}
