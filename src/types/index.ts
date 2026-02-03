export type BundleType = 'essential' | 'complete' | 'professional';

export interface Bundle {
  id: BundleType;
  name: string;
  price: number;
  priceInPence: number;
  description: string;
  features: BundleFeature[];
  popular?: boolean;
  hasSubscription?: boolean;
  updatePrice?: number;
}

export interface BundleFeature {
  name: string;
  format?: 'XLSX' | 'DOCX';
  description?: string;
  included?: boolean;
  badge?: string;
}

export interface Tool {
  id: string;
  name: string;
  format: 'XLSX' | 'DOCX';
  description: string;
  features: string[];
  bundle: BundleType[];
  generator: string;
  filename: string;
}

export interface DownloadToken {
  bundle: BundleType;
  email: string;
  createdAt: number;
  expiresAt: number;
}

export interface CheckoutRequest {
  bundle: BundleType;
  email?: string;
}

export interface Testimonial {
  quote: string;
  author: string;
  location: string;
  rating: number;
  project: string;
}

export interface FAQ {
  question: string;
  answer: string;
}
