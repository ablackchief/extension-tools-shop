import Stripe from 'stripe';

// Lazy initialization to avoid build-time errors when env vars aren't set
let stripeInstance: Stripe | null = null;
export function getStripe(): Stripe {
  if (!stripeInstance) {
    stripeInstance = new Stripe((process.env.STRIPE_SECRET_KEY || '').trim());
  }
  return stripeInstance;
}

// Keep for backwards compatibility but prefer getStripe()
export const stripe = {
  get checkout() { return getStripe().checkout; },
  get webhooks() { return getStripe().webhooks; },
};

export const STRIPE_PRODUCTS = {
  essential: {
    name: 'Essential Tools Bundle',
    price: 1999,
    description: 'Extension Calculator, Budget Planner, PD Flowchart',
  },
  complete: {
    name: 'Complete Toolkit Bundle',
    price: 3999,
    description: 'All 7 core tools for your extension project',
  },
  professional: {
    name: 'Professional Plus Bundle',
    price: 5999,
    description: 'All 10 tools plus one free update',
  },
} as const;
