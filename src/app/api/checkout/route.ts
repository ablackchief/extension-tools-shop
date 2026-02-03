import { NextResponse } from 'next/server';
import { STRIPE_PRODUCTS } from '@/lib/stripe';
import { BundleType } from '@/types';

export async function POST(req: Request) {
  try {
    const key = process.env.STRIPE_SECRET_KEY?.trim();
    if (!key) {
      return NextResponse.json({ error: 'Stripe key not configured' }, { status: 500 });
    }

    const { bundle, email } = await req.json() as { bundle: BundleType; email?: string };

    if (!STRIPE_PRODUCTS[bundle]) {
      return NextResponse.json({ error: 'Invalid bundle' }, { status: 400 });
    }

    const product = STRIPE_PRODUCTS[bundle];
    const baseUrl = 'https://extensionsurvivalguide.co.uk';

    const params = new URLSearchParams();
    params.append('payment_method_types[]', 'card');
    params.append('mode', 'payment');
    params.append('success_url', `${baseUrl}/success?session_id={CHECKOUT_SESSION_ID}`);
    params.append('cancel_url', `${baseUrl}/#pricing`);
    params.append('metadata[bundle]', bundle);

    // Main product
    params.append('line_items[0][price_data][currency]', 'gbp');
    params.append('line_items[0][price_data][product_data][name]', product.name);
    params.append('line_items[0][price_data][product_data][description]', product.description);
    params.append('line_items[0][price_data][unit_amount]', String(product.price));
    params.append('line_items[0][quantity]', '1');

    if (email) {
      params.append('customer_email', email);
    }

    const response = await fetch('https://api.stripe.com/v1/checkout/sessions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${key}`,
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: params.toString(),
    });

    const session = await response.json();

    if (!response.ok) {
      console.error('Stripe API error:', session);
      return NextResponse.json(
        { error: 'Failed to create checkout session' },
        { status: 500 }
      );
    }

    return NextResponse.json({ url: session.url });
  } catch (error) {
    console.error('Checkout error:', error);
    return NextResponse.json(
      { error: 'Failed to create checkout session' },
      { status: 500 }
    );
  }
}
