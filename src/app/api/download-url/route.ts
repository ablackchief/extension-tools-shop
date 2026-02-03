import { NextResponse } from 'next/server';
import { generateDownloadToken } from '@/lib/tokens';
import { BundleType } from '@/types';

export async function GET(req: Request) {
  const { searchParams } = new URL(req.url);
  const sessionId = searchParams.get('session_id');

  if (!sessionId) {
    return NextResponse.json({ error: 'Missing session_id' }, { status: 400 });
  }

  const key = process.env.STRIPE_SECRET_KEY?.trim();
  if (!key) {
    return NextResponse.json({ error: 'Stripe key not configured' }, { status: 500 });
  }

  try {
    const response = await fetch(
      `https://api.stripe.com/v1/checkout/sessions/${encodeURIComponent(sessionId)}`,
      {
        headers: {
          'Authorization': `Bearer ${key}`,
        },
      }
    );

    if (!response.ok) {
      console.error('Stripe API error:', await response.json());
      return NextResponse.json({ error: 'Failed to retrieve session' }, { status: 500 });
    }

    const session = await response.json();

    if (session.payment_status !== 'paid') {
      return NextResponse.json({ error: 'Payment not completed' }, { status: 400 });
    }

    const bundle = session.metadata?.bundle as BundleType | undefined;
    const email = session.customer_email || session.customer_details?.email;

    if (!bundle) {
      return NextResponse.json({ error: 'Bundle not found' }, { status: 400 });
    }

    // Generate download token
    const token = generateDownloadToken(bundle, email || 'unknown');
    const baseUrl = 'https://extensionsurvivalguide.co.uk';

    return NextResponse.json({
      downloadUrl: `${baseUrl}/download/${token}`,
      email,
      bundle,
    });
  } catch (error) {
    console.error('Error retrieving session:', error);
    return NextResponse.json(
      { error: 'Failed to retrieve session' },
      { status: 500 }
    );
  }
}
