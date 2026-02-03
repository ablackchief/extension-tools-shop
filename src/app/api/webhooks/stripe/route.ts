import { NextResponse } from 'next/server';
import { stripe } from '@/lib/stripe';
import { generateDownloadToken } from '@/lib/tokens';
import { sendDownloadEmail, sendOrderNotification } from '@/lib/email';
import { BundleType } from '@/types';

export async function POST(req: Request) {
  const body = await req.text();
  const sig = req.headers.get('stripe-signature');

  if (!sig) {
    return NextResponse.json({ error: 'Missing signature' }, { status: 400 });
  }

  let event;

  try {
    event = stripe.webhooks.constructEvent(
      body,
      sig,
      (process.env.STRIPE_WEBHOOK_SECRET || '').trim()
    );
  } catch (err: any) {
    console.error('Webhook signature verification failed:', err.message);
    return NextResponse.json(
      { error: 'Webhook signature verification failed' },
      { status: 400 }
    );
  }

  if (event.type === 'checkout.session.completed') {
    const session = event.data.object as any;
    const bundle = session.metadata?.bundle as BundleType | undefined;
    const email = session.customer_email || session.customer_details?.email;
    const amountTotal = session.amount_total || 0;

    if (bundle && email) {
      try {
        // Generate secure download token (expires in 30 days)
        const token = generateDownloadToken(bundle, email);

        // Send email with download link to customer
        await sendDownloadEmail(email, bundle, token);

        // Send order notification to business
        await sendOrderNotification(email, bundle, amountTotal);

        console.log(`Download email sent to ${email} for bundle ${bundle}`);
      } catch (error) {
        console.error('Failed to process successful payment:', error);
      }
    }
  }

  return NextResponse.json({ received: true });
}
