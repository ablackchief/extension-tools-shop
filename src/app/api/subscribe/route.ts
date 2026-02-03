import { NextRequest, NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';

export async function POST(request: NextRequest) {
  try {
    const { email, source } = await request.json();

    if (!email || !email.includes('@')) {
      return NextResponse.json(
        { error: 'Please provide a valid email address' },
        { status: 400 }
      );
    }

    // For now, store in a JSON file. In production, use Resend or another service.
    const subscribersPath = path.join(process.cwd(), 'subscribers.json');

    let subscribers: { email: string; source: string; date: string }[] = [];

    if (fs.existsSync(subscribersPath)) {
      const data = fs.readFileSync(subscribersPath, 'utf8');
      subscribers = JSON.parse(data);
    }

    // Check for duplicate
    if (subscribers.some((s) => s.email.toLowerCase() === email.toLowerCase())) {
      return NextResponse.json(
        { error: 'This email is already subscribed' },
        { status: 400 }
      );
    }

    subscribers.push({
      email,
      source: source || 'unknown',
      date: new Date().toISOString(),
    });

    fs.writeFileSync(subscribersPath, JSON.stringify(subscribers, null, 2));

    return NextResponse.json({ success: true });
  } catch (error) {
    console.error('Subscribe error:', error);
    return NextResponse.json(
      { error: 'Something went wrong. Please try again.' },
      { status: 500 }
    );
  }
}
