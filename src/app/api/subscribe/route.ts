import { NextRequest, NextResponse } from 'next/server';
import { Resend } from 'resend';

const resend = new Resend(process.env.RESEND_API_KEY);
const BUSINESS_EMAIL = process.env.BUSINESS_EMAIL || 'support@extensionsurvivalguide.co.uk';
const BASE_URL = 'https://extensionsurvivalguide.co.uk';

export async function POST(request: NextRequest) {
  try {
    const { email, source } = await request.json();

    if (!email || !email.includes('@')) {
      return NextResponse.json(
        { error: 'Please provide a valid email address' },
        { status: 400 }
      );
    }

    // Send checklist email to subscriber
    await resend.emails.send({
      from: `Extension Survival Guide <${BUSINESS_EMAIL}>`,
      to: email,
      subject: 'Your Extension Planning Checklist',
      html: `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #1e293b; max-width: 600px; margin: 0 auto; padding: 20px;">
  <div style="text-align: center; margin-bottom: 30px;">
    <h1 style="color: #1a56f5; margin: 0;">Extension Survival Guide</h1>
    <p style="color: #64748b; margin: 5px 0 0;">Professional Extension Planning Tools</p>
  </div>

  <div style="background: linear-gradient(135deg, #eef5ff 0%, #dbeafe 100%); border-radius: 12px; padding: 30px; margin-bottom: 30px;">
    <h2 style="margin: 0 0 15px; color: #1e293b;">Your Extension Planning Checklist</h2>
    <p style="margin: 0; color: #475569;">Here's the one-page checklist covering the 15 steps most homeowners skip when planning an extension.</p>
  </div>

  <div style="text-align: center; margin: 30px 0;">
    <a href="${BASE_URL}/Extension-Planning-Checklist.pdf" style="display: inline-block; background: linear-gradient(135deg, #f97316 0%, #ea580c 100%); color: white; text-decoration: none; padding: 16px 32px; border-radius: 8px; font-weight: 600; font-size: 16px;">
      Download Your Checklist (PDF)
    </a>
  </div>

  <div style="background: #f8fafc; border-radius: 8px; padding: 20px; margin-bottom: 30px;">
    <h3 style="margin: 0 0 15px; color: #1e293b;">What's inside:</h3>
    <ul style="margin: 0; padding-left: 20px; color: #475569;">
      <li>15 critical planning steps in order</li>
      <li>Common mistakes to avoid at each stage</li>
      <li>Questions to ask before you start</li>
      <li>Timeline guidance for each phase</li>
    </ul>
  </div>

  <div style="background: #f0fdf4; border-left: 4px solid #22c55e; padding: 15px; margin-bottom: 30px;">
    <p style="margin: 0; color: #166534;">
      <strong>Next step:</strong> Once you've reviewed the checklist, explore our professional planning tools to help you execute each step properly.
    </p>
    <p style="margin: 10px 0 0;">
      <a href="${BASE_URL}/#pricing" style="color: #1a56f5; text-decoration: underline;">View the Extension Planning Tools →</a>
    </p>
  </div>

  <div style="margin-bottom: 30px;">
    <h3 style="color: #1e293b;">Free Guides:</h3>
    <p style="color: #475569;">We also publish free guides on extension planning. Recent articles:</p>
    <ul style="color: #475569; padding-left: 20px;">
      <li><a href="${BASE_URL}/guides/extension-over-budget-what-to-do" style="color: #1a56f5;">Extension Over Budget: What to Do Right Now</a></li>
      <li><a href="${BASE_URL}/guides/hidden-costs-house-extension" style="color: #1a56f5;">Hidden Costs of a House Extension</a></li>
      <li><a href="${BASE_URL}/guides/extend-or-move-house-calculator" style="color: #1a56f5;">Extend or Move? How to Decide</a></li>
    </ul>
  </div>

  <div style="border-top: 1px solid #e2e8f0; padding-top: 20px; text-align: center; color: #64748b; font-size: 14px;">
    <p>Questions? Reply to this email or contact us at:<br>
    <a href="mailto:${BUSINESS_EMAIL}" style="color: #1a56f5;">${BUSINESS_EMAIL}</a></p>

    <p style="margin-top: 20px;">
      <strong>Abre Etteh</strong><br>
      ARB Registered Architect<br>
      Extension Survival Guide
    </p>

    <p style="margin-top: 20px; font-size: 12px; color: #94a3b8;">
      You're receiving this because you signed up for the Extension Planning Checklist.<br>
      <a href="${BASE_URL}" style="color: #94a3b8;">extensionsurvivalguide.co.uk</a>
    </p>
  </div>
</body>
</html>
      `,
    });

    // Send notification to yourself
    await resend.emails.send({
      from: `Extension Survival Guide <${BUSINESS_EMAIL}>`,
      to: BUSINESS_EMAIL,
      subject: `New Checklist Signup: ${email}`,
      html: `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
</head>
<body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #1e293b; max-width: 600px; margin: 0 auto; padding: 20px;">
  <h2 style="color: #16a34a;">New Checklist Signup</h2>

  <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
    <tr>
      <td style="padding: 8px 0; border-bottom: 1px solid #e2e8f0; font-weight: 600;">Email</td>
      <td style="padding: 8px 0; border-bottom: 1px solid #e2e8f0;">${email}</td>
    </tr>
    <tr>
      <td style="padding: 8px 0; border-bottom: 1px solid #e2e8f0; font-weight: 600;">Source</td>
      <td style="padding: 8px 0; border-bottom: 1px solid #e2e8f0;">${source || 'unknown'}</td>
    </tr>
    <tr>
      <td style="padding: 8px 0; font-weight: 600;">Date</td>
      <td style="padding: 8px 0;">${new Date().toLocaleString('en-GB', { timeZone: 'Europe/London' })}</td>
    </tr>
  </table>

  <p style="color: #64748b; font-size: 14px;">Checklist email has been sent to the subscriber automatically.</p>
</body>
</html>
      `,
    });

    return NextResponse.json({ success: true });
  } catch (error) {
    console.error('Subscribe error:', error);
    return NextResponse.json(
      { error: 'Something went wrong. Please try again.' },
      { status: 500 }
    );
  }
}
