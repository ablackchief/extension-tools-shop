import { Resend } from 'resend';
import { BundleType } from '@/types';
import { BUNDLES } from './constants';

// Lazy initialization to avoid build-time errors
let resend: Resend | null = null;
function getResend(): Resend {
  if (!resend) {
    resend = new Resend((process.env.RESEND_API_KEY || '').trim());
  }
  return resend;
}

const BUSINESS_EMAIL = (process.env.BUSINESS_EMAIL || 'support@extensionsurvivalguide.co.uk').trim();
const BASE_URL = 'https://extensionsurvivalguide.co.uk';

export async function sendDownloadEmail(
  email: string,
  bundle: BundleType,
  token: string
): Promise<boolean> {
  const bundleData = BUNDLES.find(b => b.id === bundle);
  const downloadUrl = `${BASE_URL}/download/${token}`;

  try {
    await getResend().emails.send({
      from: `Extension Survival Guide <${BUSINESS_EMAIL}>`,
      to: email,
      subject: `Your ${bundleData?.name || 'Extension Tools'} Download Link`,
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
    <h2 style="margin: 0 0 15px; color: #1e293b;">Thank you for your purchase!</h2>
    <p style="margin: 0; color: #475569;">Your <strong>${bundleData?.name}</strong> is ready to download.</p>
  </div>

  <div style="text-align: center; margin: 30px 0;">
    <a href="${downloadUrl}" style="display: inline-block; background: linear-gradient(135deg, #f97316 0%, #ea580c 100%); color: white; text-decoration: none; padding: 16px 32px; border-radius: 8px; font-weight: 600; font-size: 16px;">
      Download Your Tools
    </a>
  </div>

  <div style="background: #f8fafc; border-radius: 8px; padding: 20px; margin-bottom: 30px;">
    <h3 style="margin: 0 0 15px; color: #1e293b;">What's included:</h3>
    <ul style="margin: 0; padding-left: 20px; color: #475569;">
      ${bundleData?.features.map(f => `<li>${f.name}${f.format ? ` (${f.format})` : ''}</li>`).join('')}
    </ul>
  </div>

  <div style="background: #fef3c7; border-left: 4px solid #f59e0b; padding: 15px; margin-bottom: 30px;">
    <p style="margin: 0; color: #92400e;">
      <strong>Important:</strong> This download link expires in 30 days. Save your files to your computer after downloading.
    </p>
  </div>

  <div style="margin-bottom: 30px;">
    <h3 style="color: #1e293b;">Getting Started:</h3>
    <ol style="color: #475569; padding-left: 20px;">
      <li>Download and unzip your files</li>
      <li>Start with the Extension Decision Calculator</li>
      <li>Move to the Budget Planner once you've decided</li>
      <li>Use the Project Brief before talking to architects</li>
    </ol>
  </div>

  <div style="border-top: 1px solid #e2e8f0; padding-top: 20px; text-align: center; color: #64748b; font-size: 14px;">
    <p>Need help? Reply to this email or contact us at:<br>
    <a href="mailto:${BUSINESS_EMAIL}" style="color: #1a56f5;">${BUSINESS_EMAIL}</a></p>

    <p style="margin-top: 20px;">
      <strong>Abre Etteh</strong><br>
      ARB Registered Architect<br>
      Extension Survival Guide
    </p>
  </div>
</body>
</html>
      `,
    });

    return true;
  } catch (error) {
    console.error('Failed to send download email:', error);
    return false;
  }
}

export async function sendOrderNotification(
  customerEmail: string,
  bundle: BundleType,
  amountPaid: number
): Promise<boolean> {
  const bundleData = BUNDLES.find(b => b.id === bundle);

  try {
    await getResend().emails.send({
      from: `Extension Survival Guide <${BUSINESS_EMAIL}>`,
      to: BUSINESS_EMAIL,
      subject: `New Order: ${bundleData?.name || bundle} - £${(amountPaid / 100).toFixed(2)}`,
      html: `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
</head>
<body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #1e293b; max-width: 600px; margin: 0 auto; padding: 20px;">
  <h2 style="color: #16a34a;">New Order Received</h2>

  <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
    <tr>
      <td style="padding: 8px 0; border-bottom: 1px solid #e2e8f0; font-weight: 600;">Bundle</td>
      <td style="padding: 8px 0; border-bottom: 1px solid #e2e8f0;">${bundleData?.name || bundle}</td>
    </tr>
    <tr>
      <td style="padding: 8px 0; border-bottom: 1px solid #e2e8f0; font-weight: 600;">Amount</td>
      <td style="padding: 8px 0; border-bottom: 1px solid #e2e8f0;">£${(amountPaid / 100).toFixed(2)}</td>
    </tr>
    <tr>
      <td style="padding: 8px 0; border-bottom: 1px solid #e2e8f0; font-weight: 600;">Customer</td>
      <td style="padding: 8px 0; border-bottom: 1px solid #e2e8f0;">${customerEmail}</td>
    </tr>
    <tr>
      <td style="padding: 8px 0; font-weight: 600;">Date</td>
      <td style="padding: 8px 0;">${new Date().toLocaleString('en-GB', { timeZone: 'Europe/London' })}</td>
    </tr>
  </table>

  <p style="color: #64748b; font-size: 14px;">Download email has been sent to the customer automatically.</p>
</body>
</html>
      `,
    });

    return true;
  } catch (error) {
    console.error('Failed to send order notification:', error);
    return false;
  }
}
