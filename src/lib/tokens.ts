import crypto from 'crypto';
import { BundleType, DownloadToken, Tool } from '@/types';
import { TOOLS } from './constants';

const TOKEN_SECRET = process.env.DOWNLOAD_TOKEN_SECRET || 'default-secret';
const TOKEN_EXPIRY_DAYS = 30;

// In-memory token store (in production, use a database)
const tokenStore = new Map<string, DownloadToken>();

export function generateDownloadToken(bundle: BundleType, email: string): string {
  const token = crypto.randomBytes(32).toString('hex');
  const now = Date.now();

  const tokenData: DownloadToken = {
    bundle,
    email,
    createdAt: now,
    expiresAt: now + (TOKEN_EXPIRY_DAYS * 24 * 60 * 60 * 1000),
  };

  tokenStore.set(token, tokenData);

  return token;
}

export function verifyDownloadToken(token: string): DownloadToken | null {
  const tokenData = tokenStore.get(token);

  if (!tokenData) {
    return null;
  }

  if (Date.now() > tokenData.expiresAt) {
    tokenStore.delete(token);
    return null;
  }

  return tokenData;
}

export function getFilesForBundle(bundle: BundleType): Tool[] {
  return TOOLS.filter(tool => tool.bundle.includes(bundle));
}

export function getReadmeContent(bundle: BundleType): string {
  const bundleName = bundle === 'essential' ? 'Essential Tools' :
                     bundle === 'complete' ? 'Complete Toolkit' :
                     'Professional Plus';

  return `
THE EXTENSION SURVIVAL GUIDE
${bundleName} Bundle
====================================

Thank you for your purchase!

GETTING STARTED
---------------
1. Start with the Extension Decision Calculator to confirm extending is right for you
2. Use the Budget Planner to set realistic expectations
3. Complete the Project Brief before speaking to architects or builders

SUPPORT
-------
If you need help with any of these tools:
- Email: support@extensionsurvivalguide.co.uk
- Website: https://extensionsurvivalguide.co.uk

LICENCE
-------
These tools are licensed for personal use on one project.
You may share with household members working on the same project.
Please do not redistribute or resell.

Created by Abre Etteh
ARB Registered Architect

(c) ${new Date().getFullYear()} Extension Survival Guide
`.trim();
}
