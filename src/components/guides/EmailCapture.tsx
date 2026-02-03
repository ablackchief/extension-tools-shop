'use client';

import { useState } from 'react';
import { Send, Check, Loader2 } from 'lucide-react';
import { Button } from '@/components/ui/button';

export function EmailCapture() {
  const [email, setEmail] = useState('');
  const [status, setStatus] = useState<'idle' | 'loading' | 'success' | 'error'>('idle');
  const [errorMessage, setErrorMessage] = useState('');

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setStatus('loading');
    setErrorMessage('');

    try {
      const response = await fetch('/api/subscribe', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, source: 'guides' }),
      });

      if (response.ok) {
        setStatus('success');
        setEmail('');
      } else {
        const data = await response.json();
        setErrorMessage(data.error || 'Something went wrong. Please try again.');
        setStatus('error');
      }
    } catch {
      setErrorMessage('Something went wrong. Please try again.');
      setStatus('error');
    }
  };

  return (
    <section className="bg-gradient-to-br from-slate-900 via-blue-900 to-slate-900 py-16 md:py-20">
      <div className="max-w-3xl mx-auto px-4 text-center">
        <h2 className="text-2xl md:text-3xl font-bold text-white mb-4">
          Get the Extension Planning Checklist Free
        </h2>
        <p className="text-slate-300 mb-8 max-w-xl mx-auto">
          A one-page PDF covering the 15 steps most homeowners skip. Sent straight to your inbox.
        </p>

        {status === 'success' ? (
          <div className="bg-green-500/20 border border-green-500/30 rounded-xl p-6 max-w-md mx-auto">
            <div className="flex items-center justify-center gap-2 text-green-400 mb-2">
              <Check className="w-5 h-5" />
              <span className="font-semibold">Check your inbox</span>
            </div>
            <p className="text-slate-300 text-sm">
              We&apos;ve sent the checklist to your email address.
            </p>
          </div>
        ) : (
          <form onSubmit={handleSubmit} className="flex flex-col sm:flex-row gap-3 max-w-md mx-auto">
            <input
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              placeholder="Enter your email"
              required
              className="flex-1 px-4 py-3 rounded-lg bg-white/10 border border-white/20 text-white placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
            <Button
              type="submit"
              variant="orange"
              disabled={status === 'loading'}
              className="whitespace-nowrap"
            >
              {status === 'loading' ? (
                <Loader2 className="w-4 h-4 animate-spin" />
              ) : (
                <>
                  <Send className="w-4 h-4 mr-2" />
                  Send Me the Checklist
                </>
              )}
            </Button>
          </form>
        )}

        {status === 'error' && (
          <p className="text-red-400 text-sm mt-3">{errorMessage}</p>
        )}

        <p className="text-slate-500 text-xs mt-4">
          No spam. Unsubscribe anytime. Your email is used only for extension planning resources.
        </p>
      </div>
    </section>
  );
}
