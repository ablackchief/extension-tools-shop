"use client";

import { useEffect, useState, Suspense } from "react";
import { useSearchParams } from "next/navigation";
import { CheckCircle, Download, Mail, Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import Link from "next/link";

function SuccessContent() {
  const searchParams = useSearchParams();
  const sessionId = searchParams.get("session_id");
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [email, setEmail] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (sessionId) {
      fetch(`/api/download-url?session_id=${sessionId}`)
        .then((res) => res.json())
        .then((data) => {
          if (data.downloadUrl) {
            setDownloadUrl(data.downloadUrl);
            setEmail(data.email);
          } else {
            setError(data.error || "Failed to get download link");
          }
        })
        .catch(() => {
          setError("Failed to connect to server");
        })
        .finally(() => {
          setLoading(false);
        });
    } else {
      setLoading(false);
      setError("No session ID provided");
    }
  }, [sessionId]);

  if (loading) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-green-50 to-blue-50 flex items-center justify-center p-4">
        <div className="text-center">
          <Loader2 className="w-12 h-12 text-blue-600 animate-spin mx-auto mb-4" />
          <p className="text-slate-600">Preparing your download...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-green-50 to-blue-50 flex items-center justify-center p-4">
      <div className="max-w-lg w-full bg-white rounded-2xl shadow-xl p-8 text-center">
        <div className="w-20 h-20 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-6">
          <CheckCircle className="w-10 h-10 text-green-600" />
        </div>

        <h1 className="text-3xl font-bold text-gray-900 mb-4">
          Thank You for Your Purchase!
        </h1>

        <p className="text-gray-600 mb-8">
          Your Extension Survival Guide tools are ready to download.
        </p>

        {error ? (
          <div className="bg-red-50 text-red-700 p-4 rounded-lg mb-6">
            <p>{error}</p>
            <p className="mt-2 text-sm">
              Please contact support@extensionsurvivalguide.co.uk for assistance.
            </p>
          </div>
        ) : (
          downloadUrl && (
            <Button
              size="lg"
              className="w-full bg-gradient-to-r from-blue-600 to-blue-700 mb-4"
              onClick={() => (window.location.href = downloadUrl)}
            >
              <Download className="w-5 h-5 mr-2" />
              Download Your Tools (ZIP)
            </Button>
          )
        )}

        {email && (
          <div className="flex items-center justify-center gap-2 text-sm text-gray-500 mb-8">
            <Mail className="w-4 h-4" />
            <span>Download link also sent to {email}</span>
          </div>
        )}

        <div className="bg-gray-50 rounded-lg p-6 text-left">
          <h3 className="font-semibold text-gray-900 mb-3">What&apos;s Next?</h3>
          <ol className="space-y-2 text-sm text-gray-600">
            <li>1. Download and unzip your files</li>
            <li>2. Start with the Extension Decision Calculator</li>
            <li>3. Move to the Budget Planner once you&apos;ve decided</li>
            <li>4. Use the Project Brief before talking to architects</li>
          </ol>
        </div>

        <div className="mt-8 flex flex-col gap-4">
          <Link href="/" className="text-blue-600 hover:underline text-sm">
            Return to Homepage
          </Link>
          <p className="text-xs text-gray-400">
            Download available for 30 days. Need help? Contact
            support@extensionsurvivalguide.co.uk
          </p>
        </div>
      </div>
    </div>
  );
}

export default function SuccessPage() {
  return (
    <Suspense
      fallback={
        <div className="min-h-screen bg-gradient-to-br from-green-50 to-blue-50 flex items-center justify-center p-4">
          <Loader2 className="w-12 h-12 text-blue-600 animate-spin" />
        </div>
      }
    >
      <SuccessContent />
    </Suspense>
  );
}
