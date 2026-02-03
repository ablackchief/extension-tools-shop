"use client";

import { useState, use } from "react";
import { Download, Loader2, AlertCircle, CheckCircle, FileArchive } from "lucide-react";
import { Button } from "@/components/ui/button";
import Link from "next/link";

interface DownloadPageProps {
  params: Promise<{ token: string }>;
}

export default function DownloadPage({ params }: DownloadPageProps) {
  const { token } = use(params);
  const [downloading, setDownloading] = useState(false);
  const [downloaded, setDownloaded] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleDownload = async () => {
    setDownloading(true);
    setError(null);

    try {
      const response = await fetch(`/api/download/${token}`);

      if (!response.ok) {
        const data = await response.json();
        throw new Error(data.error || "Download failed");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "Extension-Survival-Guide-Tools.zip";
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      setDownloaded(true);
    } catch (err: any) {
      setError(err.message || "Download failed. Please try again.");
    } finally {
      setDownloading(false);
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-slate-50 flex items-center justify-center p-4">
      <div className="max-w-lg w-full bg-white rounded-2xl shadow-xl p-8 text-center">
        <div className="w-20 h-20 bg-blue-100 rounded-full flex items-center justify-center mx-auto mb-6">
          {downloaded ? (
            <CheckCircle className="w-10 h-10 text-green-600" />
          ) : error ? (
            <AlertCircle className="w-10 h-10 text-red-600" />
          ) : (
            <FileArchive className="w-10 h-10 text-blue-600" />
          )}
        </div>

        <h1 className="text-2xl font-bold text-gray-900 mb-4">
          {downloaded
            ? "Download Complete!"
            : error
            ? "Download Error"
            : "Your Tools Are Ready"}
        </h1>

        {error ? (
          <div className="bg-red-50 text-red-700 p-4 rounded-lg mb-6">
            <p>{error}</p>
            <p className="mt-2 text-sm">
              If this problem persists, please contact
              support@extensionsurvivalguide.co.uk
            </p>
          </div>
        ) : downloaded ? (
          <div className="bg-green-50 text-green-700 p-4 rounded-lg mb-6">
            <p>
              Your files have been downloaded. Check your Downloads folder for the
              ZIP file.
            </p>
          </div>
        ) : (
          <p className="text-gray-600 mb-8">
            Click the button below to download your Extension Survival Guide tools
            as a ZIP file.
          </p>
        )}

        <Button
          size="lg"
          className="w-full bg-gradient-to-r from-blue-600 to-blue-700 mb-6"
          onClick={handleDownload}
          disabled={downloading}
        >
          {downloading ? (
            <>
              <Loader2 className="w-5 h-5 mr-2 animate-spin" />
              Preparing Download...
            </>
          ) : downloaded ? (
            <>
              <Download className="w-5 h-5 mr-2" />
              Download Again
            </>
          ) : (
            <>
              <Download className="w-5 h-5 mr-2" />
              Download Tools (ZIP)
            </>
          )}
        </Button>

        {downloaded && (
          <div className="bg-slate-50 rounded-lg p-6 text-left mb-6">
            <h3 className="font-semibold text-gray-900 mb-3">Getting Started:</h3>
            <ol className="space-y-2 text-sm text-gray-600">
              <li>1. Unzip the downloaded file</li>
              <li>2. Open the Extension Decision Calculator first</li>
              <li>3. Work through the Budget Planner next</li>
              <li>4. Complete the Project Brief before hiring</li>
            </ol>
          </div>
        )}

        <Link href="/" className="text-blue-600 hover:underline text-sm">
          Return to Homepage
        </Link>

        <p className="mt-6 text-xs text-gray-400">
          This download link expires in 30 days. Need help?
          <br />
          Contact support@extensionsurvivalguide.co.uk
        </p>
      </div>
    </div>
  );
}
