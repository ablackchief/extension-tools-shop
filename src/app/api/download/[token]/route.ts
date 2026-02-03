import { NextResponse } from 'next/server';
import { verifyDownloadToken, getFilesForBundle, getReadmeContent } from '@/lib/tokens';
import JSZip from 'jszip';

// Dynamic imports for generators to avoid loading them all at once
async function getGenerator(generatorName: string) {
  try {
    const generators = await import('@/lib/generators');
    return (generators as any)[generatorName];
  } catch (error) {
    console.error(`Generator ${generatorName} not found:`, error);
    return null;
  }
}

export async function GET(
  req: Request,
  { params }: { params: Promise<{ token: string }> }
) {
  const { token } = await params;

  // Verify token
  const tokenData = verifyDownloadToken(token);

  if (!tokenData) {
    return NextResponse.json(
      { error: 'Invalid or expired download link' },
      { status: 403 }
    );
  }

  const { bundle } = tokenData;

  try {
    // Generate all files for the bundle
    const zip = new JSZip();
    const folder = zip.folder('Extension-Survival-Guide-Tools');

    if (!folder) {
      throw new Error('Failed to create zip folder');
    }

    const files = getFilesForBundle(bundle);

    for (const file of files) {
      try {
        const generator = await getGenerator(`generate${file.generator.charAt(0).toUpperCase() + file.generator.slice(1).replace(/-([a-z])/g, (_, c) => c.toUpperCase())}`);

        if (generator) {
          const content = await generator();
          folder.file(file.filename, content);
        } else {
          // Create placeholder file if generator not found
          folder.file(
            file.filename,
            `Placeholder for ${file.name}. Generator not yet implemented.`
          );
        }
      } catch (error) {
        console.error(`Error generating ${file.filename}:`, error);
        folder.file(
          file.filename.replace(/\.(xlsx|docx)$/, '.txt'),
          `Error generating ${file.name}. Please contact support.`
        );
      }
    }

    // Add README
    folder.file('README.txt', getReadmeContent(bundle));

    const zipContent = await zip.generateAsync({ type: 'blob' });

    const bundleName = bundle.charAt(0).toUpperCase() + bundle.slice(1);

    return new NextResponse(zipContent, {
      headers: {
        'Content-Type': 'application/zip',
        'Content-Disposition': `attachment; filename="Extension-Tools-${bundleName}.zip"`,
      },
    });
  } catch (error) {
    console.error('Download error:', error);
    return NextResponse.json(
      { error: 'Failed to generate download' },
      { status: 500 }
    );
  }
}
