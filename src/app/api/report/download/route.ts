import { NextRequest, NextResponse } from 'next/server';
import { access, readFile } from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';

export const runtime = 'nodejs';

function detectContentType(filePath: string): string {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.pptx') {
    return 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  }
  if (ext === '.pdf') {
    return 'application/pdf';
  }
  if (ext === '.json') {
    return 'application/json; charset=utf-8';
  }
  return 'application/octet-stream';
}

export async function GET(req: NextRequest) {
  try {
    const rawFile = req.nextUrl.searchParams.get('file') || '';
    if (!rawFile.trim()) {
      return NextResponse.json(
        { success: false, error: { message: 'file 파라미터가 필요합니다.' } },
        { status: 400 }
      );
    }

    const normalizedRelativePath = rawFile.replace(/\\/g, '/').replace(/^\/+/, '');
    const dataRoot = path.resolve(process.cwd(), 'data');
    const targetPath = path.resolve(process.cwd(), normalizedRelativePath);

    if (!(targetPath === dataRoot || targetPath.startsWith(`${dataRoot}${path.sep}`))) {
      return NextResponse.json(
        { success: false, error: { message: '허용되지 않은 경로입니다.' } },
        { status: 403 }
      );
    }

    await access(targetPath, constants.R_OK);
    const payload = await readFile(targetPath);
    const fileName = path.basename(targetPath);
    const encodedName = encodeURIComponent(fileName);

    return new NextResponse(payload, {
      headers: {
        'Content-Type': detectContentType(targetPath),
        'Content-Disposition': `attachment; filename*=UTF-8''${encodedName}`,
        'Cache-Control': 'no-store'
      }
    });
  } catch (error) {
    return NextResponse.json(
      {
        success: false,
        error: {
          message: error instanceof Error ? error.message : '파일 다운로드 중 오류가 발생했습니다.'
        }
      },
      { status: 500 }
    );
  }
}
