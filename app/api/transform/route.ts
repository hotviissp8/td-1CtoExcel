import { NextRequest, NextResponse } from 'next/server';
import { transformExcel } from '@/lib/transform';

export const runtime = 'nodejs';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;
    if (!file) return NextResponse.json({ error: 'No file' }, { status: 400 });

    const { buffer: out, warnings, rowCount } = await transformExcel(await file.arrayBuffer());

    const baseName = file.name.replace(/\.xlsx?$/i, '');
    return new NextResponse(new Uint8Array(out), {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(baseName + '_result.xlsx')}`,
        'X-Warnings': encodeURIComponent(JSON.stringify(warnings)),
        'X-Row-Count': String(rowCount),
        'X-Output-Name': encodeURIComponent(baseName + '_result.xlsx'),
      },
    });
  } catch (e) {
    return NextResponse.json(
      { error: e instanceof Error ? e.message : 'Unknown error' },
      { status: 500 }
    );
  }
}
