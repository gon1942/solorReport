import { NextRequest, NextResponse } from 'next/server';
import ExcelJS from 'exceljs';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;

    if (!file) {
      return NextResponse.json(
        { success: false, error: 'No file provided' },
        { status: 400 }
      );
    }

    // Excel 파일 분석
    const buffer = Buffer.from(await file.arrayBuffer()) as any;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    const sheets = workbook.worksheets.map(worksheet => {
      const jsonData: any[][] = [];
      worksheet.eachRow((row, rowNumber) => {
        const rowData: any[] = [];
        row.eachCell({ includeEmpty: true }, (cell) => {
          let value = cell.value;

          // richText 객체를 텍스트로 변환
          if (typeof value === 'object' && value !== null && 'richText' in value) {
            value = value.richText.map((t: any) => t.text).join('');
          }

          rowData.push(value ?? '');
        });
        jsonData.push(rowData);
      });

      const headers = jsonData[0] as string[] || [];
      const preview = jsonData.slice(0, 11); // 첫 10행 미리보기

      return {
        name: worksheet.name,
        headers,
        preview,
        rowCount: jsonData.length,
        columnCount: headers.length,
        isPVSolarData: true
      };
    });

    return NextResponse.json({
      success: true,
      data: {
        sheets,
        totalSheets: sheets.length
      }
    });

  } catch (error) {
    console.error('PV Solar Excel analysis error:', error);
    return NextResponse.json(
      { 
        success: false, 
        error: error instanceof Error ? error.message : 'Analysis failed' 
      },
      { status: 500 }
    );
  }
}