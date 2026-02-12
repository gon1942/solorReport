import { NextRequest, NextResponse } from 'next/server';
import { access, mkdir } from 'node:fs/promises';
import path from 'node:path';
import { constants } from 'node:fs';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

export const runtime = 'nodejs';

const execFileAsync = promisify(execFile);

function toSafeFolderName(fileName: string): string {
  const base = path.parse(fileName).name;
  return base.replace(/[^a-zA-Z0-9._-]/g, '_') || 'report';
}

async function resolveInputFilePath(fileName: string): Promise<string | null> {
  const candidates = [path.join(process.cwd(), 'upload', fileName)];

  for (const candidate of candidates) {
    try {
      await access(candidate, constants.R_OK);
      return candidate;
    } catch {
      // no-op
    }
  }

  return null;
}

async function runPythonCommand(args: string[], env: NodeJS.ProcessEnv = process.env) {
  const pythonCandidates = [
    process.env.PYTHON_BIN,
    path.join(process.cwd(), 'scripts', '.venv', 'bin', 'python'),
    'python3',
    'python'
  ].filter((v): v is string => Boolean(v));

  let lastError: unknown = null;

  for (const python of pythonCandidates) {
    try {
        return await execFileAsync(python, args, {
          cwd: process.cwd(),
          maxBuffer: 10 * 1024 * 1024,
          env
        });
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError;
}

export async function POST(req: NextRequest) {
  try {
    const body = await req.json().catch(() => ({}));
    const fileName = String(body?.fileName || '').trim();

    if (!fileName) {
      return NextResponse.json(
        { success: false, error: { message: 'fileName이 필요합니다.' } },
        { status: 400 }
      );
    }

    const sourceFilePath = await resolveInputFilePath(fileName);
    if (!sourceFilePath) {
      return NextResponse.json(
        { success: false, error: { message: `입력 파일을 찾지 못했습니다: ${fileName}` } },
        { status: 404 }
      );
    }

    const folderName = toSafeFolderName(fileName);
    const targetDir = path.join(process.cwd(), 'data', folderName);
    await mkdir(targetDir, { recursive: true });

    const commands = [
      {
        name: 'capex',
        args: [
          './pv_solar/convert_to_json/capex.py',
          sourceFilePath,
          '-o',
          `data/${folderName}/capex.json`
        ]
      },
      {
        name: 'opex_year1',
        args: [
          './pv_solar/convert_to_json/opex.py',
          sourceFilePath,
          '--mode',
          'opex_year1',
          '-o',
          `data/${folderName}/opex_year1.json`
        ]
      },
      {
        name: 'general',
        args: [
          './pv_solar/convert_to_json/extract_general_inputs.py',
          sourceFilePath,
          '-o',
          `data/${folderName}/general.json`
        ]
      }
    ];

    const outputDir = path.join(targetDir, 'output');
    await mkdir(outputDir, { recursive: true });
    const reportFileName = `${folderName}_capex.pptx`;
    const outputPptPath = path.join(outputDir, reportFileName);

    const logs: Array<{ step: string; stdout: string; stderr: string }> = [];

    for (const cmd of commands) {
      const result = await runPythonCommand(cmd.args);
      logs.push({
        step: cmd.name,
        stdout: result.stdout?.toString() || '',
        stderr: result.stderr?.toString() || ''
      });
    }

    const documentResult = await runPythonCommand(
      ['./pv_solar/capex_pptx.py'],
      {
        ...process.env,
        CAPEX_DATA_DIR: targetDir,
        CAPEX_OUTPUT_PATH: outputPptPath
      }
    );
    logs.push({
      step: 'capex_pptx',
      stdout: documentResult.stdout?.toString() || '',
      stderr: documentResult.stderr?.toString() || ''
    });

    return NextResponse.json({
      success: true,
      data: {
        fileName,
        sourceFilePath,
        outputDir: `data/${folderName}`,
        documentOutputDir: `data/${folderName}/output`,
        documentFile: `data/${folderName}/output/${reportFileName}`,
        files: {
          capex: `data/${folderName}/capex.json`,
          opex_year1: `data/${folderName}/opex_year1.json`,
          general: `data/${folderName}/general.json`,
          report: `data/${folderName}/output/${reportFileName}`
        },
        logs
      }
    });
  } catch (error) {
    return NextResponse.json(
      {
        success: false,
        error: {
          message: error instanceof Error ? error.message : '리포트 데이터 생성 중 오류가 발생했습니다.'
        }
      },
      { status: 500 }
    );
  }
}
