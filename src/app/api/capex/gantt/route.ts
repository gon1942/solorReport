import { promises as fs } from 'fs';
import path from 'path';
import { NextResponse } from 'next/server';

export const runtime = 'nodejs';

type GanttTask = {
  id: string;
  name: string;
  type: string;
  parentId?: string;
  startWeekIndex: number;
  endWeekIndex: number;
};

type GanttPayload = {
  phase1: GanttTask[];
  phase2: GanttTask[];
};

const dataDir = path.join(process.cwd(), 'data_gantt');
const phase1Path = path.join(dataDir, 'phase1.json');
const phase2Path = path.join(dataDir, 'phase2.json');

const normalizeTasks = (value: unknown): GanttTask[] => {
  if (!Array.isArray(value)) return [];
  return value
    .filter((item): item is GanttTask => {
      if (!item || typeof item !== 'object') return false;
      const candidate = item as Record<string, unknown>;
      return (
        typeof candidate.id === 'string' &&
        typeof candidate.name === 'string' &&
        typeof candidate.type === 'string' &&
        typeof candidate.startWeekIndex === 'number' &&
        typeof candidate.endWeekIndex === 'number'
      );
    })
    .map((task) => ({
      id: task.id,
      name: task.name,
      type: task.type,
      parentId: typeof task.parentId === 'string' ? task.parentId : undefined,
      startWeekIndex: task.startWeekIndex,
      endWeekIndex: task.endWeekIndex
    }));
};

const readTasks = async (filePath: string): Promise<GanttTask[]> => {
  try {
    const raw = await fs.readFile(filePath, 'utf-8');
    return normalizeTasks(JSON.parse(raw));
  } catch {
    return [];
  }
};

export async function GET() {
  const phase1 = await readTasks(phase1Path);
  const phase2 = await readTasks(phase2Path);
  return NextResponse.json({ phase1, phase2 });
}

export async function POST(request: Request) {
  try {
    const body = (await request.json()) as Partial<GanttPayload>;
    const phase1 = normalizeTasks(body?.phase1);
    const phase2 = normalizeTasks(body?.phase2);

    await fs.mkdir(dataDir, { recursive: true });
    await fs.writeFile(phase1Path, JSON.stringify(phase1, null, 2), 'utf-8');
    await fs.writeFile(phase2Path, JSON.stringify(phase2, null, 2), 'utf-8');

    return NextResponse.json({ success: true, saved: { phase1: phase1.length, phase2: phase2.length } });
  } catch (error: any) {
    return NextResponse.json({ success: false, error: error?.message || 'Failed to save gantt data' }, { status: 500 });
  }
}
