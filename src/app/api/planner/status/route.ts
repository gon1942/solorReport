import { NextResponse } from 'next/server';
import { getRedisClient } from '@/lib/redisClient';

type PlannerStatus = {
  id: string;
  name: string;
  status: 'queued' | 'running' | 'completed' | 'failed';
  progress: number;
  owner: string;
  priority: 'low' | 'normal' | 'high';
  updatedAt: string;
  message?: string;
};

const mockStatuses: PlannerStatus[] = [
  {
    id: 'plan-001',
    name: '풍력·태양광 믹스 계획',
    status: 'running',
    progress: 42,
    owner: 'ops-bot',
    priority: 'high',
    updatedAt: new Date().toISOString(),
    message: '시나리오 8/20 분석 중'
  },
  {
    id: 'plan-002',
    name: 'PV 유지보수 일정 재계산',
    status: 'queued',
    progress: 0,
    owner: 'planner-ui',
    priority: 'normal',
    updatedAt: new Date().toISOString()
  },
  {
    id: 'plan-003',
    name: 'ESS 배치 최적화',
    status: 'completed',
    progress: 100,
    owner: 'optimizer',
    priority: 'high',
    updatedAt: new Date(Date.now() - 1000 * 60 * 5).toISOString(),
    message: '성공적으로 완료'
  }
];

const parseRedisPayload = (value: string | null): PlannerStatus[] => {
  if (!value) return [];
  try {
    const parsed = JSON.parse(value);
    if (Array.isArray(parsed)) return parsed as PlannerStatus[];
  } catch (error) {
    return [];
  }
  return [];
};

export async function GET() {
  const redis = getRedisClient();

  if (!redis) {
    return NextResponse.json({ success: true, source: 'mock', data: mockStatuses });
  }

  try {
    if (redis.status === 'wait' || redis.status === 'end') {
      await redis.connect();
    }

    const cached = await redis.get('planner:status');
    const parsed = parseRedisPayload(cached);

    if (parsed.length > 0) {
      return NextResponse.json({ success: true, source: 'redis', data: parsed });
    }

    return NextResponse.json({ success: true, source: 'redis-miss', data: mockStatuses });
  } catch (error: any) {
    return NextResponse.json({ success: false, source: 'redis-error', error: error?.message, data: mockStatuses }, { status: 200 });
  }
}
