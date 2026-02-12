import { NextResponse } from 'next/server';
import { getApiServerUrl } from '@/config/serverConfig';

type TemplateRequest = {
  projectName?: string;
  audience?: string;
  tone?: string;
  language?: string;
  sections?: string[];
  constraints?: string[];
  dataPoints?: Array<{ label: string; value: string }>;
};

type TemplateResponse = {
  title: string;
  theme: {
    mood: string;
    palette: string[];
    typography: string[];
  };
  slides: Array<{
    id: string;
    layout: string;
    title: string;
    bullets?: string[];
    notes?: string;
    visuals?: string[];
  }>;
};

const buildPrompt = (payload: TemplateRequest) => {
  const sections = payload.sections?.length ? payload.sections.join(', ') : '표지, 개요, 핵심 지표, 일정, 리스크, 결론';
  const constraints = payload.constraints?.length ? payload.constraints.join(' | ') : '실행 가능한 PPTX 템플릿 구조, 8-12 슬라이드, 시각 요소 지시 포함';
  const dataPoints = payload.dataPoints?.length
    ? payload.dataPoints.map((d) => `${d.label}: ${d.value}`).join('\n')
    : 'N/A';

  return `너는 에너지 프로젝트용 PPTX 템플릿 설계자다. 아래 입력을 반영해 PPTX 템플릿 구조를 JSON으로만 출력한다.

[요청 정보]
- 프로젝트: ${payload.projectName || 'PV Solar Planner'}
- 대상: ${payload.audience || '운영/기획팀'}
- 톤: ${payload.tone || '정확하고 간결'}
- 언어: ${payload.language || 'ko'}
- 섹션: ${sections}
- 제약: ${constraints}
- 데이터 포인트:
${dataPoints}

[출력 형식]
다음 스키마에 맞는 JSON만 출력:
{
  "title": string,
  "theme": { "mood": string, "palette": string[], "typography": string[] },
  "slides": [
    { "id": "slide-1", "layout": string, "title": string, "bullets": string[], "notes": string, "visuals": string[] }
  ]
}

주의: 설명 문장, 코드 블록, 주석 없이 JSON만 출력.`;
};

const parseJson = (value: string) => {
  try {
    return JSON.parse(value);
  } catch (error) {
    const jsonMatch = value.match(/```json\s*([\s\S]*?)\s*```/) || value.match(/\{[\s\S]*\}/);
    if (!jsonMatch) return null;
    try {
      return JSON.parse(jsonMatch[1] || jsonMatch[0]);
    } catch (err) {
      return null;
    }
  }
};

export async function POST(req: Request) {
  const payload = (await req.json().catch(() => ({}))) as TemplateRequest;
  const prompt = buildPrompt(payload);

  try {
    const apiServerUrl = getApiServerUrl();
    const response = await fetch(`${apiServerUrl}/callai`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        messages: [{ role: 'user', content: prompt }],
        options: {
          provider: 'ollama',
          model: 'airun-chat:latest',
          temperature: 0.2,
          max_tokens: 3500
        }
      })
    });

    if (!response.ok) {
      return NextResponse.json({ success: false, error: `AI API 호출 실패: ${response.status}` }, { status: 500 });
    }

    const result = await response.json();
    const aiResponse = result?.data?.response?.content || result?.data?.response || '';
    const parsed = parseJson(String(aiResponse));

    if (!parsed) {
      return NextResponse.json({ success: false, error: 'JSON 파싱 실패', raw: aiResponse }, { status: 200 });
    }

    return NextResponse.json({ success: true, data: parsed as TemplateResponse });
  } catch (error: any) {
    return NextResponse.json({ success: false, error: error?.message || 'Unknown error' }, { status: 500 });
  }
}
