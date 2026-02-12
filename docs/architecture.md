# Architecture

## Overview

이 문서는 `PV Solar Workspace`의 현재 코드 기준 아키텍처를 간략히 설명합니다.

## Frontend

- 프레임워크: Next.js 14 + React 18 + TypeScript
- 메인 페이지: `src/app/page.tsx`
- Planner 상태 페이지: `src/app/planner/page.tsx`
- 모달 컴포넌트: `src/components/PVSolarExcelPreprocessModal.tsx`

## API Routes

- `POST /api/rag/upload`
- `POST /api/preprocess/pv-solar/excel/analyze`
- `POST /api/preprocess/pv-solar/excel/recommend`
- `GET /api/planner/status`
- `POST /api/planner/pptx-template`

## Data and State

- CAPEX 간트 기본 데이터: `src/data/capex-gantt.json`
- CAPEX 간트 편집 상태: 페이지 state로 관리
- CAPEX 간트 저장: 브라우저 localStorage (`pvsolar.capex.gantt`)
- Planner 상태 캐시: Redis 사용 가능 (`src/lib/redisClient.ts`)

## PV Solar 디렉터리 메모

- `pv_solar/test3` 디렉터리는 `pv_solar/convert_to_json`으로 이름이 변경되었습니다.
- `pv_solar/test_x`, `pv_solar/test2_x`는 임시/미사용 폴더입니다.

## External Integration

- API 서버 URL: `NEXT_PUBLIC_API_SERVER_URL` 또는 서버 전용 `API_SERVER_URL`
- RAG 서버 URL: `NEXT_PUBLIC_RAG_API_SERVER` 또는 서버 전용 `RAG_API_SERVER`

## Notes

- 상세 설계(도메인 모델/시퀀스 다이어그램/배포 구조)는 추후 보강 예정입니다.
