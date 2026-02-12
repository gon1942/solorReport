# PV Solar Workspace

PV Solar 데이터 업로드, 전처리, RAG 연계, 리포트 생성을 위한 Next.js 기반 워크스페이스입니다.

이 프로젝트는 엑셀/CSV 기반의 태양광 데이터를 웹에서 처리하고,
분석 결과를 바탕으로 리포트 생성 워크플로우를 구성하는 데 목적이 있습니다.

## 무엇을 하는 프로젝트인가요?

- 파일 업로드부터 상태 확인, 데이터 미리보기, 리포트 생성 준비까지 한 화면에서 진행합니다.
- CAPEX 카테고리에서는 간트차트 편집(Phase 1 / Phase 2 분리)을 지원합니다.
- 서버 API를 통해 PV Solar 특화 분석 및 추천 로직과 연동됩니다.

## 주요 기능

- PV Solar 파일 업로드 및 검증
  - 지원 확장자: `.xlsx`, `.xls`, `.csv`, `.txt`
  - 업로드 크기 제한: 50MB
- 시트 분석 API
  - 시트/헤더/미리보기/행·열 개수 반환
- PV Solar 추천 API
  - 시트 선택, 핵심 필드 추천, 인사이트 생성
- CAPEX 리포트 준비 UI
  - 카테고리 선택, 마일스톤 입력, 간트차트 표시/편집
  - Phase 1 / Phase 2 간트 데이터를 분리 관리
  - 간트 편집값 로컬 저장(localStorage)
- Planner 상태 모니터링 화면
  - `/planner` 페이지에서 상태/우선순위/진행률 조회

## 화면/흐름 요약

1. `pnpm dev`로 개발 서버 실행
2. 메인 화면(`/`) 접속
3. 파일 업로드 후 처리 상태 및 청크 미리보기 확인
4. 리포트 카테고리에서 `CAPEX` 선택
5. Phase 1 / Phase 2 간트차트 편집 및 저장
6. 리포트 생성 payload 확인

## Quick Start

아래는 현재 저장소에서 확인 가능한 실행 명령입니다.

```bash
# 의존성 설치
pnpm install

# 개발 서버
pnpm dev

# 프로덕션 빌드
pnpm build

# 프로덕션 실행
pnpm start

# 타입 체크
pnpm type-check
```

> 참고: `package.json` 스크립트는 `next dev`, `next build`, `next start`, `next lint`, `tsc --noEmit`를 사용합니다.

## 기술 스택

- Frontend: Next.js 14, React 18, TypeScript
- UI: CSS Modules, lucide-react
- Data/Excel: exceljs
- Infra 연동: ioredis (Planner 상태 캐시), 외부 API 서버 연동

## API 엔드포인트(저장소 기준)

- `POST /api/rag/upload`
- `POST /api/preprocess/pv-solar/excel/analyze`
- `POST /api/preprocess/pv-solar/excel/recommend`
- `GET /api/planner/status`
- `POST /api/planner/pptx-template`

## 환경 변수

프로젝트 코드에서 사용 중인 주요 환경 변수는 다음과 같습니다.

- `NEXT_PUBLIC_API_SERVER_URL`
- `NEXT_PUBLIC_RAG_API_SERVER`
- `NEXT_PUBLIC_RAG_STATUS_KEY`
- `NEXT_PUBLIC_RAG_STATUS_BASE`
- `NEXT_PUBLIC_RAG_FILES_API`
- `NEXT_PUBLIC_RAG_DELETE_API`
- `NEXT_PUBLIC_RAG_CHUNKS_API`
- `NEXT_PUBLIC_RAG_DEFAULT_USER_ID`
- `REDIS_URL` 또는 `NEXT_PUBLIC_REDIS_URL`

## 프로젝트 구조(핵심)

```text
src/
  app/
    page.tsx                        # 메인 업로드/리포트 화면
    planner/page.tsx                # Planner 상태 화면
    api/
      rag/upload/route.ts
      preprocess/pv-solar/excel/analyze/route.ts
      preprocess/pv-solar/excel/recommend/route.ts
      planner/status/route.ts
      planner/pptx-template/route.ts
  components/
    PVSolarExcelPreprocessModal.tsx
  data/
    capex-gantt.json
pv_solar/
  capex_pptx.py
  solar_pptx.py
  convert_to_json/
  test_x/                         # 임시/미사용
  test2_x/                        # 임시/미사용
```

## Documentation

상세 문서는 아래 링크로 분리 관리합니다.

- [Architecture](docs/architecture.md)
- [Runbook](docs/runbook.md)
- [Contributing](CONTRIBUTING.md)

## 폴더 네이밍 메모

- `pv_solar/test3`는 `pv_solar/convert_to_json`으로 변경되었습니다.
- `pv_solar/test_x`, `pv_solar/test2_x`는 임시/미사용 폴더입니다.

## License / Contact

- License: 저장소 루트 라이선스 파일은 현재 확인되지 않습니다.
- Contact: 프로젝트 담당 정보는 추후 문서화 예정입니다.



## -------------------- ##

# PV Solar Workspace 코드 분석

## 개요
- Next.js 14(App Router) + React 18 + TypeScript 기반 단일 페이지 업로드/리포트 생성 UI.
- 목적: PV Solar 관련 Excel/CSV 업로드 → RAG 백엔드 연계 전처리, 청크 미리보기, CAPEX 간트/마일스톤 입력, 리포트 생성 흐름 제공.
- 주요 의존성: `exceljs`(서버 라우트에서 Excel 파싱), `lucide-react`(아이콘), Next 빌트인 API 라우트.

## 프런트엔드 흐름
- `src/app/page.tsx`: 클라이언트 컴포넌트. 업로드 드롭존, 상태 카드, 최근 파일 리스트, 시트 선택, 계산 토글(SMP/PPA/REC/PSD), 청크 미리보기, 리포트 생성 UI, CAPEX 간트 편집, 마일스톤 입력 모달 제어.
  - 업로드 시 `/api/rag/upload` 호출 후 처리 상태를 폴링(`NEXT_PUBLIC_RAG_STATUS_BASE`, `NEXT_PUBLIC_RAG_STATUS_KEY`).
  - 최근 파일/삭제/청크 미리보기는 별도 API(`NEXT_PUBLIC_RAG_FILES_API`, `NEXT_PUBLIC_RAG_DELETE_API`, `NEXT_PUBLIC_RAG_CHUNKS_API`)를 사용하며 key 헤더 필요.
  - CAPEX 간트 기본 값은 `src/data/capex-gantt.json`을 매핑해 그룹/작업 구조와 주차 범위를 계산, UI에서 직접 수정 가능.
  - 리포트 생성 버튼은 현재 콘솔 출력만 수행(실제 API 연계 미구현).
- `src/components/PVSolarExcelPreprocessModal.tsx`: 업로드 보조 모달. Excel 업로드 → `/api/preprocess/pv-solar/excel/analyze`로 시트/미리보기 수집 → AI 모드 시 `/api/preprocess/pv-solar/excel/recommend` 호출해 추천 컬럼·인사이트·메타데이터 표시. 에너지 단위/정규화/메타데이터 설정과 행 필터, 파일 분할 옵션을 수집해 `onProcess`로 반환.
- 스타일: `src/app/page.module.css`, `src/app/globals.css` 사용. Tailwind 미사용, CSS 모듈 기반.

## 서버(API) 라우트
- `src/app/api/rag/upload/route.ts`: 업로드 유효성 검사(확장자, 50MB), 기본 파라미터 강제, PV Solar 특화 옵션(FormData) 구성 후 백엔드 `getApiServerUrl()/rag/upload`로 프록시. 실패 시 모의 성공 응답 반환. API Key(`NEXT_PUBLIC_RAG_STATUS_KEY`)와 쿠키의 Bearer 토큰을 헤더에 전달.
- `src/app/api/preprocess/pv-solar/excel/analyze/route.ts`: `exceljs`로 모든 시트 로딩 → 헤더/미리보기(10행) 추출 → 시트별 메타 반환.
- `src/app/api/preprocess/pv-solar/excel/recommend/route.ts`: 전달된 시트 메타와 원본 파일로 PV Solar 점수화 후 최적 시트/컬럼/인사이트 추천. AI 호출(`/callai`, provider: ollama, model: airun-chat:latest)로 JSON 응답 시도, 실패 시 규칙 기반 대체.

## 설정/데이터
- `src/config/serverConfig.ts`: API/RAG 기본 URL을 환경변수로 해석. 서버/클라이언트 모두 참조.
- `src/config/milestones.ts`: CAPEX 마일스톤 템플릿 정의.
- `src/data/capex-gantt.json`: 기본 간트 주차 데이터.
- 루트에 `.env`, `.env.local` 존재(민감정보 가능). `.next/`와 `node_modules/`가 포함되어 빌드 산출물이 남아 있음.

## 스크립트/도구
- npm 스크립트: `dev`, `build`, `start`, `lint`, `type-check`.
- Python 스크립트(`scripts/`)와 샘플 문서(`pv_solar_sample_docs/`)가 포함되어 있으나 Next 앱 실행과는 별개.

## 주의/확인사항
- 외부 API 키와 URL 환경변수가 없으면 업로드/상태 조회/미리보기 호출이 실패하거나 모의 응답으로 대체될 수 있음.
- 현재 리포트 생성 동작은 콘솔 로깅만 하고 있어 백엔드 트리거가 필요.
- `.next` 산출물은 컨텍스트를 불필요하게 키울 수 있으므로 Docker/배포 시 `.dockerignore`에 포함해야 함.
- Excel/파일 업로드는 Next Edge가 아닌 Node 런타임에서 동작해야 함(기본은 `nodejs` 라우트).

## 빠른 검증 커맨드
- `npm install` (의존성 설치)
- `npm run lint` (Next ESLint)
- `npm run type-check`
- `npm run dev` 후 http://localhost:3000 에서 업로드 흐름 확인
