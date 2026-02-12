# Runbook

## 개발 환경 실행

```bash
pnpm install
pnpm dev
```

- 기본 개발 서버: `http://localhost:3000`

## 빌드 및 실행

```bash
pnpm build
pnpm start
```

## 타입 체크

```bash
pnpm type-check
```

## 주요 화면 점검

1. `/` 접속
2. 파일 업로드 후 분석/미리보기 확인
3. 카테고리 `CAPEX` 선택
4. Phase 1 / Phase 2 간트 편집
5. 간트 저장 버튼 클릭 후 새로고침하여 값 유지 확인

## 환경 변수 체크

- `NEXT_PUBLIC_API_SERVER_URL`
- `NEXT_PUBLIC_RAG_API_SERVER`
- `NEXT_PUBLIC_RAG_STATUS_KEY`
- `NEXT_PUBLIC_RAG_STATUS_BASE`
- `NEXT_PUBLIC_RAG_FILES_API`
- `NEXT_PUBLIC_RAG_DELETE_API`
- `NEXT_PUBLIC_RAG_CHUNKS_API`
- `NEXT_PUBLIC_RAG_DEFAULT_USER_ID`
- `REDIS_URL` 또는 `NEXT_PUBLIC_REDIS_URL`

## 장애 시 기본 확인

1. 브라우저 콘솔 에러 확인
2. API route 응답 코드 및 메시지 확인
3. 환경 변수 누락 여부 확인
4. Redis 연결 설정 유무 확인 (`/api/planner/status`는 Redis 미연결 시 mock 반환)

## Notes

- 운영/배포용 상세 절차는 추후 추가 예정입니다.
