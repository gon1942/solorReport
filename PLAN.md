### 1) 요구사항 요약
- 기존 `pv_solar/capex_readme.md` 분석 문서를 한국어로 재작성하고, 팀이 바로 참고 가능한 형태로 다듬는다.

### 2) 영향 범위(파일/모듈)
- P0: `pv_solar/capex_readme.md`
- P1: `PLAN.md`
- P2: 없음

### 3) 실행 계획(단계)
- Step 1:
  - 변경 파일: `pv_solar/capex_readme.md`
  - 변경 내용: 영문 분석 내용을 한국어로 전체 번역하고, 섹션 구조(목적/흐름/함수/입출력/리스크/개선)를 유지해 재정렬한다.
  - 검증: 문서 내 모든 핵심 함수와 실행 단계가 한국어로 누락 없이 설명되는지 확인한다.
- Step 2:
  - 변경 파일: `pv_solar/capex_readme.md`
  - 변경 내용: 운영 가이드 관점(실행 명령, 환경변수 우선순위, 유지보수 포인트)을 명확히 정리한다.
  - 검증: 신규 입문자가 문서만 읽고 실행/수정 지점을 파악할 수 있는지 점검한다.

### 3-1) 코드 스니펫 예시 (Pseudocode):
```python
data_dir = _resolve_capex_data_dir()
output_path = _resolve_capex_output_path(default_output_dir)
content_list = build_content_list_from_api(data_dir=data_dir)
generate_with_template(content_list, output_path)
for content in ordered_content:
    apply_layout_specific_postprocess(content.layout)
prune_slides(output_path, keep=len(content_list))
```

### 4) 리스크 & 롤백
- 리스크: 번역 과정에서 기술 용어가 원문의 의도와 다르게 해석될 수 있음.
- 롤백(명령/방법): `pv_solar/capex_readme.md`를 직전 커밋 버전으로 복원.

### 5) 완료 기준(Definition of Done)
- `pv_solar/capex_readme.md`가 한국어로 전면 정리된다.
- 함수 책임, 입출력, 환경변수, 리스크, 개선안이 문서에 유지된다.
- 실행 예시 명령이 한국어 설명과 함께 제공된다.
