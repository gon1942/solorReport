# capex_pptx.py 분석 문서

## 1) 문서 목적

`pv_solar/capex_pptx.py`는 CAPEX 보고용 PPTX를 생성하고, 생성된 슬라이드를 레이아웃별로 후처리하는 실행 스크립트다.

이 파일은 역할이 크게 두 가지다.

- 실행 엔트리포인트(`test_airun_pptx`) 역할
- 레이아웃별 후처리 유틸 함수 모음 역할

즉, "데이터 기반 콘텐츠 생성"과 "템플릿 기반 시각 보정"이 한 파일에 결합된 오케스트레이션 스크립트다.

---

## 2) 전체 동작 흐름

메인 실행 경로는 `test_airun_pptx()`다.

1. 입력/출력 경로를 결정한다.
   - 입력: `CAPEX_DATA_DIR` 우선, 없으면 `../data`
   - 출력: `CAPEX_OUTPUT_PATH` 우선, 없으면 `CAPEX_OUTPUT_DIR/result.pptx`, 둘 다 없으면 `../resuot_output/result.pptx`
2. API 형태 JSON으로부터 콘텐츠 리스트를 생성한다.
   - `build_content_list_from_api(data_dir=...)`
3. 템플릿 기반으로 1차 PPTX를 생성한다.
   - `PPTXGenerator(...).generate_with_template(content_list, output_path)`
4. 레이아웃 타입별 후처리를 적용한다.
   - 승인 요청 수치 치환
   - `cod_pipeline`, `cod_pipeline_phase`
   - `main_agreements`
   - `equipment_procurement_case`
5. 남은 템플릿 기본 슬라이드를 제거해 최종 장수를 맞춘다.
   - `prune_slides(output_path, keep=len(content_list))`
6. 생성 요약(경로, 장수, 로고 개수)을 출력한다.

---

## 3) 주요 함수별 책임

### `_apply_approval_request_replacements(...)`
- 승인 요청 슬라이드(기본 index 3)의 텍스트를 지정 값으로 치환한다.
- KRW/EUR/용량 같은 하드코딩 문자열 교체에 사용한다.
- 결과는 원본 PPTX에 즉시 저장된다.

### `prune_slides(...)`
- `keep` 장수보다 많은 뒤쪽 슬라이드를 삭제한다.
- 내부적으로 `prs.slides._sldIdLst`를 직접 조작한다.

### `apply_main_agreements_slide(...)`
- `main_agreements` 레이아웃 후처리 함수다.
- 동작 방식:
  - placeholder 토큰(`{{agreement_1}}`, `{...}`, `<<...>>`) 치환
  - 토큰 실패 시 `Current Status`/`Next Steps` 컬럼 텍스트 박스를 찾아 값 채우기
- 제목을 제외한 텍스트 폰트를 9pt로 통일한다.

### `apply_cod_pipeline_slide(...)`
- `cod_pipeline`, `cod_pipeline_phase` 레이아웃을 처리한다.
- 제목/서브타이틀/본문 텍스트 재구성, 위치 보정, 폰트/색상 규칙을 적용한다.
- 표 스타일링이 가장 복잡한 구간이며 다음을 포함한다.
  - XML 레벨 경계선/배경색 지정(`lxml.etree` 사용)
  - 헤더/본문 정렬, 열 폭, 행 높이 규칙 적용
  - 병합 셀(SPV, Phase, Sub Total, TOTAL) 구성
  - 최종적으로 표 내 폰트 9pt 통일

### `apply_equipment_procurement_title(...)`
- 장비 조달 슬라이드에서 제목 영역을 찾아 강제 적용한다.
- 제목 shape를 못 찾으면 새 textbox를 만든다.
- 필요 시 본문 bullet도 content 기반으로 주입한다.

### `_build_project_detail_values(...)`
- `general.json`, `opex_year1.json`, `equipment_vendors.json`를 읽어 파생값을 계산한다.
- 주요 계산 항목:
  - location/cod
  - dc_wp
  - total_quote
  - equip_budget
  - delta
  - cubicle amount, won/wp
  - cubicle 업체 목록

### `apply_equipment_project_detail_values(...)`
- 프로젝트 상세 영역(슬라이드 내 텍스트/표 영역)에 계산값을 주입한다.
- 라벨 텍스트를 기준으로 위치를 추정해 값 textbox를 오버레이한다.
- 특정 하단 영역의 폰트를 6pt로 일괄 정규화한다.

### `_resolve_capex_data_dir()` / `_resolve_capex_output_path(...)`
- 환경변수 기반 경로 해석 유틸 함수다.
- 실행 환경(로컬/서버)에서 경로를 유연하게 바꿀 수 있게 해준다.

### `test_airun_pptx()`
- 전체 생성 파이프라인을 오케스트레이션하는 메인 함수다.
- 콘텐츠 정렬, 레이아웃별 후처리, 슬라이드 정리, 요약 로그 출력까지 수행한다.

---

## 4) 입력/출력 및 실행 조건

### 필수 라이브러리
- `python-pptx`
- `lxml`

### 입력 파일(데이터 디렉터리 기준)
- `general.json`
- `opex_year1.json`
- `equipment_vendors.json` (선택적 override 용도)

### 출력 파일
- 최종 PPTX (`result.pptx`)
- 기본 위치: `resuot_output/result.pptx`

---

## 5) 구현 특성 및 유지보수 포인트

1. 템플릿 의존성이 높다.
   - 텍스트 매칭/위치 추론이 많아 템플릿 문구가 바뀌면 후처리가 조용히 스킵될 수 있다.

2. run 단위 텍스트 치환 한계가 있다.
   - placeholder가 run 여러 개로 쪼개져 있으면 치환 누락 가능성이 있다.

3. XML 직접 조작 코드가 포함되어 있다.
   - 표현력은 높지만 템플릿 버전 변화에 민감하고 디버깅 난도가 높다.

4. 좌표/크기 하드코딩이 많다.
   - `Inches(...)` 기반 고정값이 많아서 마스터/비율 변경 시 레이아웃 깨짐 위험이 있다.

5. 파일 말단에 정리 대상 코드가 있다.
   - `if __name__ == "__main__":` 아래에 사용되지 않는 `_as_dict` 선언이 남아 있다.

---

## 6) 개선 제안 (저위험, 점진적)

1. 후처리 결과 로깅 강화
   - 레이아웃별로 "적용됨/스킵됨"과 매칭 건수를 남기면 운영 추적이 쉬워진다.

2. 검증 리포트 모드 추가
   - 치환 건수, 스킵 섹션, 최종 슬라이드 수를 JSON/MD로 출력하면 회귀 확인이 쉬워진다.

3. 상수 분리
   - 라벨 문자열/좌표/폰트 설정을 별도 constants 모듈로 분리하면 유지보수성이 올라간다.

4. placeholder 치환 보강
   - run 단위 실패 시 paragraph 단위 재조합 치환 fallback을 추가하면 안정성이 높아진다.

5. 사전 무결성 체크
   - 실행 전 필수 JSON 키/슬라이드 인덱스 점검을 하면 후반부 실패를 줄일 수 있다.

---

## 7) 실행 방법

기본 실행:

```bash
python pv_solar/capex_pptx.py
```

환경변수로 경로 오버라이드:

```bash
CAPEX_DATA_DIR=/path/to/data CAPEX_OUTPUT_PATH=/path/to/result.pptx python pv_solar/capex_pptx.py
```

운영 관점에서 이 스크립트는 "코드 정확성"만큼 "템플릿 안정성"이 중요한 구조다.
