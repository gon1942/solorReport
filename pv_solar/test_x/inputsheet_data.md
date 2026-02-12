당신은 CAPEX/DEVEX Approval 문서 작성자입니다.
아래 InputSheet에서 추출한 JSON을 근거로, PPT 슬라이드에 바로 꽂을 수 있는 “구조화 JSON”만 출력하십시오.
추측하지 말고, 값이 없으면 null 또는 "TBD"로 표기하십시오.
통화/단위는 InputSheet에 표기된 내용을 우선하며, 없으면 "TBD"로 두십시오.

[요청 1] Executive Summary 슬라이드용 표 데이터 생성
- 출력은 2열 테이블: (항목, 값)
- 섹션 헤더 행을 포함 (예: General / Schedule / Capacity / CAPEX / Development cost)
- 최소 포함 항목 예:
  - Seller/developer
  - Project name
  - Country, Region
  - Technology
  - Project Status
  - RTB date, COD date
  - Total Capacity DC, Total Capacity AC, DC/AC ratio
  - CAPEX Total (있으면)
  - Development cost Total (있으면)

[요청 2] Approval request 슬라이드 문구 생성
- Expected return 문구 (있으면 IRR/Return 관련 값 활용)
- CAPEX 승인 문구 (총 CAPEX, 통화 포함)
- 없으면 "TBD"로 둠

[출력 형식: 반드시 아래 JSON 스키마로만]
{
  "executive_summary": {
    "title": "Executive Summary",
    "table": [
      {"type": "section", "left": "General", "right": ""},
      {"type": "row", "left": "Seller/developer", "right": "..."},
      {"type": "row", "left": "Project name", "right": "..."},
      {"type": "section", "left": "Schedule", "right": ""},
      {"type": "row", "left": "RTB date", "right": "..."},
      {"type": "row", "left": "COD date", "right": "..."},
      {"type": "section", "left": "Capacity", "right": ""},
      {"type": "row", "left": "Total Capacity DC", "right": "..."},
      {"type": "row", "left": "Total Capacity AC", "right": "..."},
      {"type": "row", "left": "DC/AC ratio", "right": "..."},
      {"type": "section", "left": "CAPEX", "right": ""},
      {"type": "row", "left": "CAPEX Total", "right": "..."},
      {"type": "section", "left": "Development cost", "right": ""},
      {"type": "row", "left": "Development cost Total", "right": "..."}
    ]
  },
  "approval_request": {
    "title": "Approval request",
    "expected_return": {
      "label": "Expected return",
      "text": "..."
    },
    "capex": {
      "label": "CAPEX",
      "text": "..."
    }
  },
  "trace": {
    "used_fields": [
      {"label": "...", "value": "...", "section": "..."}
    ]
  }
}

[InputSheet JSON]
{
  "executive_summary": {
    "title": "Executive Summary",
    "table": [
      {"type":"section","left":"General","right":""},
      {"type":"row","left":"Seller/developer","right":"TBD"},
      {"type":"row","left":"Project name","right":"Project 1"},

      {"type":"section","left":"Schedule","right":""},
      {"type":"row","left":"RTB date","right":"2025-11-01"},
      {"type":"row","left":"COD date","right":"2026-06-30"},

      {"type":"section","left":"Capacity","right":""},
      {"type":"row","left":"Total Capacity DC","right":"0.817 MWp"},
      {"type":"row","left":"Total Capacity AC","right":"0.681 MWac"},
      {"type":"row","left":"DC/AC ratio","right":"1.20"},

      {"type":"section","left":"CAPEX","right":""},
      {"type":"row","left":"CAPEX Total","right":"TBD"},

      {"type":"section","left":"Development cost","right":""},
      {"type":"row","left":"Development cost Total","right":"TBD"}
    ]
  },
  "approval_request": {
    "title": "Approval request",
    "expected_return": {
      "label":"Expected return",
      "text":"Approve the expected return based on the latest revenues and capex. Approval criterion: Portfolio blended Equity IRR of 9.64% (≈9.65%). Average IRR (auxiliary, unweighted average of project IRRs): 8.95%."
    },
    "capex": {
      "label":"CAPEX",
      "text":"Approval of total CAPEX (TBD). Please input CAPEX Total."
    }
  }
}
