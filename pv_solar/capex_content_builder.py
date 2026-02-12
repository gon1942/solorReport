# -*- coding: utf-8 -*-
"""CAPEX content/data builders split from capex_pptx."""

import os
import json
import calendar
import re
from pathlib import Path
from typing import Any, List, Optional
from datetime import datetime, time

try:
    from solar_pptx import SlideContent
except ModuleNotFoundError:
    from pv_solar.solar_pptx import SlideContent

DEFAULT_API_DOCS = [doc.strip() for doc in os.getenv("HAMONIZE_API_DOCUMENTS", "test2.xlsx").split(",") if doc.strip()]
LEASE_PREPAY_RATE = 35000.0
LEASE_ANNUAL_RATE = 40000.0
LEASE_FX_RATE = 1470.0


def _resolve_data_dir(data_dir: Optional[Path | str] = None) -> Path:
    if data_dir:
        return Path(data_dir).resolve()
    env_data_dir = os.getenv("CAPEX_DATA_DIR", "").strip()
    if env_data_dir:
        return Path(env_data_dir).resolve()
    return (Path(__file__).resolve().parent.parent / "data").resolve()


def _resolve_gantt_phase_path(data_dir: Path, phase_key: str) -> Path:
    project_root = Path(__file__).resolve().parent.parent
    candidates = [
        (data_dir.parent / "data_gantt" / f"{phase_key}.json").resolve(),
        (data_dir.parent.parent / "data_gantt" / f"{phase_key}.json").resolve(),
        (project_root / "data_gantt" / f"{phase_key}.json").resolve(),
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]

def _parse_phase_bullets(text: str) -> tuple[list[str], list[str]]:
    """Executive Summary 텍스트에서 Phase 1/2 불릿을 분리해 반환."""

    phase1: list[str] = []
    phase2: list[str] = []
    current: Optional[str] = None

    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        lower = stripped.lower()
        if lower.startswith("bolt #2 - phase 1"):
            current = "phase1"
            continue
        if lower.startswith("bolt #2 - phase 2"):
            current = "phase2"
            continue
        if stripped.startswith("•"):
            bullet = stripped.lstrip("•").strip()
            if current == "phase1":
                phase1.append(bullet)
            elif current == "phase2":
                phase2.append(bullet)

    return phase1, phase2

def _inject_region_line(lines: list[str], region_line: str, *, max_items: int = 7) -> list[str]:
    """지역 요약을 1순위로 넣되 기존 지역 관련 불릿은 제거한다."""

    filtered = [ln for ln in lines if not any(key in ln for key in ("지역", "위치", "location", "Location"))]
    merged = [region_line] + filtered
    return merged[:max_items]

def _summarize_regions(projects: list[dict]) -> str:
    from collections import Counter

    regions: list[str] = []
    for p in projects:
        region = p.get("region")
        if region is None:
            continue
        if not isinstance(region, str):
            region = str(region)
        region = region.strip()
        if region:
            regions.append(region)

    if not regions:
        return "지역 정보 없음"

    counter = Counter(regions)
    parts = [f"{region}({count})" for region, count in sorted(counter.items())]
    return ", ".join(parts)

def _cod_window(projects: list[dict]) -> str:
    dates = sorted({p["cod"] for p in projects if p.get("cod")})
    if not dates:
        return "데이터 확인 필요"
    return f"{dates[0]} ~ {dates[-1]}" if len(dates) > 1 else dates[0]

def _to_date_str(val: Any) -> Optional[str]:
    if val is None or (isinstance(val, float) and str(val) == "nan"):
        return None
    if isinstance(val, datetime):
        return val.date().isoformat()
    try:
        dt = datetime.fromisoformat(str(val).replace(" ", "T"))
        return dt.date().isoformat()
    except Exception:
        try:
            dt = datetime.strptime(str(val).split()[0], "%Y-%m-%d")
            return dt.date().isoformat()
        except Exception:
            return str(val)

def _clean_val(val: Any) -> Optional[Any]:
    if val is None:
        return None
    if isinstance(val, time):
        # 00:00:00 등은 빈 값 취급
        return None if val == time(0, 0, 0) else val.isoformat()
    if isinstance(val, datetime):
        return val.isoformat()
    if isinstance(val, float) and str(val) == "nan":
        return None
    if isinstance(val, str):
        v = val.strip()
        if not v or v == "00:00:00":
            return None
        return v
    return val

def _province_from_region(region: Optional[str]) -> Optional[str]:
    if not region:
        return None
    if "," in region:
        return region.split(",")[-1].strip()
    return region.strip()

def _load_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:  # pylint: disable=broad-except
        print(f"⚠️ JSON 로드 실패: {path} ({exc})")
        return {}

def _to_float(val: Any) -> float:
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        cleaned = val.strip().replace(",", "")
        if cleaned == "":
            return 0.0
        try:
            return float(cleaned)
        except Exception:
            return 0.0
    return 0.0

def _to_optional_float(val: Any) -> Optional[float]:
    if val is None:
        return None
    if isinstance(val, float) and str(val) == "nan":
        return None
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        cleaned = val.strip().replace(",", "")
        if cleaned == "":
            return None
        try:
            return float(cleaned)
        except Exception:
            return None
    return None

def _format_optional_number(val: Any, decimals: int = 2) -> str:
    num = _to_optional_float(val)
    if num is None:
        return ""
    return f"{num:,.{decimals}f}"

def _format_cod_month(val: Any) -> str:
    if val is None:
        return ""
    if isinstance(val, datetime):
        return f"{calendar.month_name[val.month]} {val.year}"
    if isinstance(val, str):
        cleaned = val.strip()
        if not cleaned:
            return ""
        try:
            dt = datetime.fromisoformat(cleaned.replace(" ", "T"))
            return f"{calendar.month_name[dt.month]} {dt.year}"
        except Exception:
            try:
                dt = datetime.strptime(cleaned.split()[0], "%Y-%m-%d")
                return f"{calendar.month_name[dt.month]} {dt.year}"
            except Exception:
                return ""
    return ""

def _extract_rec_tariff_krw_per_kwh(data_dir: Path, default: float = 189.0) -> float:
    pattern = re.compile(r"KRW\s*([0-9][0-9,]*)\s*/\s*kWh", re.IGNORECASE)
    for json_path in sorted(data_dir.glob("*.json")):
        try:
            text = json_path.read_text(encoding="utf-8")
        except Exception:
            continue
        match = pattern.search(text)
        if match:
            try:
                return float(match.group(1).replace(",", ""))
            except Exception:
                continue
    return default

def _build_route_status_text(data_dir: Path, total_capacity_kwp: float) -> str:
    rec_tariff = _extract_rec_tariff_krw_per_kwh(data_dir)
    dg_mw = total_capacity_kwp / 1000.0 if total_capacity_kwp > 0 else 0.0

    rec_tariff_text = f"{int(rec_tariff):,}" if abs(rec_tariff - int(rec_tariff)) < 1e-9 else f"{rec_tariff:,.2f}"
    dg_mw_text = f"{dg_mw:.1f}" if dg_mw > 0 else "0.0"

    return (
        f"▪ SBP has confirmed a fixed REC tariff of KRW {rec_tariff_text}/kWh for a {dg_mw_text}MW DG portfolio.\n"
        "▪ OR is currently screening and reviewing candidate projects to be proposed together with the GV PPA.\n"
        "▪ Upon final selection, the BOLT#2 projects will be formally submitted to SBP, with no change in the agreed pricing."
    )

def _build_route_to_market_table(data_dir: Path) -> tuple[list[str], list[list[str]]]:
    general_path = data_dir / "general.json"
    general = _load_json(general_path)

    headers = [
        "SPV",
        "Phase #",
        "No.",
        "PJT Name",
        "Region",
        "Capacity\n(kWp)",
        "Specific\nProduction\n(kWh/kWp/yr)",
        "REC Tariff\n(KRW/kWh)",
        "Offtaker",
        "Status",
    ]

    rows: list[list[str]] = []
    projects = general.get("projects") or []
    total_capacity_kwp = 0.0
    metrics: list[tuple[Optional[float], Optional[float]]] = []

    for idx, project in enumerate(projects, start=1):
        inputs = project.get("general_inputs") or {}
        production = project.get("production") or {}
        production_block = production.get("Production") or {}

        project_name = inputs.get("Project name") or ""
        region = inputs.get("Region") or ""
        status = inputs.get("Project Status") or ""
        capacity_mwp = inputs.get("Total Capacity DC")
        capacity_kwp = _to_optional_float(capacity_mwp)
        if capacity_kwp is not None:
            capacity_kwp *= 1000
            total_capacity_kwp += capacity_kwp
        specific_prod = production_block.get("Specific Production")
        specific_prod_num = _to_optional_float(specific_prod)

        rows.append([
            "BOLT#2" if idx == 1 else "",
            "1" if idx == 1 else ("2" if idx == 8 else ""),
            str(idx),
            str(project_name),
            str(region),
            _format_optional_number(capacity_kwp, 2),
            _format_optional_number(specific_prod, 2),
            "",
            "",
            "",
        ])
        metrics.append((capacity_kwp, specific_prod_num))

    def _calc_subtotal(start_idx: int, end_idx: int) -> tuple[float, Optional[float]]:
        subtotal_capacity = 0.0
        subtotal_specific_weighted = 0.0
        subtotal_specific_sum = 0.0
        subtotal_specific_count = 0

        for capacity_kwp, specific_prod_num in metrics[start_idx:end_idx]:
            if capacity_kwp is not None:
                subtotal_capacity += capacity_kwp
            if specific_prod_num is not None:
                subtotal_specific_count += 1
                subtotal_specific_sum += specific_prod_num
                if capacity_kwp is not None and capacity_kwp > 0:
                    subtotal_specific_weighted += specific_prod_num * capacity_kwp

        if subtotal_capacity > 0:
            subtotal_specific = subtotal_specific_weighted / subtotal_capacity
        elif subtotal_specific_count > 0:
            subtotal_specific = subtotal_specific_sum / subtotal_specific_count
        else:
            subtotal_specific = None

        return subtotal_capacity, subtotal_specific

    if rows:
        phase1_end = min(7, len(rows))
        cap1, spec1 = _calc_subtotal(0, phase1_end)
        rows.insert(phase1_end, [
            "",
            "",
            "",
            "SubTotal",
            "",
            _format_optional_number(cap1, 2),
            _format_optional_number(spec1, 2),
            "",
            "",
            "",
        ])

        if len(projects) > 7:
            cap2, spec2 = _calc_subtotal(7, len(projects))
            phase2_insert_idx = len(rows)
            rows.insert(phase2_insert_idx, [
                "",
                "",
                "",
                "SubTotal",
                "",
                _format_optional_number(cap2, 2),
                _format_optional_number(spec2, 2),
                "",
            "",
            "",
        ])

    status_text = _build_route_status_text(data_dir, total_capacity_kwp)
    if rows:
        rows[0][9] = status_text

    if rows:
        rows.append([
            "",
            "",
            "",
            "TOTAL",
            "",
            _format_optional_number(total_capacity_kwp, 2),
            "",
            "",
            "",
            "",
        ])

    return headers, rows

def _build_cod_pipeline_table(data_dir: Path) -> tuple[list[str], list[list[str]]]:
    general_path = data_dir / "general.json"
    general = _load_json(general_path)

    headers = [
        "SPV",
        "Phase #",
        "No.",
        "PJT Name",
        "Region",
        "Capacity\n(kWp)",
        "COD",
    ]

    rows: list[list[str]] = []
    total_capacity_kwp = 0.0
    phase1_capacity_kwp = 0.0
    phase2_capacity_kwp = 0.0
    projects = general.get("projects") or []
    current_year = datetime.now().year
    for idx, project in enumerate(projects, start=1):
        inputs = project.get("general_inputs") or {}
        project_name = inputs.get("Project name") or ""
        region = inputs.get("Region") or ""
        capacity_mwp = inputs.get("Total Capacity DC")
        capacity_kwp = _to_optional_float(capacity_mwp)
        if capacity_kwp is not None:
            capacity_kwp *= 1000
            total_capacity_kwp += capacity_kwp
            if idx <= 7:
                phase1_capacity_kwp += capacity_kwp
            else:
                phase2_capacity_kwp += capacity_kwp

        phase_label = "1" if idx == 1 else ("2" if idx == 8 else "")
        cod_month = ""
        if idx == 1:
            cod_month = f"May {current_year}"
        elif idx == 8:
            cod_month = f"June {current_year}"

        rows.append([
            "BOLT#2" if idx == 1 else "",
            phase_label,
            str(idx),
            str(project_name),
            str(region),
            _format_optional_number(capacity_kwp, 2),
            cod_month,
        ])

    if rows:
        phase1_end = min(7, len(rows))
        rows.insert(phase1_end, [
            "",
            "",
            "",
            "Sub Total",
            "",
            _format_optional_number(phase1_capacity_kwp, 2),
            "",
        ])

        if len(projects) > 7:
            rows.append([
                "",
                "",
                "",
                "Sub Total",
                "",
                _format_optional_number(phase2_capacity_kwp, 2),
                "",
            ])

    if rows:
        rows.append([
            "",
            "",
            "",
            "TOTAL",
            "",
            _format_optional_number(total_capacity_kwp, 2),
            "",
        ])

    return headers, rows


def _build_technical_solution_text(data_dir: Path) -> str:
    """Technical Solution 본문 텍스트를 데이터 기반으로 구성한다."""
    general_path = data_dir / "general.json"
    general = _load_json(general_path)
    projects = general.get("projects") or []

    panel_names: list[str] = []
    structure_types: list[str] = []
    for project in projects:
        inputs = project.get("general_inputs") or {}
        panel_name = str(inputs.get("WTG / Panels name") or "").strip()
        if panel_name and panel_name.upper() != "N/A" and panel_name not in panel_names:
            panel_names.append(panel_name)

        structure = str(inputs.get("Transaction structure") or "").strip()
        if structure and structure not in structure_types:
            structure_types.append(structure)

    panel_desc = panel_names[0] if panel_names else ""
    structure_desc = ", ".join(structure_types) if structure_types else ""

    placeholders = {
        "{{panel_desc}}": panel_desc,
        "{{structure_desc}}": structure_desc,
    }

    config_path = data_dir / "common" / "technical_solution.json"
    loaded = _load_json(config_path)
    summary_title = str((loaded.get("summary_title") if isinstance(loaded, dict) else "") or "Equipment Configuration Summary").strip()
    sections_raw = loaded.get("sections")
    sections: list[Any] = sections_raw if isinstance(sections_raw, list) else []

    def _apply_placeholders(text: str) -> str:
        value = text
        for key, rep in placeholders.items():
            value = value.replace(key, rep)
        return value

    lines = [summary_title]

    # technical_solution.json이 유효하게 제공되면 해당 데이터만 사용
    if sections:
        for section in sections:
            if not isinstance(section, dict):
                continue
            title = str(section.get("title") or "").strip()
            if title:
                lines.append(title)
            bullets_raw = section.get("bullets")
            bullets: list[Any] = bullets_raw if isinstance(bullets_raw, list) else []
            for bullet in bullets:
                txt = _apply_placeholders(str(bullet or "").strip())
                if txt:
                    lines.append(f"- {txt}")
        return "\n".join(lines)

    # technical_solution.json이 없을 때는 general.json에서 읽은 항목만 최소 구성
    if panel_desc:
        lines.append("Solar Panels")
        lines.append(f"- {panel_desc}")
    if structure_desc:
        lines.append("Structure")
        lines.append(f"- {structure_desc}")

    return "\n".join(lines)

def _build_permits_table(data_dir: Path, split_index: int = 7) -> tuple[list[str], list[list[str]]]:
    general_path = data_dir / "general.json"
    general = _load_json(general_path)
    permit_overrides = _load_dg_pv_mna_permit_overrides(data_dir)

    headers = [
        "SPV",
        "Phase #",
        "No.",
        "PJT Name",
        "Capacity\n(kWp)",
        "EBL\nCompletion",
        "EBL\nValidity & Status",
        "EIA\nCompletion",
        "EIA\nValidity & Status",
        "DAP\nCompletion",
        "DAP\nValidity & Status",
        "Hanjeon PPA (GIA)\nApproval & Status",
    ]

    rows: list[list[str]] = []
    projects = general.get("projects") or []
    total_capacity_kwp = 0.0
    phase_capacity = [0.0, 0.0]

    def _as_text(value: Any) -> str:
        if value is None:
            return ""
        text = str(value).strip()
        return "" if text.lower() in {"", "none", "nan"} else text

    def _pick_first(*values: Any) -> str:
        for value in values:
            text = _as_text(value)
            if text:
                return text
        return ""

    for idx, project in enumerate(projects, start=1):
        inputs = project.get("general_inputs") or {}
        permits = project.get("permits") or {}
        hanjeon_block = permits.get("Hanjeon") or permits.get("hanjeon") or {}
        project_name = inputs.get("Project name") or ""
        capacity_mwp = inputs.get("Total Capacity DC")
        capacity_kwp = _to_optional_float(capacity_mwp)
        if capacity_kwp is not None:
            capacity_kwp *= 1000
            total_capacity_kwp += capacity_kwp

        phase = 1 if idx <= split_index else 2
        phase_capacity[phase - 1] += capacity_kwp or 0.0

        ebl_completion = _pick_first(
            inputs.get("EBL Completion"),
            permits.get("EBL Completion"),
            (permits.get("EBL") or {}).get("Completion") if isinstance(permits.get("EBL"), dict) else None,
            "",
        )
        ebl_validity = _pick_first(
            inputs.get("EBL Validity & Status"),
            permits.get("EBL Validity & Status"),
            (permits.get("EBL") or {}).get("Validity & Status") if isinstance(permits.get("EBL"), dict) else None,
            "",
        )
        eia_completion = _pick_first(
            inputs.get("EIA Completion"),
            permits.get("EIA Completion"),
            (permits.get("EIA") or {}).get("Completion") if isinstance(permits.get("EIA"), dict) else None,
            "",
        )
        eia_validity = _pick_first(
            inputs.get("EIA Validity & Status"),
            permits.get("EIA Validity & Status"),
            (permits.get("EIA") or {}).get("Validity & Status") if isinstance(permits.get("EIA"), dict) else None,
            "",
        )
        dap_completion = _pick_first(
            inputs.get("DAP Completion"),
            permits.get("DAP Completion"),
            (permits.get("DAP") or {}).get("Completion") if isinstance(permits.get("DAP"), dict) else None,
            "",
        )
        dap_validity = _pick_first(
            inputs.get("DAP Validity & Status"),
            permits.get("DAP Validity & Status"),
            (permits.get("DAP") or {}).get("Validity & Status") if isinstance(permits.get("DAP"), dict) else None,
            "",
        )
        hanjeon_status = _pick_first(
            inputs.get("Hanjeon PPA (GIA) Approval & Status"),
            inputs.get("GIA Approval & Status"),
            permits.get("Hanjeon PPA (GIA) Approval & Status"),
            permits.get("GIA Approval & Status"),
            hanjeon_block.get("Approval & Status") if isinstance(hanjeon_block, dict) else None,
            "",
        )

        override = permit_overrides.get(idx)
        if isinstance(override, dict):
            ebl_completion = _pick_first(override.get("EBL Completion"), ebl_completion)
            ebl_validity = _pick_first(override.get("EBL Validity & Status"), ebl_validity)
            eia_completion = _pick_first(override.get("EIA Completion"), eia_completion)
            eia_validity = _pick_first(override.get("EIA Validity & Status"), eia_validity)
            dap_completion = _pick_first(override.get("DAP Completion"), dap_completion)
            dap_validity = _pick_first(override.get("DAP Validity & Status"), dap_validity)
            hanjeon_status = _pick_first(override.get("Hanjeon PPA (GIA) Approval & Status"), hanjeon_status)
        elif not eia_completion and not eia_validity:
            eia_completion = "N/A"
            eia_validity = "N/A"

        rows.append([
            "BOLT#2" if idx == 1 else "",
            str(phase),
            str(idx),
            str(project_name),
            _format_optional_number(capacity_kwp, 2),
            ebl_completion,
            ebl_validity,
            eia_completion,
            eia_validity,
            dap_completion,
            dap_validity,
            hanjeon_status,
        ])

        if idx == split_index:
            rows.append([
                "",
                "",
                "",
                "Sub Total",
                _format_optional_number(phase_capacity[0], 2),
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            ])

    if projects:
        rows.append([
            "",
            "",
            "",
            "Sub Total",
            _format_optional_number(phase_capacity[1], 2),
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ])

        rows.append([
            "",
            "",
            "",
            "TOTAL",
            _format_optional_number(total_capacity_kwp, 2),
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ])

    return headers, rows

def _load_dg_pv_mna_permit_overrides(data_dir: Path) -> dict[int, dict[str, str]]:
    permits_path = data_dir / "permits_dg_pv_mna.json"
    mapping_path = data_dir / "project1_10_posn2_mapping.json"

    permits_payload = _load_json(permits_path)
    mapping_payload = _load_json(mapping_path)
    permit_projects = permits_payload.get("projects") if isinstance(permits_payload, dict) else None
    mappings = mapping_payload.get("mappings") if isinstance(mapping_payload, dict) else None

    if not isinstance(permit_projects, list) or not isinstance(mappings, list):
        return {}

    def _as_text(value: Any) -> str:
        if value is None:
            return ""
        text = str(value).strip()
        return "" if text.lower() in {"", "none", "nan"} else text

    posn_lookup: dict[str, dict[str, Any]] = {}
    for item in permit_projects:
        if not isinstance(item, dict):
            continue
        posn2_id = _as_text(item.get("posn2_id"))
        if not posn2_id:
            continue
        posn_lookup[posn2_id.upper()] = item

    overrides: dict[int, dict[str, str]] = {}
    for mapping in mappings:
        if not isinstance(mapping, dict):
            continue
        project_no = mapping.get("project_no")
        posn2_id = _as_text(mapping.get("posn2_id"))
        if not posn2_id:
            continue
        if not isinstance(project_no, int):
            continue

        source_item = posn_lookup.get(posn2_id.upper())
        if not isinstance(source_item, dict):
            continue

        permits_value = source_item.get("permits")
        permits_block: dict[str, Any] = permits_value if isinstance(permits_value, dict) else {}

        ebl_value = permits_block.get("EBL")
        eia_value = permits_block.get("EIA")
        dap_value = permits_block.get("DAP")
        hanjeon_value = permits_block.get("Hanjeon PPA (GIA)")

        ebl: dict[str, Any] = ebl_value if isinstance(ebl_value, dict) else {}
        eia: dict[str, Any] = eia_value if isinstance(eia_value, dict) else {}
        dap: dict[str, Any] = dap_value if isinstance(dap_value, dict) else {}
        hanjeon: dict[str, Any] = hanjeon_value if isinstance(hanjeon_value, dict) else {}

        overrides[project_no] = {
            "EBL Completion": _as_text(ebl.get("raw")),
            "EBL Validity & Status": _as_text(ebl.get("normalized")),
            "EIA Completion": _as_text(eia.get("raw")),
            "EIA Validity & Status": _as_text(eia.get("normalized")),
            "DAP Completion": _as_text(dap.get("raw")),
            "DAP Validity & Status": _as_text(dap.get("normalized")),
            "Hanjeon PPA (GIA) Approval & Status": _as_text(hanjeon.get("raw")),
        }

    return overrides

def _build_epc_status_and_next_steps(data_dir: Path) -> tuple[str, str]:
    general = _load_json(data_dir / "general.json")
    opex = _load_json(data_dir / "opex_year1.json")
    cfg = _ensure_main_agreements_json(data_dir)
    capex_vals = _build_capex_approval_values(data_dir)

    capacities_mw: dict[str, float] = {}
    for project in general.get("projects") or []:
        if not isinstance(project, dict):
            continue
        name = str(project.get("name") or "").strip()
        if not name:
            continue
        inputs = project.get("general_inputs") or {}
        cap_mw = _to_optional_float(inputs.get("Total Capacity DC"))
        if cap_mw is None or cap_mw <= 0:
            continue
        capacities_mw[name] = cap_mw

    total_mw = 0.0
    total_epc_krw = 0.0
    total_contingency_krw = 0.0
    for project in opex.get("projects") or []:
        if not isinstance(project, dict):
            continue
        name = str(project.get("name") or "").strip()
        cap_mw = capacities_mw.get(name)
        if cap_mw is None or cap_mw <= 0:
            continue
        block = project.get("opex_year1") or {}
        bos = _to_float((block.get("BOS") or {}).get("Cost"))
        modules = _to_float((block.get("Modules") or {}).get("Cost"))
        grid = _to_float((block.get("Grid Connection Cost") or {}).get("Cost"))
        contingency = _to_float((block.get("Contingency") or {}).get("Cost"))

        total_mw += cap_mw
        total_epc_krw += (bos + modules + grid)
        total_contingency_krw += contingency

    if total_mw <= 0:
        return (
            "• Bulk EPC contract value is pending (missing CAPEX baseline).\n"
            "• Contractual conditions review is pending.\n"
            "• Contract signature schedule is pending.",
            "• GPK's execution of the agreement shall follow CAPEX approval and the contractor's signing of the EPC contract.",
        )

    epc_krw_per_mw = total_epc_krw / total_mw
    contingency_krw_per_mw = total_contingency_krw / total_mw
    epc_million_krw_per_mw = epc_krw_per_mw / 1_000_000
    contingency_million_krw_per_mw = contingency_krw_per_mw / 1_000_000
    epc_eur_k_per_mw = (epc_krw_per_mw / LEASE_FX_RATE) / 1000

    epc_cfg = cfg.get("epc") if isinstance(cfg, dict) else {}
    if not isinstance(epc_cfg, dict):
        epc_cfg = {}
    review_status = str(epc_cfg.get("review_status") or "Approved by Greenvolt").strip()
    signature_target = str(epc_cfg.get("signature_target") or "TBD").strip()

    contractor_list: list[str] = []
    for item in epc_cfg.get("contractors") or []:
        if not isinstance(item, dict):
            continue
        name = str(item.get("name") or "").strip()
        cnt = _to_optional_float(item.get("project_count"))
        if not name:
            continue
        if cnt is None:
            contractor_list.append(name)
        else:
            contractor_list.append(f"{name} ({int(cnt)} projects)")
    contractors_text = ", ".join(contractor_list) if contractor_list else "EPC contractors"

    review_sentence = review_status.rstrip(".")
    if not review_sentence:
        review_sentence = "All contractual conditions reviewed and approved by Greenvolt"

    current_status = (
        f"• The bulk EPC contract value is set at KRW {epc_million_krw_per_mw:,.0f} million per MW "
        f"(approx. EUR {epc_eur_k_per_mw:,.0f}k), inclusive of contingency KRW {contingency_million_krw_per_mw:,.1f} million per MW.\n"
        f"• {review_sentence}.\n"
        f"• Contracts with {contractors_text} are scheduled for signature in {signature_target}."
    )
    capex_krw_text = capex_vals.get("capex_krw_text") or "CAPEX"
    capex_eur_text = capex_vals.get("capex_eur_text") or "EUR"
    capacity_text = capex_vals.get("capacity_text") or "portfolio"
    next_steps = (
        "• GPK's execution of the agreement shall follow CAPEX approval "
        f"({capex_krw_text} / {capex_eur_text} for {capacity_text}) and the contractor's signing of the EPC contract."
    )
    return current_status, next_steps

def _build_rec_sales_status_and_next_steps(data_dir: Path) -> tuple[str, str]:
    cfg = _ensure_main_agreements_json(data_dir)
    rec_cfg = cfg.get("rec_sales") if isinstance(cfg, dict) else {}
    if not isinstance(rec_cfg, dict):
        rec_cfg = {}

    negotiation_status = str(rec_cfg.get("negotiation_status") or "Under discussion").strip()
    offtaker_name = str(rec_cfg.get("offtaker_name") or "Offtaker to be confirmed").strip()
    alias = str(rec_cfg.get("alias") or "TBD").strip()
    offtaker_type = str(rec_cfg.get("offtaker_type") or "RPS obligor").strip()
    term_years_num = _to_optional_float(rec_cfg.get("term_years"))
    term_years_text = f"{int(term_years_num)}-year" if term_years_num is not None and term_years_num > 0 else "TBD-year"

    price = _to_optional_float(rec_cfg.get("price_krw_per_kwh"))
    if price is None or price <= 0:
        price = _extract_rec_tariff_krw_per_kwh(data_dir)
    price_text = f"{price:,.0f}" if price is not None else "TBD"

    offer_valid_until = str(rec_cfg.get("offer_valid_until") or f"end of {datetime.now().year}").strip()
    management_approval = str(rec_cfg.get("management_approval") or "pending").strip()

    current_status = (
        f"• {negotiation_status} with {offtaker_name} ({alias}, {offtaker_type}) for a {term_years_text} "
        f"REC Sales Agreement at a fixed price of KRW {price_text}/kWh.\n"
        f"• The price offer remains valid until {offer_valid_until}; management approval is {management_approval} "
        "prior to contract execution."
    )
    next_steps = (
        f"• Complete management approval and proceed to REC Sales Agreement execution with {alias}."
    )
    return current_status, next_steps

def _build_om_status_and_next_steps(data_dir: Path) -> tuple[str, str]:
    cfg = _ensure_main_agreements_json(data_dir)
    om_cfg = cfg.get("om") if isinstance(cfg, dict) else {}
    if not isinstance(om_cfg, dict):
        om_cfg = {}

    general = _load_json(data_dir / "general.json")
    opex = _load_json(data_dir / "opex_year1.json")

    capacities_mw: dict[str, float] = {}
    for project in general.get("projects") or []:
        if not isinstance(project, dict):
            continue
        name = str(project.get("name") or "").strip()
        if not name:
            continue
        inputs = project.get("general_inputs") or {}
        cap_mw = _to_optional_float(inputs.get("Total Capacity DC"))
        if cap_mw is None or cap_mw <= 0:
            continue
        capacities_mw[name] = cap_mw

    total_mw = 0.0
    total_om_krw = 0.0
    for project in opex.get("projects") or []:
        if not isinstance(project, dict):
            continue
        name = str(project.get("name") or "").strip()
        cap_mw = capacities_mw.get(name)
        if cap_mw is None or cap_mw <= 0:
            continue
        block = project.get("opex_year1") or {}
        om_cost = _to_float((block.get("O&M") or {}).get("Cost"))
        total_mw += cap_mw
        total_om_krw += om_cost

    om_krw_per_mw = (total_om_krw / total_mw) if total_mw > 0 else 0.0
    om_krw_per_mw_text = f"{om_krw_per_mw:,.0f}" if om_krw_per_mw > 0 else "TBD"

    rfp_status = str(om_cfg.get("rfp_status") or "RfP conducted for O&M providers").strip()
    modelling_alignment = str(om_cfg.get("modelling_alignment") or "in line with profitability modelling").strip()
    selection_basis = str(om_cfg.get("selection_basis") or "penalty for non-compliance track records").strip()
    leading_candidate = str(om_cfg.get("leading_candidate") or "to be confirmed").strip()

    gv_review_status = str(om_cfg.get("greenvolt_review_status") or "pending review and comments").strip()
    selected_vendor = str(om_cfg.get("selected_vendor") or "selected vendor").strip()
    proceed_terms = str(om_cfg.get("proceed_terms") or "the same terms and conditions").strip()

    current_status = (
        f"• OR has {rfp_status}.\n"
        f"• O&M commercial benchmark is KRW {om_krw_per_mw_text}/MW and remains {modelling_alignment}; "
        f"vendor selection will prioritize {selection_basis}.\n"
        f"• Current leading candidate: {leading_candidate}."
    )
    next_steps = (
        f"• Following Greenvolt's review/comments ({gv_review_status}), {selected_vendor} will proceed under {proceed_terms}."
    )
    return current_status, next_steps

def _build_gantt_table_from_saved_phase(data_dir: Path, phase: str) -> tuple[list[str], list[list[str]]]:
    """data_gantt/{phase}.json을 읽어 슬라이드용 간트 표를 생성한다."""
    phase_key = phase.strip().lower()
    if phase_key not in {"phase1", "phase2"}:
        phase_key = "phase1"

    phase_path = _resolve_gantt_phase_path(data_dir, phase_key)
    raw_tasks = _load_json(phase_path)
    tasks = raw_tasks if isinstance(raw_tasks, list) else []

    parsed: list[dict[str, Any]] = []
    for item in tasks:
        if not isinstance(item, dict):
            continue
        name = str(item.get("name") or "").strip()
        ttype = str(item.get("type") or "task").strip().lower()
        start = _to_optional_float(item.get("startWeekIndex"))
        end = _to_optional_float(item.get("endWeekIndex"))
        if not name or start is None or end is None:
            continue
        start_idx = int(start)
        end_idx = int(end)
        if end_idx < start_idx:
            start_idx, end_idx = end_idx, start_idx
        parsed.append({
            "name": name,
            "type": ttype,
            "start": start_idx,
            "end": end_idx,
        })

    if not parsed:
        return ["Task", "Start W", "End W", "Duration", "Timeline"], [["No saved tasks", "", "", "", ""]]

    max_week = max((t["end"] for t in parsed), default=0)
    max_week = max(0, min(max_week, 51))

    headers = ["Task", "Start W", "End W", "Duration", "Timeline"]
    rows: list[list[str]] = []
    timeline_slots = min(28, max(12, max_week + 1))

    for task in parsed:
        start_idx = max(0, min(task["start"], max_week))
        end_idx = max(0, min(task["end"], max_week))
        duration = (end_idx - start_idx) + 1
        timeline_chars: list[str] = []
        for slot_idx in range(timeline_slots):
            if timeline_slots <= 1 or max_week <= 0:
                week_idx = 0
            else:
                week_idx = round((slot_idx / (timeline_slots - 1)) * max_week)
            timeline_chars.append("=" if start_idx <= week_idx <= end_idx else " ")

        timeline = f"W{start_idx + 1:02d}-W{end_idx + 1:02d} |{''.join(timeline_chars)}|"
        task_name = task["name"] if task["type"] == "group" else f"  {task['name']}"
        rows.append([
            str(task_name),
            str(start_idx + 1),
            str(end_idx + 1),
            str(duration),
            timeline,
        ])

    return headers, rows

def _build_construction_planning_slide13(data_dir: Path) -> tuple[list[str], list[list[str]], str]:
    """슬라이드 13용 Construction Planning 요약 표/설명 생성."""
    general = _load_json(data_dir / "general.json")
    projects = general.get("projects") if isinstance(general, dict) else []
    if not isinstance(projects, list):
        projects = []

    def _fmt_dd_mmm_yy(value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, datetime):
            dt = value
        else:
            txt = str(value).strip()
            if not txt:
                return ""
            dt = None
            for parser in (lambda s: datetime.fromisoformat(s.replace(" ", "T")), lambda s: datetime.strptime(s.split()[0], "%Y-%m-%d")):
                try:
                    dt = parser(txt)
                    break
                except Exception:
                    continue
            if dt is None:
                return txt
        return dt.strftime("%b %d, '%y")

    headers = [
        "SPV",
        "Phase\n#",
        "No.",
        "PJT Name",
        "Capacity\n(kWp)",
        "RtB Date",
        "Expected SOC",
        "Expected COD",
    ]

    rows: list[list[str]] = []
    phase1_capacity = 0.0
    phase2_capacity = 0.0
    total_capacity = 0.0
    phase_split_idx = 7
    phase1_soc: list[datetime] = []
    phase2_soc: list[datetime] = []
    phase1_rtb: list[datetime] = []
    phase2_rtb: list[datetime] = []
    phase1_cod: list[datetime] = []
    phase2_cod: list[datetime] = []

    def _as_datetime(value: Any) -> Optional[datetime]:
        if value is None:
            return None
        if isinstance(value, datetime):
            return value
        txt = str(value).strip()
        if not txt:
            return None
        for parser in (lambda s: datetime.fromisoformat(s.replace(" ", "T")), lambda s: datetime.strptime(s.split()[0], "%Y-%m-%d")):
            try:
                return parser(txt)
            except Exception:
                continue
        return None

    def _month_window(dates: list[datetime]) -> str:
        if not dates:
            return "N/A"
        dmin = min(dates)
        dmax = max(dates)
        start = dmin.strftime("%b %Y")
        end = dmax.strftime("%b %Y")
        return start if start == end else f"{start} - {end}"

    for idx, project in enumerate(projects, start=1):
        inputs = project.get("general_inputs") if isinstance(project, dict) else {}
        if not isinstance(inputs, dict):
            inputs = {}
        project_name = str(inputs.get("Project name") or "")
        capacity_mw = _to_optional_float(inputs.get("Total Capacity DC")) or 0.0
        capacity_kwp = capacity_mw * 1000
        total_capacity += capacity_kwp
        if idx <= phase_split_idx:
            phase1_capacity += capacity_kwp
            rtb_dt = _as_datetime(inputs.get("RTB date"))
            soc_dt = _as_datetime(inputs.get("SOC date"))
            cod_dt = _as_datetime(inputs.get("COD date"))
            if rtb_dt is not None:
                phase1_rtb.append(rtb_dt)
            if soc_dt is not None:
                phase1_soc.append(soc_dt)
            if cod_dt is not None:
                phase1_cod.append(cod_dt)
        else:
            phase2_capacity += capacity_kwp
            rtb_dt = _as_datetime(inputs.get("RTB date"))
            soc_dt = _as_datetime(inputs.get("SOC date"))
            cod_dt = _as_datetime(inputs.get("COD date"))
            if rtb_dt is not None:
                phase2_rtb.append(rtb_dt)
            if soc_dt is not None:
                phase2_soc.append(soc_dt)
            if cod_dt is not None:
                phase2_cod.append(cod_dt)

        rows.append([
            "BOLT#2" if idx == 1 else "",
            "1" if idx == 1 else ("2" if idx == phase_split_idx + 1 else ""),
            str(idx),
            project_name,
            _format_optional_number(capacity_kwp, 2),
            _fmt_dd_mmm_yy(inputs.get("RTB date")),
            _fmt_dd_mmm_yy(inputs.get("SOC date")),
            _fmt_dd_mmm_yy(inputs.get("COD date")),
        ])

    if rows:
        rows.insert(min(phase_split_idx, len(rows)), [
            "", "", "", "Sub Total", _format_optional_number(phase1_capacity, 2), "", "", ""
        ])
        if len(projects) > phase_split_idx:
            rows.append(["", "", "", "Sub Total", _format_optional_number(phase2_capacity, 2), "", "", ""])
        rows.append(["", "", "", "TOTAL", _format_optional_number(total_capacity, 2), "", "", ""])

    phase1_count = min(phase_split_idx, len(projects))
    phase2_count = max(0, len(projects) - phase_split_idx)
    right_text = (
        "Engineering (Phase 1,2)\n"
        f"- Phase 1: {phase1_count} projects, {_format_optional_number(phase1_capacity, 2)} kWp, RTB window {_month_window(phase1_rtb)}.\n"
        f"- Phase 2: {phase2_count} projects, {_format_optional_number(phase2_capacity, 2)} kWp, RTB window {_month_window(phase2_rtb)}.\n"
        "Procurement (Phase 1,2)\n"
        f"- Phase 1: SOC window {_month_window(phase1_soc)} for procurement readiness.\n"
        f"- Phase 2: SOC window {_month_window(phase2_soc)} for procurement readiness.\n"
        "Construction (Phase 1,2)\n"
        f"- Phase 1: COD window {_month_window(phase1_cod)}.\n"
        f"- Phase 2: COD window {_month_window(phase2_cod)}.\n"
        f"- Portfolio baseline: {_format_optional_number(total_capacity, 2)} kWp across {len(projects)} projects."
    )

    return headers, rows, right_text


def _build_equipment_procurement_case_text(data_dir: Path) -> str:
    """data/*.json 수치를 기반으로 procurement narrative 불릿을 구성한다."""
    general = _load_json(data_dir / "general.json")
    opex = _load_json(data_dir / "opex_year1.json")

    def _as_dict(value: Any) -> dict[str, Any]:
        return value if isinstance(value, dict) else {}

    projects = general.get("projects") if isinstance(general, dict) else []
    if not isinstance(projects, list):
        projects = []

    capacities_kwp: list[float] = []
    cod_years: list[int] = []
    panel_models: set[str] = set()
    panel_count_total = 0.0

    for p in projects:
        if not isinstance(p, dict):
            continue
        gi = _as_dict(p.get("general_inputs"))

        cap_mw = _to_optional_float(gi.get("Total Capacity DC"))
        if cap_mw is not None and cap_mw > 0:
            capacities_kwp.append(cap_mw * 1000)

        cod = _to_date_str(gi.get("COD date"))
        if cod and re.match(r"^\d{4}-\d{2}-\d{2}$", cod):
            cod_years.append(int(cod[:4]))

        panel_name = str(gi.get("WTG / Panels name") or "").strip()
        if panel_name:
            panel_models.add(panel_name)
        panel_num = _to_optional_float(gi.get("WTG / Panels number"))
        if panel_num is not None and panel_num > 0:
            panel_count_total += panel_num

    avg_kwp = (sum(capacities_kwp) / len(capacities_kwp)) if capacities_kwp else 0.0
    min_kwp = min(capacities_kwp) if capacities_kwp else 0.0
    max_kwp = max(capacities_kwp) if capacities_kwp else 0.0
    target_year = min(cod_years) if cod_years else datetime.now().year

    # Cost metrics from opex_year1 (portfolio-weighted)
    cost_projects = opex.get("projects") if isinstance(opex, dict) else []
    if not isinstance(cost_projects, list):
        cost_projects = []

    capacity_lookup: dict[str, float] = {}
    for p in projects:
        if not isinstance(p, dict):
            continue
        name = str(p.get("name") or "").strip()
        gi = _as_dict(p.get("general_inputs"))
        cap_mw = _to_optional_float(gi.get("Total Capacity DC"))
        if name and cap_mw is not None and cap_mw > 0:
            capacity_lookup[name] = cap_mw

    total_mw = 0.0
    modules_total = 0.0
    bos_total = 0.0
    grid_total = 0.0
    contingency_total = 0.0
    for cp in cost_projects:
        if not isinstance(cp, dict):
            continue
        name = str(cp.get("name") or "").strip()
        cap_mw = capacity_lookup.get(name)
        if cap_mw is None:
            continue
        oy = _as_dict(cp.get("opex_year1"))
        modules_total += _to_float(_as_dict(oy.get("Modules")).get("Cost"))
        bos_total += _to_float(_as_dict(oy.get("BOS")).get("Cost"))
        grid_total += _to_float(_as_dict(oy.get("Grid Connection Cost")).get("Cost"))
        contingency_total += _to_float(_as_dict(oy.get("Contingency")).get("Cost"))
        total_mw += cap_mw

    def _per_mw(total_cost: float) -> float:
        return (total_cost / total_mw) if total_mw > 0 else 0.0

    modules_per_mw_m = _per_mw(modules_total) / 1_000_000
    bos_per_mw_m = _per_mw(bos_total) / 1_000_000
    grid_per_mw_m = _per_mw(grid_total) / 1_000_000
    contingency_per_mw_m = _per_mw(contingency_total) / 1_000_000
    procurement_per_mw_m = modules_per_mw_m + bos_per_mw_m + grid_per_mw_m

    bullets = [
        f"Each cluster set is assessed with representative capacity around {_format_optional_number(avg_kwp, 0)}kW (range {_format_optional_number(min_kwp, 0)}-{_format_optional_number(max_kwp, 0)}kW).",
        f"Target procurement pricing baseline is set for the {target_year} COD pipeline covering {len(projects)} projects.",
        f"Equipment standardization baseline includes {len(panel_models)} panel model(s) and approximately {_format_optional_number(panel_count_total, 0)} modules across the portfolio.",
        f"Current procurement cost benchmark is about KRW {_format_optional_number(procurement_per_mw_m, 1)}m/MW (Modules {_format_optional_number(modules_per_mw_m, 1)}m, BOS {_format_optional_number(bos_per_mw_m, 1)}m, Grid {_format_optional_number(grid_per_mw_m, 1)}m).",
        f"Contingency allowance is tracked at KRW {_format_optional_number(contingency_per_mw_m, 1)}m/MW and remains a key buffer against supplier-side price movement.",
        "All values remain subject to ongoing supplier negotiations and periodic cost refresh cycles.",
    ]

    return "\n".join(f"- {b}" for b in bullets)

def _ensure_main_agreements_json(data_dir: Path) -> dict[str, Any]:
    """슬라이드 10 생성 시 main_agreements.json을 자동 생성/갱신한다."""
    path = data_dir / "main_agreements.json"
    general = _load_json(data_dir / "general.json")
    existing = _load_json(path) if path.exists() else {}

    project_count = len(general.get("projects") or []) if isinstance(general, dict) else 0
    rec_tariff = _extract_rec_tariff_krw_per_kwh(data_dir)

    def _merge_dict(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
        merged: dict[str, Any] = dict(base)
        for key, value in override.items():
            if isinstance(value, dict) and isinstance(merged.get(key), dict):
                merged[key] = _merge_dict(merged[key], value)
            else:
                merged[key] = value
        return merged

    generated = {
        "source": {
            "generated_by": "capex_content_builder._ensure_main_agreements_json",
            "generated_at": datetime.now().strftime("%Y-%m-%d"),
            "project_count": project_count,
        },
        "epc": {
            "review_status": "Contractual conditions review status to be confirmed",
            "signature_target": "TBD",
            "contractors": [
                {
                    "name": "EPC contractor allocation pending",
                    "project_count": project_count,
                }
            ],
        },
        "rec_sales": {
            "negotiation_status": "Under discussion",
            "offtaker_name": "Offtaker to be confirmed",
            "alias": "TBD",
            "offtaker_type": "RPS obligor",
            "term_years": None,
            "price_krw_per_kwh": rec_tariff,
            "offer_valid_until": f"end of {datetime.now().year}",
            "management_approval": "pending",
        },
        "om": {
            "rfp_status": "conducted an RfP for O&M providers",
            "modelling_alignment": "in line with profitability modelling",
            "selection_basis": "penalties for non-compliance track records",
            "leading_candidate": "to be confirmed",
            "greenvolt_review_status": "pending review and comments",
            "selected_vendor": "the selected vendor",
            "proceed_terms": "the same terms and conditions"
        }
    }

    payload = generated
    if isinstance(existing, dict) and existing:
        payload = _merge_dict(generated, existing)

    try:
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:  # pylint: disable=broad-except
        print(f"⚠️ main_agreements.json 생성 실패: {exc}")
        return payload

    return payload

def _build_lease_agreement_table(data_dir: Path, split_index: int = 7) -> tuple[list[str], list[list[str]]]:
    general_path = data_dir / "general.json"
    general = _load_json(general_path)

    headers = [
        "SPV",
        "Phase #",
        "No.",
        "PJT Name",
        "Capacity\n(kWp)",
        "Basis",
        "Lease Fee\n(KRW/kW)",
        "1-5 yrs\nKRW",
        "1-5 yrs\nc.EUR",
        "6-20 yrs\nKRW",
        "6-20 yrs\nc.EUR",
    ]

    rows: list[list[str]] = []
    projects = general.get("projects") or []

    prepay_rate = LEASE_PREPAY_RATE
    annual_rate = LEASE_ANNUAL_RATE
    fx_rate = LEASE_FX_RATE

    total_capacity = 0.0
    total_prepay = 0.0
    total_annual = 0.0
    phase_capacity = [0.0, 0.0]
    phase_prepay = [0.0, 0.0]
    phase_annual = [0.0, 0.0]
    has_phase2 = False

    for idx, project in enumerate(projects, start=1):
        inputs = project.get("general_inputs") or {}
        project_name = inputs.get("Project name") or ""
        capacity_mwp = inputs.get("Total Capacity DC")
        capacity_kwp = _to_optional_float(capacity_mwp)
        if capacity_kwp is not None:
            capacity_kwp *= 1000
            total_capacity += capacity_kwp

        prepay_krw = capacity_kwp * prepay_rate if capacity_kwp is not None else None
        annual_krw = capacity_kwp * annual_rate if capacity_kwp is not None else None
        prepay_eur = prepay_krw / fx_rate if prepay_krw is not None else None
        annual_eur = annual_krw / fx_rate if annual_krw is not None else None

        phase = 1 if idx <= split_index else 2
        if phase == 2:
            has_phase2 = True
        if capacity_kwp is not None:
            phase_capacity[phase - 1] += capacity_kwp
        if prepay_krw is not None:
            phase_prepay[phase - 1] += prepay_krw
        if annual_krw is not None:
            phase_annual[phase - 1] += annual_krw
        if prepay_krw is not None:
            total_prepay += prepay_krw
        if annual_krw is not None:
            total_annual += annual_krw

        rows.append([
            "BOLT#2" if idx == 1 else "",
            str(phase),
            str(idx),
            str(project_name),
            _format_optional_number(capacity_kwp, 2),
            "Prepayment for 5 years",
            _format_optional_number(prepay_rate, 0),
            _format_optional_number(prepay_krw, 0),
            _format_optional_number(prepay_eur, 0),
            "",
            "",
        ])

        rows.append([
            "",
            "",
            "",
            "",
            "",
            "Annual lease fee",
            _format_optional_number(annual_rate, 0),
            "",
            "",
            _format_optional_number(annual_krw, 0),
            _format_optional_number(annual_eur, 0),
        ])

        if idx == split_index:
            rows.append([
                "",
                "",
                "",
                "Sub Total",
                _format_optional_number(phase_capacity[0], 2),
                "",
                "",
                _format_optional_number(phase_prepay[0], 0),
                _format_optional_number(phase_prepay[0] / fx_rate, 0) if phase_prepay[0] else "",
                _format_optional_number(phase_annual[0], 0),
                _format_optional_number(phase_annual[0] / fx_rate, 0) if phase_annual[0] else "",
            ])

    if rows and has_phase2:
        rows.append([
            "",
            "",
            "",
            "Sub Total",
            _format_optional_number(phase_capacity[1], 2),
            "",
            "",
            _format_optional_number(phase_prepay[1], 0),
            _format_optional_number(phase_prepay[1] / fx_rate, 0) if phase_prepay[1] else "",
            _format_optional_number(phase_annual[1], 0),
            _format_optional_number(phase_annual[1] / fx_rate, 0) if phase_annual[1] else "",
        ])

    if rows:
        rows.append([
            "",
            "",
            "",
            "TOTAL",
            _format_optional_number(total_capacity, 2),
            "",
            "",
            _format_optional_number(total_prepay, 0),
            _format_optional_number(total_prepay / fx_rate, 0) if total_prepay else "",
            _format_optional_number(total_annual, 0),
            _format_optional_number(total_annual / fx_rate, 0) if total_annual else "",
        ])

    return headers, rows

def _build_capex_approval_values(data_dir: Path) -> dict[str, str]:
    general_path = data_dir / "general.json"
    opex_path = data_dir / "opex_year1.json"

    general = _load_json(general_path)
    opex = _load_json(opex_path)

    projects = general.get("projects") or []
    total_capacity = 0.0
    for proj in projects:
        inputs = proj.get("general_inputs") or {}
        total_capacity += _to_float(inputs.get("Total Capacity DC"))

    capex_projects = opex.get("projects") or []
    capex_keys = [
        "BOS",
        "Modules",
        "Grid Connection Cost",
        "Development cost",
        "Land Purchase",
        "Other DEVEX(Broker Fee)",
        "Other Capex",
    ]
    total_capex_krw = 0.0
    for proj in capex_projects:
        blocks = (proj.get("opex_year1") or {})
        for key in capex_keys:
            block = blocks.get(key) or {}
            total_capex_krw += _to_float(block.get("Cost"))

    capex_m_krw = round(total_capex_krw / 1_000_000)
    fx_krw_per_eur = 1470.0
    capex_eur = round(total_capex_krw / fx_krw_per_eur)

    capex_krw_text = f"KRW {capex_m_krw:,.0f}M"
    capex_eur_text = f"EUR {capex_eur:,.0f}"
    capacity_text = f"{total_capacity:.3f}MWp"

    approval_text = (
        "Expected return: Approve the expected return based on the latest revenues and capex:\n"
        "- Approval criterion: Portfolio blended Equity IRR of 9.64% (≈9.65%)\n"
        "- Average IRR (auxiliary, unweighted average of project IRRs): 8.95%\n"
        f"capex :  Approval of total capex of {capex_krw_text} approximately {capex_eur_text} "
        f"for the {capacity_text} portfolio"
    )

    return {
        "approval_text": approval_text,
        "capex_krw_text": capex_krw_text,
        "capex_eur_text": capex_eur_text,
        "capacity_text": capacity_text,
    }

def _build_capex_approval_text(data_dir: Path) -> str:
    return _build_capex_approval_values(data_dir)["approval_text"]

def _build_phase_data_from_inputsheet() -> tuple[list[dict], list[dict]]:
    """레그 입력시트(Excel)를 직접 파싱해 Project 메타데이터를 추출."""

    base_dir = Path(__file__).resolve().parent
    candidates = [
        base_dir / "pv_solar_sample_docs" / "test2.xlsx",
        base_dir.parent / "pv_solar_sample_docs" / "test2.xlsx",
    ]

    xlsx_path = next((p for p in candidates if p.exists()), None)
    if xlsx_path is None:
        print("⚠️ test2.xlsx 파일을 찾을 수 없습니다.")
        return [], []

    try:
        import pandas as pd  # type: ignore
    except ImportError:
        print("⚠️ pandas 미설치: 입력시트를 파싱하지 못합니다.")
        return [], []

    excel = pd.ExcelFile(xlsx_path)
    sheet_name = next(
        (name for name in excel.sheet_names if str(name).upper().startswith("INPUTSHEET")),
        None,
    )
    if sheet_name is None:
        print("⚠️ INPUTSHEET로 시작하는 시트를 찾을 수 없습니다.")
        return [], []

    df = pd.read_excel(excel, sheet_name=sheet_name, header=None)

    header_row = df.iloc[1]
    project_cols: list[tuple[int, str, int]] = []
    for col_idx, val in header_row.items():
        if isinstance(val, str) and val.strip().lower().startswith("project"):
            phase = 1 if col_idx <= 14 else 2
            project_cols.append((col_idx, val.strip(), phase))

    def get(row_idx: int, col_idx: int, *, is_date: bool = False):
        raw = df.iat[row_idx, col_idx] if row_idx < len(df.index) and col_idx < len(df.columns) else None
        raw = _clean_val(raw)
        if is_date:
            return _to_date_str(raw)
        return raw

    projects: list[dict] = []
    for col_idx, proj_id, phase in project_cols:
        entry = {
            "project_id": proj_id,
            "phase": phase,
            "seller": get(4, col_idx),
            "project_name": get(5, col_idx),
            "country": get(6, col_idx),
            "region": get(7, col_idx),
            "technology": get(8, col_idx),
            "status": get(9, col_idx),
            "transaction_structure": get(10, col_idx),
            "input_source": get(11, col_idx),
            "currency": get(13, col_idx),
            "rtb_date": get(15, col_idx, is_date=True),
            "cod_date": get(18, col_idx, is_date=True),
            "capacity_dc_mwp": get(23, col_idx),
        }
        projects.append(entry)

    # JSON 파일로도 보관 (LLM 컨텍스트용)
    try:
        out_path = xlsx_path.parent / "projects.json"
        out_path.write_text(json.dumps(projects, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        print(f"✅ 프로젝트 JSON 추출 완료: {out_path}")
    except Exception as exc:  # pylint: disable=broad-except
        print(f"⚠️ 프로젝트 JSON 저장 실패: {exc}")

    phase1_projects = [p for p in projects if p.get("phase") == 1]
    phase2_projects = [p for p in projects if p.get("phase") == 2]
    return phase1_projects, phase2_projects

def build_content_list_from_api(
    documents: Optional[List[str]] = None,
    data_dir: Optional[Path | str] = None,
) -> List[SlideContent]:
    """슬라이드마다 API를 호출해 내용을 채운 뒤 SlideContent 리스트를 생성."""

    docs = documents if documents else DEFAULT_API_DOCS

    # cover_subtitle = get_api_text_or_default(
    #     "레그에 등록된 test2.xlsx 데이터를 기반으로 이 문서를 한 줄로 요약해 표지 부제로만 반환하세요.",
    #     "Project overview subtitle",
    #     documents=docs,
    #     rag=True,
    #     username=DEFAULT_API_USERNAME,
    # )
    cover_subtitle = "대한민국 분산형 태양광 발전 프로젝트 BOLT #2에 대한 CAPEX 승인 요청---ㅅㄷㅌㅅ"

    # 입력시트 데이터를 코드 상 상수에서 읽어 Phase별 요약을 동적으로 생성
    # phase1_raw, phase2_raw = _build_phase_data_from_inputsheet()
    phase1_raw, phase2_raw = [], []

    phase1_regions = _summarize_regions(phase1_raw)
    phase1_cod_window = _cod_window(phase1_raw)
    phase1_count = len(phase1_raw)

    phase1_total_dc = 0.0
    phase1_provinces: list[str] = []
    phase1_statuses: list[str] = []
    for p in phase1_raw:
        val = p.get("capacity_dc_mwp")
        if isinstance(val, (int, float)):
            phase1_total_dc += float(val)
        prov = _province_from_region(p.get("region"))
        if prov:
            phase1_provinces.append(prov)
        st = p.get("status")
        if st:
            phase1_statuses.append(str(st))

    phase1_provinces_txt = ", ".join(sorted(set(phase1_provinces))) if phase1_provinces else "지역 정보 없음"
    if phase1_statuses:
        from collections import Counter
        phase1_top_status, _ = Counter(phase1_statuses).most_common(1)[0]
    else:
        phase1_top_status = "데이터 확인 필요"

    total_dc_phase2 = 0.0
    for p in phase2_raw:
        val = p.get("capacity_dc_mwp")
        if isinstance(val, (int, float)):
            total_dc_phase2 += float(val)
    phase2_regions = _summarize_regions(phase2_raw)
    phase2_cod_window = _cod_window(phase2_raw)
    phase2_count = len(phase2_raw)

    phase1_table = [
        [
            "1",
            (
                f"RTB 기준 {phase1_count}개 프로젝트, 총 {phase1_total_dc:.2f} MWp. "
                f"주요 권역: {phase1_provinces_txt}. COD 목표: {phase1_cod_window}. "
                f"상태: {phase1_top_status}. 계통 연계비는 미납 상태이며 COD 이전 CAPEX에서 충당 예정"
            ),
        ],
        [
            "2",
            (
                "상태: RTB 유지, 거래 구조 Ground, 통화 KRW. 계통연계비 CAPEX 지급 가정. "
                "REC 단가/수익성 지표는 입력 데이터에 없어 추가 확인 필요"
            ),
        ],
        [
            "3",
            (
                f"RTB 기준 {phase1_count}개 프로젝트, 총 {phase1_total_dc:.2f} MWp. "
                f"주요 권역: {phase1_provinces_txt}. COD 목표: {phase1_cod_window}. "
                f"상태: {phase1_top_status}. 계통 연계비는 미납 상태이며 COD 이전 CAPEX에서 충당 예정"
            ),
        ],
        [
            "4",
            (
                "상태: RTB 유지, 거래 구조 Ground, 통화 KRW. 계통연계비 CAPEX 지급 가정. "
                "REC 단가/수익성 지표는 입력 데이터에 없어 추가 확인 필요"
            ),
        ],
        [
            "5",
            (
                f"RTB 기준 {phase1_count}개 프로젝트, 총 {phase1_total_dc:.2f} MWp. "
                f"주요 권역: {phase1_provinces_txt}. COD 목표: {phase1_cod_window}. "
                f"상태: {phase1_top_status}. 계통 연계비는 미납 상태이며 COD 이전 CAPEX에서 충당 예정"
            ),
        ],
        [
            "6",
            (
                "상태: RTB 유지, 거래 구조 Ground, 통화 KRW. 계통연계비 CAPEX 지급 가정. "
                "REC 단가/수익성 지표는 입력 데이터에 없어 추가 확인 필요"
            ),
        ],
        [
            "7",
            (
                "상태: RTB 유지, 거래 구조 Ground, 통화 KRW. 계통연계비 CAPEX 지급 가정. "
                "REC 단가/수익성 지표는 입력 데이터에 없어 추가 확인 필요"
            ),
        ],
    ]

    phase2_table = [
        [
            "1",
            (
                f"{phase2_count}개 루프탑 PV, 총 {total_dc_phase2:.3f} MWp DC, 지역: {phase2_regions}; "
                f"COD 목표: {phase2_cod_window}"
            ),
        ],
        [
            "2",
            (
                "상태: EBL 2건, RtB 4건 (입력시트 기준). 계통연계비 CAPEX 포함 가정, "
                "REC/수익성 지표는 추가 데이터 확인 필요"
            ),
        ],
        [
            "3",
            (
                f"{phase2_count}개 루프탑 PV, 총 {total_dc_phase2:.3f} MWp DC, 지역: {phase2_regions}; "
                f"COD 목표: {phase2_cod_window}"
            ),
        ],
        [
            "4",
            (
                "상태: EBL 2건, RtB 4건 (입력시트 기준). 계통연계비 CAPEX 포함 가정, "
                "REC/수익성 지표는 추가 데이터 확인 필요"
            ),
        ],
        [
            "5",
            (
                f"{phase2_count}개 루프탑 PV, 총 {total_dc_phase2:.3f} MWp DC, 지역: {phase2_regions}; "
                f"COD 목표: {phase2_cod_window}"
            ),
        ],
        [
            "6",
            (
                "상태: EBL 2건, RtB 4건 (입력시트 기준). 계통연계비 CAPEX 포함 가정, "
                "REC/수익성 지표는 추가 데이터 확인 필요"
            ),
        ],
        [
            "7",
            (
                f"{phase2_count}개 루프탑 PV, 총 {total_dc_phase2:.3f} MWp DC, 지역: {phase2_regions}; "
                f"COD 목표: {phase2_cod_window}"
            ),
        ],
        
    ]

    exec_summary_text = (
        "BOLT #2 - Phase 1\n" + "\n".join(f"• {row[1]}" for row in phase1_table) +
        "\nBOLT #2 - Phase 2\n" + "\n".join(f"• {row[1]}" for row in phase2_table)
    )

    # session_text = get_api_text_or_default(
    #     "이 문서의 섹션 표지에 들어갈 한 문장 설명을 작성하세요. 레그 test2.xlsx 내용에서 전체 프로젝트의 핵심 목적을 1문장으로 정리해 주세요.",
    #     "프로젝트 핵심 목적 요약",
    #     documents=docs,
    #     rag=True,
    #     username=DEFAULT_API_USERNAME,
    # )
    session_text = "대한민국 분산형 태양광 발전 프로젝트 BOLT #2의 시장 진출 전략"

    content_list: List[SlideContent] = []

    # 슬라이드 1: 표지 (템플릿 1페이지, layout='title')
    content_list.append(SlideContent(
        layout="title",
        title="Capex Approval – S. Korea DG BOLT#2",
        subtitle=cover_subtitle,
        date="November 15, 2023",
        order=1,
    ))

    # # 슬라이드 2: Executive Summary (템플릿 18페이지, layout='')
    content_list.append(SlideContent(
        layout="two_tables",
        title="Executive Summary",
        content=exec_summary_text,
        table_headers=["#", "BOLT #2 - Phase 1"],
        table_headers_2=["#", "BOLT #2 - Phase 2"],
        table_data=phase1_table,
        table_data_2=phase2_table,
        order=2,
    ))



    # 슬라이드 3: Approval request (템플릿 17페이지, layout='approval_request')
    data_dir = _resolve_data_dir(data_dir)
    approval_text = _build_capex_approval_text(data_dir)
    content_list.append(SlideContent(
        layout="approval_request",
        title="Approval request",
        content=approval_text,
        order=3,
    ))

    # 슬라이드 4: 섹션 표지 (템플릿 16페이지, layout='session_title')
    content_list.append(SlideContent(
        layout="session_title",
        title="01",
        content='Route to Market',
        # content=session_text,
        order=4,
    ))

    # 슬라이드 5: Route to Market (템플릿 10페이지, layout='text_large_table')
    route_headers, route_rows = _build_route_to_market_table(data_dir)
    content_list.append(SlideContent(
        layout="title_subtitle_table",
        title="Renewable Energy Certificate(REC) Sales Agreement Negotiation",
        content=(
            "REC sales under discussion with Samcheok Bluepower(SBP) at KRW 189/kWh "
            "(20-year fixed term)."
        ),
        table_headers=route_headers,
        table_data=route_rows,
        order=5,
    ))
    
    
    # 슬라이드 6: 섹션 표지 (템플릿 16페이지, layout='session_title')
    content_list.append(SlideContent(
        layout="session_title",
        title="02",
        content='Project Description',
        # content=session_text,
        order=6,
    ))

    # 슬라이드 7: COD 2026 Pipeline + Technical Solution 공간
    cod_headers, cod_rows = _build_cod_pipeline_table(data_dir)
    technical_solution_text = _build_technical_solution_text(data_dir)
    content_list.append(SlideContent(
        layout="cod_pipeline",
        title=(
            "BOLT#2 is Korean PV Portfolio with COD in May-Jun. 2026, "
            "Amounting to 4.2MWp"
        ),
        content=technical_solution_text,
        subtitle=f"COD {datetime.now().year} Pipeline",
        table_headers=cod_headers,
        table_data=cod_rows,
        order=7,
    ))

    # 슬라이드 8: Permits (템플릿 10페이지, layout='permits')
    permits_headers, permits_rows = _build_permits_table(data_dir)
    content_list.append(SlideContent(
        layout="permits",
        title="Permits",
        table_headers=permits_headers,
        table_data=permits_rows,
        order=8,
    ))
    
    # 슬라이드 9: Main Agreements : Lease Agreement
    # - 생성 시 layout='lease_agreement'은 전용 매핑이 없어 기본 레이아웃을 사용
    lease_headers, lease_rows = _build_lease_agreement_table(data_dir, split_index=7)
    content_list.append(SlideContent(
        layout="main_agreements_lease",
        title="Main Agreements : Lease Agreement",
        content=(
            "The following projects have executed lease agreements and are required to pay "
            "a five-year lease prepayment upon COD. [1EUR=1,470 KRW / excl. VAT]"
        ),
        table_headers=lease_headers,
        table_data=lease_rows,
        order=9,
    ))

    # 슬라이드 10: Main Agreements : EPC, REC Sales Agreement, O&M
    # - 템플릿 19페이지(layout='main_agreements') 사용
    main_headers = ["Agreement", "Completion", "Current Status", "Next Steps"]
    epc_current_status, epc_next_steps = _build_epc_status_and_next_steps(data_dir)
    rec_current_status, rec_next_steps = _build_rec_sales_status_and_next_steps(data_dir)
    om_current_status, om_next_steps = _build_om_status_and_next_steps(data_dir)
    main_rows = [
        ["EPC Contract", "Contractor Signature\nPending", epc_current_status, epc_next_steps],
        ["REC Sales\nAgreement", "On-going\nReview", rec_current_status, rec_next_steps],
        ["O&M", "On-going\nReview", om_current_status, om_next_steps],
    ]
    content_list.append(SlideContent(
        layout="main_agreements",
        title="Main Agreements : EPC, REC Sales Agreement, O&M",
        table_headers=main_headers,
        table_data=main_rows,
        order=10,
    ))

    # 슬라이드 11: Project Construction Plan - Phase 1 (템플릿 10페이지)
    gantt_headers_p1, gantt_rows_p1 = _build_gantt_table_from_saved_phase(data_dir, "phase1")
    content_list.append(SlideContent(
        layout="gantt_chart_template10",
        title="Project Construction Plan_Phase 1",
        subtitle="Data source: data_gantt/phase1.json",
        table_headers=gantt_headers_p1,
        table_data=gantt_rows_p1,
        order=11,
    ))

    # 슬라이드 12: Project Construction Plan - Phase 2 (템플릿 10페이지)
    gantt_headers_p2, gantt_rows_p2 = _build_gantt_table_from_saved_phase(data_dir, "phase2")
    content_list.append(SlideContent(
        layout="gantt_chart_template10",
        title="Project Construction Plan_Phase 2",
        subtitle="Data source: data_gantt/phase2.json",
        table_headers=gantt_headers_p2,
        table_data=gantt_rows_p2,
        order=12,
    ))

    # 슬라이드 13: Construction Planning - BOLT#2 Phase1,2 (템플릿 cod_pipeline)
    cp_headers, cp_rows, cp_text = _build_construction_planning_slide13(data_dir)
    content_list.append(SlideContent(
        layout="cod_pipeline_phase",
        title="Construction Planning - BOLT#2 Phase1,2",
        content=cp_text,
        subtitle="Construction baseline summary",
        table_headers=cp_headers,
        table_data=cp_rows,
        order=13,
    ))

    # 슬라이드 14: Equipment procurement process – Illustrative case (템플릿 23)
    content_list.append(SlideContent(
        layout="equipment_procurement_case",
        title="Equipment procurement process – Illustrative case",
        content=_build_equipment_procurement_case_text(data_dir),
        order=14,
    ))

    # 슬라이드 15: BOS negotiation (템플릿 24페이지, 표 데이터는 data 기반 입력 전까지 공란 유지)
    content_list.append(SlideContent(
        layout="bos_negotiation",
        title="BOS negotiation",
        subtitle="Data-driven table placeholder (values intentionally blank)",
        content="",
        order=15,
    ))
    


    # # 슬라이드 2: Executive Summary (표 2개 + 요약 텍스트)
    # content_list.append(SlideContent(
    #     title="Executive Summary",
    #     subtitle=table_caption,
    #     content=table_caption,
    #     table_headers=["#", "BOLT #2 - Phase 1"],
    #     table_data=[
    #         ["1", "240.69"],
    #         ["2", "386.24"],
    #         ["3", "210.00"],
    #         ["4", "171.00"],
    #     ],
    #     table_headers_2=["#", "BOLT #2 - Phase 2"],
    #     table_data_2=[
    #         ["1", "100"],
    #         ["2", "80"],
    #         ["3", "20"],
    #     ],
    #     layout='two_tables',
    #     order=2,
    # ))

    # # 슬라이드 3: Approval request
    # content_list.append(SlideContent(
    #     layout="approval_request",
    #     title="Approval request",
    #     content=approval_text,
    #     order=3,
    # ))

    
    return content_list
