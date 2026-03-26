import openpyxl
from datetime import datetime
from pathlib import Path

from pptx import Presentation
from pptx.util import Mm


def read_excel(filepath: str) -> list[dict]:
    """
    엑셀 파일을 읽어 인원 정보 리스트 반환.
    헤더: 이름, 영문이름, 부서, 직급, 회사명
    오류 시: [{"error": "메시지"}]
    """
    try:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if not rows:
            return [{"error": "엑셀 파일이 비어 있습니다."}]

        HEADER_MAP = {
            "이름":    "name_kr",
            "영문이름": "name_en",
            "부서":    "dept",
            "직급":    "rank",
            "회사명":  "company",
        }

        col_idx: dict[str, int] = {}
        for i, cell in enumerate(rows[0]):
            if cell is None:
                continue
            key = str(cell).strip()
            if key in HEADER_MAP:
                col_idx[HEADER_MAP[key]] = i

        if "name_kr" not in col_idx:
            return [{"error": "'이름' 헤더를 찾을 수 없습니다."}]

        people: list[dict] = []
        for row in rows[1:]:
            def _get(field: str) -> str | None:
                i = col_idx.get(field)
                if i is None or i >= len(row):
                    return None
                val = row[i]
                s = str(val).strip() if val is not None else ""
                return s if s else None

            name_kr = _get("name_kr")
            if not name_kr:
                continue

            people.append({
                "name_kr": name_kr,
                "name_en": _get("name_en"),
                "dept":    _get("dept"),
                "rank":    _get("rank"),
                "company": _get("company"),
            })

        return people

    except Exception as exc:
        return [{"error": f"엑셀 읽기 오류: {exc}"}]


def generate_badges(
    people,
    logo_path,
    company_name,
    badge_type,
    design,
    color,
    size,
    layout,
    event_mode,
    event_info,
    output_dir,
) -> str | None:
    """
    명찰 PPT 생성 후 저장 경로 반환.
    badge_type : "badge" | "lanyard"
    """
    BASE_DIR = Path(__file__).parent

    if badge_type == "badge":
        from designs.badge import place_on_slide, BADGE_SIZES, BADGE_LAYOUTS
        size_mm    = BADGE_SIZES.get(size, (70.0, 25.0))
        cols, rows = BADGE_LAYOUTS.get(layout, (2, 2))
        _color     = color or "Gold"
        slide_w, slide_h = 210.0, 297.0
        layout_key = layout

    elif badge_type == "lanyard":
        from designs.lanyard import place_on_slide, LANYARD_SIZES, LANYARD_LAYOUTS
        size_mm = LANYARD_SIZES.get(size, (90.0, 120.0))
        _color  = color or "Coral"

        if size == "대형 103×133mm":
            slide_w, slide_h = 297.0, 210.0   # 가로 A4
            layout_key = "대형_2개"
        else:
            slide_w, slide_h = 210.0, 297.0   # 세로 A4
            layout_key = "일반_4개" if layout == "A4에 4개" else "일반_2개"

        cols, rows = LANYARD_LAYOUTS.get(layout_key, (1, 2))

    else:
        return None

    # 프레젠테이션 생성 (방향 반영)
    prs = Presentation()
    prs.slide_width  = Mm(slide_w)
    prs.slide_height = Mm(slide_h)
    blank    = prs.slide_layouts[6]
    per_page = cols * rows

    for start in range(0, len(people), per_page):
        slide = prs.slides.add_slide(blank)
        chunk = people[start : start + per_page]

        if badge_type == "badge":
            place_on_slide(
                slide, chunk, layout_key, size_mm,
                design, _color, logo_path, company_name,
            )
        else:
            place_on_slide(
                slide, chunk, layout_key, size_mm,
                design, _color, logo_path, company_name,
                event_mode, event_info,
                slide_w=slide_w, slide_h=slide_h,
            )

    # 저장
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path  = str(Path(output_dir) / f"badges_{timestamp}.pptx")
    prs.save(out_path)

    # 폰트 임베딩 (PPT 열기 경고 방지를 위해 비활성화)
    # try:
    #     from font_embed import embed_fonts
    #     embed_fonts(out_path, BASE_DIR / "fonts")
    # except Exception as e:
    #     print(f"  [warn] 폰트 embed 실패: {e}")

    return out_path


def check_fonts() -> list[str]:
    """설치되지 않은 폰트 파일 목록 반환. 빈 리스트면 전부 OK."""
    import json
    spec_path = Path(__file__).parent / "designs" / "design_spec.json"
    with open(spec_path, encoding="utf-8") as f:
        spec = json.load(f)
    base    = Path(__file__).parent
    missing = []
    for rel in spec["common"]["fonts"].values():
        if not (base / rel).exists():
            missing.append(rel)
    return missing
