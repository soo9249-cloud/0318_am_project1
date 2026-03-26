"""
font_embed.py — OOXML 폰트 임베딩 (ECMA-376 §22.5.2)

PPTX 저장 후 호출하면 PPT 안에 폰트 바이너리를 포함시켜
다른 컴퓨터에서 열어도 폰트가 그대로 표시됩니다.
"""
from __future__ import annotations

import uuid
import zipfile
from pathlib import Path

from lxml import etree

_P_NS   = "http://schemas.openxmlformats.org/presentationml/2006/main"
_A_NS   = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R_NS   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS  = "http://schemas.openxmlformats.org/package/2006/content-types"

_FONT_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
)
_FONT_CT = "application/vnd.openxmlformats-officedocument.obfuscatedFont"


def _obfuscate(font_data: bytes, guid: str) -> bytes:
    """ECMA-376 §22.5.2 폰트 XOR 난독화.

    guid : '{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}' 형식 GUID
    처음 32 바이트를 GUID 역순 키로 XOR.
    """
    hex_str = guid.replace("{", "").replace("}", "").replace("-", "")
    if len(hex_str) != 32:
        return font_data                    # 잘못된 GUID — 그대로 반환
    key    = bytes.fromhex(hex_str)[::-1]   # 16 바이트 역순
    result = bytearray(font_data)
    for i in range(min(32, len(result))):
        result[i] ^= key[i % 16]
    return bytes(result)


# 임베드할 폰트 목록: (타입페이스명, 변형, fonts/ 내 파일명)
_EMBED_LIST: list[tuple[str, str, str]] = [
    ("Noto Serif KR", "regular", "NotoSerifKR-Bold.ttf"),
    ("Noto Serif KR", "bold",    "NotoSerifKR-Bold.ttf"),
    ("Noto Sans KR",  "regular", "NotoSansKR-Regular.ttf"),
    ("Montserrat",    "regular", "Montserrat-Regular.ttf"),
    ("Montserrat",    "bold",    "Montserrat-Bold.ttf"),
]


def embed_fonts(pptx_path: str, fonts_dir: Path) -> None:
    """
    PPTX 파일에 폰트를 embed.

    pptx_path : 저장된 .pptx 경로 (이미 존재해야 함)
    fonts_dir : fonts/ 폴더 경로
    """
    embeds = [
        (tf, var, fonts_dir / fn)
        for tf, var, fn in _EMBED_LIST
        if (fonts_dir / fn).exists()
    ]
    if not embeds:
        return

    # ── PPTX(ZIP) 읽기
    with zipfile.ZipFile(pptx_path, "r") as zin:
        files = {n: zin.read(n) for n in zin.namelist()}

    prs_root = etree.fromstring(files["ppt/presentation.xml"])

    rels_key  = "ppt/_rels/presentation.xml.rels"
    rels_root = (etree.fromstring(files[rels_key]) if rels_key in files
                 else etree.Element(f"{{{_PKG_NS}}}Relationships"))

    ct_root = etree.fromstring(files["[Content_Types].xml"])

    # ── embeddedFontLst 찾기 / 생성
    font_lst = prs_root.find(f"{{{_P_NS}}}embeddedFontLst")
    if font_lst is None:
        font_lst = etree.SubElement(prs_root, f"{{{_P_NS}}}embeddedFontLst")

    def _get_or_make_ef(typeface: str):
        """타입페이스에 해당하는 p:embeddedFont 요소 반환 (없으면 생성)."""
        for ef in font_lst:
            fe = ef.find(f"{{{_A_NS}}}font")
            if fe is not None and fe.get("typeface") == typeface:
                return ef
        ef = etree.SubElement(font_lst, f"{{{_P_NS}}}embeddedFont")
        fe = etree.SubElement(ef, f"{{{_A_NS}}}font")
        fe.set("typeface", typeface)
        fe.set("panose",       "00000000000000000000")
        fe.set("pitchFamily",  "0")
        fe.set("charset",      "0")
        return ef

    # ── 각 폰트 처리
    for part_idx, (typeface, variant, font_path) in enumerate(embeds, start=1):
        guid      = "{" + str(uuid.uuid4()).upper() + "}"
        font_data = font_path.read_bytes()
        obf_data  = _obfuscate(font_data, guid)

        part_name = f"font{part_idx}"
        part_path = f"ppt/fonts/{part_name}.fntdata"
        files[part_path] = obf_data

        # 관계 추가
        rel = etree.SubElement(rels_root, f"{{{_PKG_NS}}}Relationship")
        rel.set("Id",     guid)
        rel.set("Type",   _FONT_REL_TYPE)
        rel.set("Target", f"fonts/{part_name}.fntdata")

        # Content_Types 추가
        ov = etree.SubElement(ct_root, f"{{{_CT_NS}}}Override")
        ov.set("PartName",    f"/{part_path}")
        ov.set("ContentType", _FONT_CT)

        # embeddedFont 변형 요소 추가
        ef      = _get_or_make_ef(typeface)
        var_el  = etree.SubElement(ef, f"{{{_P_NS}}}{variant}")
        var_el.set(f"{{{_R_NS}}}id", guid)

    # ── XML 직렬화
    xml_opts = dict(xml_declaration=True, encoding="UTF-8", standalone=True)
    files["ppt/presentation.xml"]   = etree.tostring(prs_root,  **xml_opts)
    files[rels_key]                  = etree.tostring(rels_root, **xml_opts)
    files["[Content_Types].xml"]     = etree.tostring(ct_root,   **xml_opts)

    # ── PPTX(ZIP) 다시 쓰기
    with zipfile.ZipFile(pptx_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)

    print(f"  [embed] {len(embeds)}개 폰트 embed 완료 → {pptx_path}")
