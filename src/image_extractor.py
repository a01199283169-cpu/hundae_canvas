"""엑셀 내장 이미지 추출 및 주문 행 매핑."""

from __future__ import annotations

import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path

from src.config_loader import ensure_dirs, load_config
from src.database import connect, insert_order_image, mark_order_has_image


NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}
R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"


@dataclass
class ImageAnchor:
    sheet_name: str
    excel_row: int  # 1-based
    excel_col: int  # 1-based
    media_path: str
    media_filename: str


def _sheet_index_map(z: zipfile.ZipFile) -> dict[str, int]:
    """시트명 → sheetN.xml 번호."""
    wb_xml = z.read("xl/workbook.xml")
    root = ET.fromstring(wb_xml)
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    sheets = root.findall(".//m:sheet", ns)
    mapping = {}
    for i, sheet in enumerate(sheets, start=1):
        name = sheet.attrib.get("name", "")
        mapping[name] = i
    return mapping


def _drawing_targets(z: zipfile.ZipFile, sheet_idx: int) -> list[str]:
    """시트에 연결된 drawing XML 경로."""
    rel_path = f"xl/worksheets/_rels/sheet{sheet_idx}.xml.rels"
    if rel_path not in z.namelist():
        return []
    rel_root = ET.fromstring(z.read(rel_path))
    drawings = []
    for rel in rel_root.findall("rel:Relationship", NS):
        if "drawing" in rel.attrib.get("Type", ""):
            target = rel.attrib["Target"].replace("../", "xl/")
            drawings.append(target)
    return drawings


def _resolve_media(z: zipfile.ZipFile, drawing_path: str) -> dict[str, str]:
    """drawing rId → media 파일 경로."""
    rel_path = drawing_path.replace("drawings/", "drawings/_rels/") + ".rels"
    if rel_path not in z.namelist():
        return {}
    rel_root = ET.fromstring(z.read(rel_path))
    rid_map = {}
    for rel in rel_root.findall("rel:Relationship", NS):
        target = rel.attrib.get("Target", "")
        if "media/" in target:
            rid_map[rel.attrib["Id"]] = "xl/" + target.replace("../", "")
    return rid_map


def extract_image_anchors(xlsx_path: Path) -> list[ImageAnchor]:
    """xlsx에서 이미지 앵커(행/열) 목록 추출."""
    anchors: list[ImageAnchor] = []
    with zipfile.ZipFile(xlsx_path, "r") as z:
        sheet_map = _sheet_index_map(z)
        # sheet index → name 역매핑
        idx_to_name = {v: k for k, v in sheet_map.items()}

        for sheet_idx, sheet_name in idx_to_name.items():
            for drawing_path in _drawing_targets(z, sheet_idx):
                if drawing_path not in z.namelist():
                    continue
                rid_media = _resolve_media(z, drawing_path)
                drawing_xml = z.read(drawing_path)
                root = ET.fromstring(drawing_xml)

                for anchor in root.findall(".//xdr:twoCellAnchor", NS):
                    from_el = anchor.find("xdr:from", NS)
                    if from_el is None:
                        continue
                    row_el = from_el.find("xdr:row", NS)
                    col_el = from_el.find("xdr:col", NS)
                    if row_el is None or col_el is None:
                        continue
                    excel_row = int(row_el.text) + 1  # 0-based → 1-based
                    excel_col = int(col_el.text) + 1

                    blip = anchor.find(".//a:blip", NS)
                    if blip is None:
                        continue
                    embed = blip.attrib.get(R_EMBED)
                    media_path = rid_media.get(embed or "")
                    if not media_path:
                        continue
                    anchors.append(
                        ImageAnchor(
                            sheet_name=sheet_name,
                            excel_row=excel_row,
                            excel_col=excel_col,
                            media_path=media_path,
                            media_filename=Path(media_path).name,
                        )
                    )
    return anchors


def extract_and_save_images(
    xlsx_path: Path,
    row_order_map: dict[tuple[str, int], int] | None = None,
) -> dict[str, int]:
    """
    이미지 파일 추출 및 주문 매핑.
    row_order_map: (sheet_name, excel_row) → order_id
    """
    config = load_config()
    paths = ensure_dirs(config)
    source_file = xlsx_path.name
    anchors = extract_image_anchors(xlsx_path)

    stats = {"extracted": 0, "mapped": 0, "unmapped": 0}
    conn = connect()

    with zipfile.ZipFile(xlsx_path, "r") as z:
        for idx, anchor in enumerate(anchors, start=1):
            if anchor.media_path not in z.namelist():
                continue

            # 출력 폴더: images/{시트}_{행}/
            folder_name = f"{anchor.sheet_name}_row{anchor.excel_row:03d}"
            out_dir = paths["images"] / source_file.replace(".xlsx", "") / folder_name
            out_dir.mkdir(parents=True, exist_ok=True)

            ext = Path(anchor.media_filename).suffix or ".png"
            out_file = out_dir / f"image_{idx}{ext}"

            with z.open(anchor.media_path) as src, out_file.open("wb") as dst:
                shutil.copyfileobj(src, dst)

            stats["extracted"] += 1

            # 주문 매핑
            order_id = None
            mapped = False
            if row_order_map:
                order_id = row_order_map.get((anchor.sheet_name, anchor.excel_row))
                if order_id is None:
                    # 같은 행 또는 인접 행(±2) 주문 탐색
                    for delta in (0, -1, 1, -2, 2):
                        order_id = row_order_map.get((anchor.sheet_name, anchor.excel_row + delta))
                        if order_id:
                            break

            if order_id:
                mark_order_has_image(conn, order_id)
                mapped = True
                stats["mapped"] += 1
            else:
                stats["unmapped"] += 1

            insert_order_image(
                conn,
                order_id=order_id,
                source_file=source_file,
                sheet_name=anchor.sheet_name,
                excel_row=anchor.excel_row,
                image_file=str(out_file),
                mapped=mapped,
            )

    conn.close()
    return stats


def build_row_order_map(xlsx_path: Path) -> dict[tuple[str, int], int]:
    """파싱된 주문의 start_row·item rows → order_id 매핑 생성."""
    from src.parser import parse_workbook

    conn = connect()
    parsed = parse_workbook(xlsx_path)
    source_file = xlsx_path.name
    row_map: dict[tuple[str, int], int] = {}

    for sheet_name, orders in parsed.items():
        for order in orders:
            row = conn.execute(
                """
                SELECT id FROM orders
                WHERE source_file=? AND sheet_name=? AND order_no=?
                """,
                (source_file, sheet_name, order["order_no"]),
            ).fetchone()
            if not row:
                continue
            order_id = int(row["id"])
            start = order.get("start_row", 0)
            if start:
                row_map[(sheet_name, start)] = order_id
            for item in order.get("items", []):
                er = item.get("excel_row")
                if er:
                    row_map[(sheet_name, er)] = order_id

    conn.close()
    return row_map
