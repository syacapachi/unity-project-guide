from __future__ import annotations

import zipfile
from dataclasses import dataclass

from summary_utils import NS, cell_reference, cm, collect_package_media_paths, local_name, open_zip, parse_relationships, parse_xml, read_zip_member, relationship_media_paths


@dataclass(frozen=True)
class SheetInfo:
    """XLSX内のシート名、シートID、XMLパスを保持するためのデータクラス。"""

    name: str
    sheet_id: str
    path: str


@dataclass(frozen=True)
class CellInfo:
    """XLSX内のセル値と数式を保持するためのデータクラス。"""

    sheet_name: str
    ref: str
    value: str
    formula: str


@dataclass(frozen=True)
class DrawingInfo:
    """XLSX内の描画オブジェクトの位置とサイズを保持するためのデータクラス。"""

    sheet_name: str
    object_id: str
    label: str
    media_path: str
    from_cell: str
    to_cell: str
    cx: int | None
    cy: int | None


@dataclass
class XlsxSummary:
    """XLSX要約に必要なシート、セル、描画、メディア情報を保持するためのデータクラス。"""

    sheets: list[SheetInfo]
    cells: list[CellInfo]
    drawings: list[DrawingInfo]
    media_paths: set[str]


def is_xlsx_package(data: bytes) -> bool:
    """OOXMLバイト列がExcelブックとして扱えそうかを判定する関数。"""
    try:
        with open_zip(data) as archive:
            names = set(archive.namelist())
    except zipfile.BadZipFile:
        return False
    return "xl/workbook.xml" in names


def _shared_strings(archive: zipfile.ZipFile) -> list[str]:
    """sharedStrings.xmlから共有文字列テーブルを読み込む関数。"""
    data = read_zip_member(archive, "xl/sharedStrings.xml")
    root = parse_xml(data or b"")
    if root is None:
        return []
    values: list[str] = []
    for item in root.findall(".//x:si", NS):
        values.append("".join(node.text or "" for node in item.findall(".//x:t", NS)))
    return values


def _parse_sheets(archive: zipfile.ZipFile) -> list[SheetInfo]:
    """workbook.xmlとRelationshipからシート一覧を読み込む関数。"""
    root = parse_xml(read_zip_member(archive, "xl/workbook.xml") or b"")
    if root is None:
        return []
    rels = parse_relationships(archive, "xl/workbook.xml")
    sheets: list[SheetInfo] = []
    for element in root.findall(".//x:sheet", NS):
        rel_id = element.attrib.get(f"{{{NS['r']}}}id", "")
        relationship = rels.get(rel_id)
        if relationship is None:
            continue
        sheets.append(
            SheetInfo(
                name=element.attrib.get("name", relationship.target),
                sheet_id=element.attrib.get("sheetId", relationship.target),
                path=relationship.target,
            )
        )
    return sheets


def _cell_value(cell, shared_strings: list[str]) -> tuple[str, str]:
    """セルXMLから表示値と数式を取り出す関数。"""
    formula_node = cell.find("x:f", NS)
    formula = formula_node.text or "" if formula_node is not None else ""
    value_node = cell.find("x:v", NS)
    raw_value = value_node.text or "" if value_node is not None else ""
    cell_type = cell.attrib.get("t", "")
    if cell_type == "s":
        try:
            return shared_strings[int(raw_value)], formula
        except (ValueError, IndexError):
            return raw_value, formula
    if cell_type == "inlineStr":
        text = "".join(node.text or "" for node in cell.findall(".//x:t", NS))
        return text, formula
    return raw_value, formula


def _parse_cells(archive: zipfile.ZipFile, sheet: SheetInfo, shared_strings: list[str]) -> list[CellInfo]:
    """ワークシートXMLから値や数式があるセルを収集する関数。"""
    root = parse_xml(read_zip_member(archive, sheet.path) or b"")
    if root is None:
        return []
    cells: list[CellInfo] = []
    for cell in root.findall(".//x:c", NS):
        value, formula = _cell_value(cell, shared_strings)
        if not value and not formula:
            continue
        cells.append(CellInfo(sheet_name=sheet.name, ref=cell.attrib.get("r", ""), value=value, formula=formula))
    return cells


def _marker_cell(marker) -> str:
    """xdr:from/xdr:toマーカーをExcelセル参照へ変換する関数。"""
    try:
        column = int(marker.find("xdr:col", NS).text or "0")
        row = int(marker.find("xdr:row", NS).text or "0")
    except (AttributeError, ValueError):
        return ""
    return cell_reference(column, row)


def _drawing_media_path(element, relationships) -> str:
    """xdr描画要素から参照している画像や動画パスを取り出す関数。"""
    blip = element.find(".//a:blip", NS)
    if blip is None:
        return ""
    rel_id = blip.attrib.get(f"{{{NS['r']}}}embed") or blip.attrib.get(f"{{{NS['r']}}}link") or ""
    relationship = relationships.get(rel_id)
    return relationship.target if relationship is not None else ""


def _parse_drawing_file(archive: zipfile.ZipFile, sheet_name: str, drawing_path: str) -> list[DrawingInfo]:
    """drawing*.xmlから画像や図形の位置・サイズ情報を収集する関数。"""
    root = parse_xml(read_zip_member(archive, drawing_path) or b"")
    if root is None:
        return []
    relationships = parse_relationships(archive, drawing_path)
    drawings: list[DrawingInfo] = []
    anchors = [element for element in root if local_name(element.tag) in {"twoCellAnchor", "oneCellAnchor", "absoluteAnchor"}]
    for index, anchor in enumerate(anchors, start=1):
        c_nv_pr = anchor.find(".//xdr:cNvPr", NS)
        object_id = c_nv_pr.attrib.get("id", str(index)) if c_nv_pr is not None else str(index)
        label = c_nv_pr.attrib.get("name", f"drawing{index}") if c_nv_pr is not None else f"drawing{index}"
        from_node = anchor.find("xdr:from", NS)
        to_node = anchor.find("xdr:to", NS)
        extent = anchor.find("xdr:ext", NS)
        cx = cy = None
        if extent is not None:
            try:
                cx = int(extent.attrib.get("cx", "0"))
                cy = int(extent.attrib.get("cy", "0"))
            except ValueError:
                cx = cy = None
        drawings.append(
            DrawingInfo(
                sheet_name=sheet_name,
                object_id=object_id,
                label=label,
                media_path=_drawing_media_path(anchor, relationships),
                from_cell=_marker_cell(from_node) if from_node is not None else "",
                to_cell=_marker_cell(to_node) if to_node is not None else "",
                cx=cx,
                cy=cy,
            )
        )
    return drawings


def _sheet_drawing_paths(archive: zipfile.ZipFile, sheet_path: str) -> list[str]:
    """ワークシートのRelationshipからdrawing*.xmlパスを収集する関数。"""
    paths: list[str] = []
    for relationship in parse_relationships(archive, sheet_path).values():
        if "drawing" in relationship.rel_type.lower() or relationship.target.startswith("xl/drawings/"):
            paths.append(relationship.target)
    return paths


def _summarize_package(data: bytes) -> XlsxSummary:
    """XLSXバイト列から要約に必要な情報を抽出する関数。"""
    with open_zip(data) as archive:
        shared_strings = _shared_strings(archive)
        sheets = _parse_sheets(archive)
        cells: list[CellInfo] = []
        drawings: list[DrawingInfo] = []
        media_paths = collect_package_media_paths(archive, "xl/media/")
        for sheet in sheets:
            cells.extend(_parse_cells(archive, sheet, shared_strings))
            media_paths.update(relationship_media_paths(archive, sheet.path))
            for drawing_path in _sheet_drawing_paths(archive, sheet.path):
                media_paths.update(relationship_media_paths(archive, drawing_path))
                drawings.extend(_parse_drawing_file(archive, sheet.name, drawing_path))
        return XlsxSummary(sheets=sheets, cells=cells, drawings=drawings, media_paths=media_paths)


def _sheet_lines(summary: XlsxSummary) -> list[str]:
    """XLSXのシート一覧行を作る関数。"""
    lines = ["## シート一覧"]
    return lines + [f"SHEET {sheet.name} id={sheet.sheet_id} ({sheet.path})" for sheet in summary.sheets] if summary.sheets else lines + ["なし"]


def _cell_lines(summary: XlsxSummary) -> list[str]:
    """XLSXのセル値と数式一覧行を作る関数。"""
    lines = ["## セル一覧"]
    if not summary.cells:
        return lines + ["なし"]
    for cell in summary.cells:
        formula = f" formula={cell.formula}" if cell.formula else ""
        lines.append(f"CELL {cell.sheet_name}!{cell.ref}: {cell.value}{formula}")
    return lines


def _drawing_lines(summary: XlsxSummary) -> list[str]:
    """XLSXの画像や図形の位置・サイズ一覧行を作る関数。"""
    lines = ["## 位置・サイズ一覧"]
    if not summary.drawings:
        return lines + ["なし"]
    for drawing in summary.drawings:
        size = "なし"
        if drawing.cx is not None and drawing.cy is not None:
            size = f"({cm(drawing.cx)}, {cm(drawing.cy)})"
        lines.append(
            f"DRAWING {drawing.sheet_name} object={drawing.object_id} label={drawing.label} media={drawing.media_path or 'なし'}: "
            f"範囲={drawing.from_cell}->{drawing.to_cell} サイズ={size}"
        )
    return lines


def _media_lines(summary: XlsxSummary) -> list[str]:
    """XLSXの画像や動画パス一覧行を作る関数。"""
    lines = ["## 画像・動画一覧"]
    return lines + [f"MEDIA {path}" for path in sorted(summary.media_paths)] if summary.media_paths else lines + ["なし"]


def summarize_xlsx_inventory(data: bytes) -> str:
    """1つのXLSXからgit textconv向けの人が読みやすい要約を返す関数。"""
    summary = _summarize_package(data)
    sections = [_sheet_lines(summary), _cell_lines(summary), _drawing_lines(summary), _media_lines(summary)]
    return "\n\n".join("\n".join(section) for section in sections)
