from __future__ import annotations

import zipfile
from dataclasses import dataclass

from summary_utils import NS, cm, collect_package_media_paths, open_zip, parse_relationships, parse_xml, read_zip_member, relationship_media_paths


@dataclass(frozen=True)
class DrawingInfo:
    """DOCX内の画像や図形の位置とサイズを保持するためのデータクラス。"""

    index: int
    label: str
    media_path: str
    x: int | None
    y: int | None
    cx: int
    cy: int


@dataclass
class DocxSummary:
    """DOCX要約に必要な本文、図形、メディア情報を保持するためのデータクラス。"""

    paragraphs: list[str]
    drawings: list[DrawingInfo]
    media_paths: set[str]


def is_docx_package(data: bytes) -> bool:
    """OOXMLバイト列がWord文書として扱えそうかを判定する関数。"""
    try:
        with open_zip(data) as archive:
            names = set(archive.namelist())
    except zipfile.BadZipFile:
        return False
    return "word/document.xml" in names


def _paragraph_texts(root) -> list[str]:
    """word/document.xmlから段落単位のテキスト一覧を取り出す関数。"""
    paragraphs: list[str] = []
    for paragraph in root.findall(".//w:p", NS):
        text = "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()
        if text:
            paragraphs.append(text)
    return paragraphs


def _position_offset(anchor, name: str) -> int | None:
    """wp:anchorから水平方向または垂直方向の位置オフセットを取り出す関数。"""
    position = anchor.find(f"wp:{name}/wp:posOffset", NS)
    if position is None or position.text is None:
        return None
    try:
        return int(position.text)
    except ValueError:
        return None


def _drawing_media_path(drawing, relationships) -> str:
    """w:drawing要素から参照している画像や動画パスを取り出す関数。"""
    blip = drawing.find(".//a:blip", NS)
    if blip is None:
        return ""
    rel_id = blip.attrib.get(f"{{{NS['r']}}}embed") or blip.attrib.get(f"{{{NS['r']}}}link") or ""
    relationship = relationships.get(rel_id)
    return relationship.target if relationship is not None else ""


def _collect_drawings(root, relationships) -> list[DrawingInfo]:
    """word/document.xmlから画像や図形のサイズ・位置情報を収集する関数。"""
    drawings: list[DrawingInfo] = []
    for index, drawing in enumerate(root.findall(".//w:drawing", NS), start=1):
        container = drawing.find(".//wp:anchor", NS) or drawing.find(".//wp:inline", NS)
        if container is None:
            continue
        extent = container.find("wp:extent", NS)
        if extent is None:
            continue
        doc_pr = container.find("wp:docPr", NS)
        label = doc_pr.attrib.get("name", f"drawing{index}") if doc_pr is not None else f"drawing{index}"
        try:
            cx = int(extent.attrib.get("cx", "0"))
            cy = int(extent.attrib.get("cy", "0"))
        except ValueError:
            continue
        drawings.append(
            DrawingInfo(
                index=index,
                label=label,
                media_path=_drawing_media_path(drawing, relationships),
                x=_position_offset(container, "positionH"),
                y=_position_offset(container, "positionV"),
                cx=cx,
                cy=cy,
            )
        )
    return drawings


def _summarize_package(data: bytes) -> DocxSummary:
    """DOCXバイト列から要約に必要な情報を抽出する関数。"""
    with open_zip(data) as archive:
        document_xml = read_zip_member(archive, "word/document.xml")
        root = parse_xml(document_xml or b"")
        relationships = parse_relationships(archive, "word/document.xml")
        media_paths = collect_package_media_paths(archive, "word/media/")
        media_paths.update(relationship_media_paths(archive, "word/document.xml"))
        if root is None:
            return DocxSummary(paragraphs=[], drawings=[], media_paths=media_paths)
        return DocxSummary(
            paragraphs=_paragraph_texts(root),
            drawings=_collect_drawings(root, relationships),
            media_paths=media_paths,
        )


def _text_lines(summary: DocxSummary) -> list[str]:
    """DOCXの本文テキスト一覧行を作る関数。"""
    lines = ["## 文書テキスト一覧"]
    if not summary.paragraphs:
        return lines + ["なし"]
    return lines + [f"TEXT p.{index}: {text}" for index, text in enumerate(summary.paragraphs, start=1)]


def _drawing_lines(summary: DocxSummary) -> list[str]:
    """DOCXの画像や図形の位置・サイズ一覧行を作る関数。"""
    lines = ["## 位置・サイズ一覧"]
    if not summary.drawings:
        return lines + ["なし"]
    for drawing in summary.drawings:
        position = "なし"
        if drawing.x is not None or drawing.y is not None:
            x = cm(drawing.x or 0)
            y = cm(drawing.y or 0)
            position = f"({x}, {y})"
        lines.append(
            f"DRAWING #{drawing.index} label={drawing.label} media={drawing.media_path or 'なし'}: "
            f"位置={position} サイズ=({cm(drawing.cx)}, {cm(drawing.cy)}) emu=({drawing.x},{drawing.y},{drawing.cx},{drawing.cy})"
        )
    return lines


def _media_lines(summary: DocxSummary) -> list[str]:
    """DOCXの画像や動画パス一覧行を作る関数。"""
    lines = ["## 画像・動画一覧"]
    return lines + [f"MEDIA {path}" for path in sorted(summary.media_paths)] if summary.media_paths else lines + ["なし"]


def summarize_docx_inventory(data: bytes) -> str:
    """1つのDOCXからgit textconv向けの人が読みやすい要約を返す関数。"""
    summary = _summarize_package(data)
    sections = [_text_lines(summary), _drawing_lines(summary), _media_lines(summary)]
    return "\n\n".join("\n".join(section) for section in sections)
