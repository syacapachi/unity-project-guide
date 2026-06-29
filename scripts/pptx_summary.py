from __future__ import annotations

import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree

from summary_utils import (
    NS,
    cm,
    collect_package_media_paths,
    local_name,
    open_zip,
    parse_relationships,
    parse_xml,
    ratio,
    read_zip_member,
    relationship_media_paths,
)


@dataclass(frozen=True)
class SlideInfo:
    """PPTX内のスライド番号、スライドID、XMLパスを保持するためのデータクラス。"""

    page: int
    slide_id: str
    path: str


@dataclass(frozen=True)
class TextItem:
    """スライド上のテキストと、その出現位置を保持するためのデータクラス。"""

    slide_key: str
    page: int
    slide_id: str
    text: str


@dataclass(frozen=True)
class TransformItem:
    """スライド上の図形や画像の位置とサイズを保持するためのデータクラス。"""

    slide_key: str
    page: int
    slide_id: str
    object_key: str
    label: str
    x: int
    y: int
    cx: int
    cy: int


@dataclass
class PptxSummary:
    """PPTX要約に必要な情報をまとめて保持するためのデータクラス。"""

    slides: dict[str, SlideInfo]
    slide_order: list[str]
    texts: list[TextItem]
    transforms: dict[tuple[str, str], TransformItem]
    media_paths: set[str]


def is_pptx_package(data: bytes) -> bool:
    """OOXMLバイト列がPowerPointプレゼンテーションとして扱えそうかを判定する関数。"""
    try:
        with open_zip(data) as archive:
            names = set(archive.namelist())
    except zipfile.BadZipFile:
        return False
    return "ppt/presentation.xml" in names or any(name.startswith("ppt/slides/slide") for name in names)


def _fallback_slides(archive: zipfile.ZipFile) -> list[SlideInfo]:
    """presentation.xmlが読めない場合にslide*.xmlの名前からスライド一覧を作る関数。"""
    paths = sorted(name for name in archive.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml"))
    return [SlideInfo(page=index + 1, slide_id=path, path=path) for index, path in enumerate(paths)]


def _parse_slides(archive: zipfile.ZipFile) -> list[SlideInfo]:
    """presentation.xmlとRelationshipからスライド順、ページ番号、スライドIDを読み取る関数。"""
    data = read_zip_member(archive, "ppt/presentation.xml")
    if data is None:
        return _fallback_slides(archive)

    root = parse_xml(data)
    if root is None:
        return _fallback_slides(archive)

    rels = parse_relationships(archive, "ppt/presentation.xml")
    slides: list[SlideInfo] = []
    for index, element in enumerate(root.findall(".//p:sldId", NS), start=1):
        rel_id = element.attrib.get(f"{{{NS['r']}}}id", "")
        relationship = rels.get(rel_id)
        if relationship is None:
            continue
        slides.append(SlideInfo(page=index, slide_id=element.attrib.get("id", relationship.target), path=relationship.target))
    return slides or _fallback_slides(archive)


def _slide_key(slide: SlideInfo) -> str:
    """差分比較で使うスライドの安定キーを作る関数。"""
    return slide.slide_id or slide.path


def _collect_texts(root: ElementTree.Element, slide: SlideInfo, slide_key: str) -> list[TextItem]:
    """スライドXMLからテキスト要素を収集する関数。"""
    items: list[TextItem] = []
    for element in root.findall(".//a:t", NS):
        text = "".join(element.itertext()).strip()
        if text:
            items.append(TextItem(slide_key=slide_key, page=slide.page, slide_id=slide.slide_id, text=text))
    return items


def _first_text_preview(element: ElementTree.Element) -> str:
    """図形ラベルに使う短いテキストプレビューを取り出す関数。"""
    texts = ["".join(node.itertext()).strip() for node in element.findall(".//a:t", NS)]
    preview = " ".join(text for text in texts if text)
    return preview[:29] + "..." if len(preview) > 30 else preview


def _object_label(element: ElementTree.Element, fallback_index: int) -> tuple[str, str]:
    """図形や画像を比較するためのキーと表示名を作る関数。"""
    c_nv_pr = element.find(".//p:cNvPr", NS)
    object_id = c_nv_pr.attrib.get("id", "") if c_nv_pr is not None else ""
    object_name = c_nv_pr.attrib.get("name", "") if c_nv_pr is not None else ""
    preview = _first_text_preview(element)
    object_type = local_name(element.tag)
    label_parts = [part for part in [object_name, preview] if part]
    label = " / ".join(label_parts) if label_parts else f"{object_type}#{object_id or fallback_index}"
    key = object_id or f"{object_type}:{object_name}:{preview}:{fallback_index}"
    return key, label


def _transform_values(element: ElementTree.Element) -> tuple[int, int, int, int] | None:
    """図形や画像のa:xfrmから位置とサイズを取り出す関数。"""
    transform = element.find(".//a:xfrm", NS)
    if transform is None:
        return None
    offset = transform.find("a:off", NS)
    extent = transform.find("a:ext", NS)
    if offset is None or extent is None:
        return None
    try:
        return (
            int(offset.attrib.get("x", "0")),
            int(offset.attrib.get("y", "0")),
            int(extent.attrib.get("cx", "0")),
            int(extent.attrib.get("cy", "0")),
        )
    except ValueError:
        return None


def _collect_transforms(root: ElementTree.Element, slide: SlideInfo, slide_key: str) -> dict[tuple[str, str], TransformItem]:
    """スライドXMLから位置や拡大率の比較対象になる図形情報を収集する関数。"""
    transforms: dict[tuple[str, str], TransformItem] = {}
    candidates = [element for element in root.iter() if local_name(element.tag) in {"sp", "pic", "graphicFrame", "cxnSp"}]
    for index, element in enumerate(candidates, start=1):
        values = _transform_values(element)
        if values is None:
            continue
        object_key, label = _object_label(element, index)
        transforms[(slide_key, object_key)] = TransformItem(
            slide_key=slide_key,
            page=slide.page,
            slide_id=slide.slide_id,
            object_key=object_key,
            label=label,
            x=values[0],
            y=values[1],
            cx=values[2],
            cy=values[3],
        )
    return transforms


def _summarize_package(data: bytes) -> PptxSummary:
    """PPTXバイト列から要約に必要な情報を抽出する関数。"""
    slides: dict[str, SlideInfo] = {}
    slide_order: list[str] = []
    texts: list[TextItem] = []
    transforms: dict[tuple[str, str], TransformItem] = {}
    media_paths: set[str] = set()

    with open_zip(data) as archive:
        media_paths.update(collect_package_media_paths(archive, "ppt/media/"))
        for slide in _parse_slides(archive):
            slide_key = _slide_key(slide)
            slides[slide_key] = slide
            slide_order.append(slide_key)
            slide_xml = read_zip_member(archive, slide.path)
            if slide_xml is None:
                continue
            root = parse_xml(slide_xml)
            if root is None:
                continue
            texts.extend(_collect_texts(root, slide, slide_key))
            transforms.update(_collect_transforms(root, slide, slide_key))
            media_paths.update(relationship_media_paths(archive, slide.path))

    return PptxSummary(slides=slides, slide_order=slide_order, texts=texts, transforms=transforms, media_paths=media_paths)


def _format_slide(slide: SlideInfo) -> str:
    """スライド情報をページ番号とID付きの文字列へ変換する関数。"""
    return f"p.{slide.page} id={slide.slide_id} ({slide.path})"


def _format_text_item(item: TextItem) -> str:
    """テキスト差分の1項目を表示用文字列へ変換する関数。"""
    return f"p.{item.page} id={item.slide_id}: {item.text}"


def _inventory_slide_lines(summary: PptxSummary) -> list[str]:
    """1ファイル要約用のスライド一覧行を作る関数。"""
    lines = ["## スライド一覧"]
    return lines + [f"SLIDE {_format_slide(summary.slides[key])}" for key in summary.slide_order] if summary.slide_order else lines + ["なし"]


def _inventory_text_lines(summary: PptxSummary) -> list[str]:
    """1ファイル要約用のテキスト一覧行を作る関数。"""
    lines = ["## テキスト一覧"]
    return lines + [f"TEXT {_format_text_item(item)}" for item in summary.texts] if summary.texts else lines + ["なし"]


def _inventory_transform_lines(summary: PptxSummary) -> list[str]:
    """1ファイル要約用の位置とサイズ一覧行を作る関数。"""
    lines = ["## 位置・サイズ一覧"]
    if not summary.transforms:
        return lines + ["なし"]
    for item in sorted(summary.transforms.values(), key=lambda value: (value.page, value.object_key)):
        lines.append(
            f"XFRM p.{item.page} id={item.slide_id} object={item.object_key} label={item.label}: "
            f"位置=({cm(item.x)}, {cm(item.y)}) サイズ=({cm(item.cx)}, {cm(item.cy)}) "
            f"emu=({item.x},{item.y},{item.cx},{item.cy})"
        )
    return lines


def _inventory_media_lines(summary: PptxSummary) -> list[str]:
    """1ファイル要約用の画像と動画パス一覧行を作る関数。"""
    lines = ["## 画像・動画一覧"]
    return lines + [f"MEDIA {path}" for path in sorted(summary.media_paths)] if summary.media_paths else lines + ["なし"]


def summarize_pptx_inventory(data: bytes) -> str:
    """1つのPPTXからgit textconv向けの人が読みやすい要約を返す関数。"""
    summary = _summarize_package(data)
    sections = [
        _inventory_slide_lines(summary),
        _inventory_text_lines(summary),
        _inventory_transform_lines(summary),
        _inventory_media_lines(summary),
    ]
    return "\n\n".join("\n".join(section) for section in sections)


def _counter_items(items: list[TextItem]) -> Counter[str]:
    """テキスト項目の出現回数を数える関数。"""
    return Counter(item.text for item in items)


def _index_text_items(items: list[TextItem]) -> dict[str, list[TextItem]]:
    """テキスト文字列から出現位置一覧へ引ける辞書を作る関数。"""
    indexed: dict[str, list[TextItem]] = defaultdict(list)
    for item in items:
        indexed[item.text].append(item)
    return indexed


def _text_diff_lines(old: PptxSummary, new: PptxSummary) -> list[str]:
    """テキストの追加と削除を要約する行リストを作る関数。"""
    lines = ["## テキスト差分"]
    old_counter = _counter_items(old.texts)
    new_counter = _counter_items(new.texts)
    old_index = _index_text_items(old.texts)
    new_index = _index_text_items(new.texts)
    additions = sorted((new_counter - old_counter).elements())
    deletions = sorted((old_counter - new_counter).elements())
    if not additions and not deletions:
        return lines + ["変更なし"]
    for text in additions:
        lines.append(f"+ {_format_text_item(new_index[text].pop(0))}")
    for text in deletions:
        lines.append(f"- {_format_text_item(old_index[text].pop(0))}")
    return lines


def _slide_diff_lines(old: PptxSummary, new: PptxSummary) -> list[str]:
    """スライドの追加と削除を要約する行リストを作る関数。"""
    lines = ["## スライド差分"]
    old_keys = set(old.slides)
    new_keys = set(new.slides)
    added = [new.slides[key] for key in new.slide_order if key in new_keys - old_keys]
    deleted = [old.slides[key] for key in old.slide_order if key in old_keys - new_keys]
    if not added and not deleted:
        return lines + ["変更なし"]
    return lines + [f"+ {_format_slide(slide)}" for slide in added] + [f"- {_format_slide(slide)}" for slide in deleted]


def _transform_diff_lines(old: PptxSummary, new: PptxSummary) -> list[str]:
    """位置と拡大率の差分を要約する行リストを作る関数。"""
    lines = ["## 位置・拡大率差分"]
    changed: list[str] = []
    for key in sorted(set(old.transforms) & set(new.transforms)):
        old_item = old.transforms[key]
        new_item = new.transforms[key]
        if (old_item.x, old_item.y, old_item.cx, old_item.cy) == (new_item.x, new_item.y, new_item.cx, new_item.cy):
            continue
        changed.append(
            f"* p.{new_item.page} id={new_item.slide_id} {new_item.label}: "
            f"位置 ({cm(old_item.x)}, {cm(old_item.y)}) -> ({cm(new_item.x)}, {cm(new_item.y)}), "
            f"サイズ ({cm(old_item.cx)}, {cm(old_item.cy)}) -> ({cm(new_item.cx)}, {cm(new_item.cy)}), "
            f"拡大率 ({ratio(old_item.cx, new_item.cx)}, {ratio(old_item.cy, new_item.cy)})"
        )
    return lines + (changed if changed else ["変更なし"])


def _media_diff_lines(old: PptxSummary, new: PptxSummary) -> list[str]:
    """画像や動画の追加と削除を要約する行リストを作る関数。"""
    lines = ["## 画像・動画差分"]
    added = sorted(new.media_paths - old.media_paths)
    deleted = sorted(old.media_paths - new.media_paths)
    if not added and not deleted:
        return lines + ["変更なし"]
    return lines + [f"+ {path}" for path in added] + [f"- {path}" for path in deleted]


def summarize_pptx_diff(old_data: bytes, new_data: bytes) -> str:
    """2つのPPTXバイト列を比較し、人が読みやすい差分要約を返す関数。"""
    old = _summarize_package(old_data)
    new = _summarize_package(new_data)
    sections = [
        _slide_diff_lines(old, new),
        _text_diff_lines(old, new),
        _transform_diff_lines(old, new),
        _media_diff_lines(old, new),
    ]
    return "\n\n".join("\n".join(section) for section in sections)
