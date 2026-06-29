from __future__ import annotations

import posixpath
import zipfile
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}

EMU_PER_CM = 360000
IMAGE_EXTENSIONS = {".bmp", ".emf", ".gif", ".jpeg", ".jpg", ".png", ".svg", ".tif", ".tiff", ".wmf"}
VIDEO_EXTENSIONS = {".avi", ".m4v", ".mov", ".mp4", ".mpeg", ".mpg", ".wmv"}
MEDIA_EXTENSIONS = IMAGE_EXTENSIONS | VIDEO_EXTENSIONS


@dataclass(frozen=True)
class Relationship:
    """OOXMLのRelationship情報を保持するためのデータクラス。"""

    rel_id: str
    rel_type: str
    target: str


def local_name(tag: str) -> str:
    """名前空間付きXMLタグからローカル名だけを取り出す関数。"""
    if tag.startswith("{"):
        return tag.rsplit("}", 1)[1]
    return tag


def parse_xml(data: bytes) -> ElementTree.Element | None:
    """XMLバイト列をElementTreeへ変換し、失敗時はNoneを返す関数。"""
    try:
        return ElementTree.fromstring(data)
    except ElementTree.ParseError:
        return None


def open_zip(data: bytes) -> zipfile.ZipFile:
    """OOXMLバイト列をzipとして読み込む関数。"""
    return zipfile.ZipFile(BytesIO(data))


def is_zip_data(data: bytes) -> bool:
    """バイト列がzipファイルとして読み込めるかを判定する関数。"""
    try:
        with open_zip(data) as archive:
            return archive.testzip() is None
    except zipfile.BadZipFile:
        return False


def read_zip_member(archive: zipfile.ZipFile, path: str) -> bytes | None:
    """zip内の指定パスを読み込み、存在しない場合はNoneを返す関数。"""
    try:
        return archive.read(path)
    except KeyError:
        return None


def resolve_target(base_dir: str, target: str) -> str:
    """Relationshipの相対TargetをOOXMLパッケージ内パスへ解決する関数。"""
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(base_dir, target))


def relationship_path(owner_path: str) -> str:
    """OOXML部品パスから対応する_rels内のRelationshipパスを作る関数。"""
    directory = posixpath.dirname(owner_path)
    filename = posixpath.basename(owner_path)
    return posixpath.join(directory, "_rels", f"{filename}.rels")


def parse_relationships(archive: zipfile.ZipFile, owner_path: str) -> dict[str, Relationship]:
    """指定OOXML部品に紐づくRelationship一覧を読み込む関数。"""
    data = read_zip_member(archive, relationship_path(owner_path))
    if data is None:
        return {}

    root = parse_xml(data)
    if root is None:
        return {}

    base_dir = posixpath.dirname(owner_path)
    relationships: dict[str, Relationship] = {}
    for element in root:
        if local_name(element.tag) != "Relationship":
            continue
        rel_id = element.attrib.get("Id", "")
        target = element.attrib.get("Target", "")
        if not rel_id or not target or element.attrib.get("TargetMode") == "External":
            continue
        relationships[rel_id] = Relationship(
            rel_id=rel_id,
            rel_type=element.attrib.get("Type", ""),
            target=resolve_target(base_dir, target),
        )
    return relationships


def is_media_path(path: str) -> bool:
    """OOXMLパッケージ内パスが画像または動画らしいかを判定する関数。"""
    return Path(path).suffix.lower() in MEDIA_EXTENSIONS


def collect_package_media_paths(archive: zipfile.ZipFile, prefix: str) -> set[str]:
    """指定prefix配下の画像や動画パスをzip全体から収集する関数。"""
    return {name for name in archive.namelist() if name.startswith(prefix) and is_media_path(name)}


def relationship_media_paths(archive: zipfile.ZipFile, owner_path: str) -> set[str]:
    """指定OOXML部品のRelationshipから画像や動画パスを収集する関数。"""
    media_paths: set[str] = set()
    for relationship in parse_relationships(archive, owner_path).values():
        rel_type = relationship.rel_type.lower()
        if "image" in rel_type or "video" in rel_type or "media" in rel_type or is_media_path(relationship.target):
            media_paths.add(relationship.target)
    return media_paths


def cm(value: int) -> str:
    """EMU単位の値をcm表記へ変換する関数。"""
    return f"{value / EMU_PER_CM:.2f}cm"


def ratio(old_value: int, new_value: int) -> str:
    """旧値と新値から拡大率の倍率表記を作る関数。"""
    if old_value == 0:
        return "n/a"
    return f"{new_value / old_value:.3f}x"


def cell_reference(column_index: int, row_index: int) -> str:
    """0始まりの列番号と行番号をExcelのセル参照へ変換する関数。"""
    column = ""
    index = column_index
    while True:
        index, remainder = divmod(index, 26)
        column = chr(ord("A") + remainder) + column
        if index == 0:
            break
        index -= 1
    return f"{column}{row_index + 1}"


def generic_zip_textconv(data: bytes) -> str:
    """未知のOOXML zipを従来どおりXML中心のテキスト表現へ変換する関数。"""
    if not is_zip_data(data):
        return data.decode("utf-8", errors="replace")

    lines: list[str] = []
    with open_zip(data) as archive:
        for name in sorted(item for item in archive.namelist() if not item.endswith("/")):
            payload = archive.read(name)
            lines.append(f"--- {name}")
            if name.lower().endswith((".xml", ".rels")):
                lines.append(payload.decode("utf-8", errors="replace"))
            else:
                lines.append(f"<binary size={len(payload)}>")
            lines.append("")
    return "\n".join(lines)
