from __future__ import annotations

import argparse
import posixpath
import subprocess
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree

from ooxml_utils import read_all_stdin, write_text_stdout_utf8


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

EMU_PER_CM = 360000
IMAGE_EXTENSIONS = {".bmp", ".emf", ".gif", ".jpeg", ".jpg", ".png", ".svg", ".tif", ".tiff", ".wmf"}
VIDEO_EXTENSIONS = {".avi", ".m4v", ".mov", ".mp4", ".mpeg", ".mpg", ".wmv"}


@dataclass(frozen=True)
class Relationship:
    """OOXMLのRelationship情報を保持するためのデータクラス。"""

    rel_id: str
    rel_type: str
    target: str


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
class PackageSummary:
    """要約に必要なOOXMLパッケージ内の情報をまとめて保持するためのデータクラス。"""

    slides: dict[str, SlideInfo]
    slide_order: list[str]
    texts: list[TextItem]
    transforms: dict[tuple[str, str], TransformItem]
    media_paths: set[str]


def _local_name(tag: str) -> str:
    """名前空間付きXMLタグからローカル名だけを取り出す関数。"""
    if tag.startswith("{"):
        return tag.rsplit("}", 1)[1]
    return tag


def _parse_xml(data: bytes) -> ElementTree.Element | None:
    """XMLバイト列をElementTreeへ変換し、失敗時はNoneを返す関数。"""
    try:
        return ElementTree.fromstring(data)
    except ElementTree.ParseError:
        return None


def _read_zip_member(archive: zipfile.ZipFile, path: str) -> bytes | None:
    """zip内の指定パスを読み込み、存在しない場合はNoneを返す関数。"""
    try:
        return archive.read(path)
    except KeyError:
        return None


def _resolve_target(base_dir: str, target: str) -> str:
    """Relationshipの相対TargetをOOXMLパッケージ内パスへ解決する関数。"""
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(base_dir, target))


def _relationship_path(owner_path: str) -> str:
    """OOXML部品パスから対応する_rels内のRelationshipパスを作る関数。"""
    directory = posixpath.dirname(owner_path)
    filename = posixpath.basename(owner_path)
    return posixpath.join(directory, "_rels", f"{filename}.rels")


def _parse_relationships(archive: zipfile.ZipFile, owner_path: str) -> dict[str, Relationship]:
    """指定OOXML部品に紐づくRelationship一覧を読み込む関数。"""
    rels_path = _relationship_path(owner_path)
    data = _read_zip_member(archive, rels_path)
    if data is None:
        return {}

    root = _parse_xml(data)
    if root is None:
        return {}

    base_dir = posixpath.dirname(owner_path)
    relationships: dict[str, Relationship] = {}
    for element in root:
        if _local_name(element.tag) != "Relationship":
            continue
        rel_id = element.attrib.get("Id", "")
        target = element.attrib.get("Target", "")
        if not rel_id or not target or element.attrib.get("TargetMode") == "External":
            continue
        relationships[rel_id] = Relationship(
            rel_id=rel_id,
            rel_type=element.attrib.get("Type", ""),
            target=_resolve_target(base_dir, target),
        )
    return relationships


def _fallback_slides(archive: zipfile.ZipFile) -> list[SlideInfo]:
    """presentation.xmlが読めない場合にslide*.xmlの名前からスライド一覧を作る関数。"""
    paths = sorted(name for name in archive.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml"))
    return [SlideInfo(page=index + 1, slide_id=path, path=path) for index, path in enumerate(paths)]


def _parse_slides(archive: zipfile.ZipFile) -> list[SlideInfo]:
    """presentation.xmlとRelationshipからスライド順、ページ番号、スライドIDを読み取る関数。"""
    data = _read_zip_member(archive, "ppt/presentation.xml")
    if data is None:
        return _fallback_slides(archive)

    root = _parse_xml(data)
    if root is None:
        return _fallback_slides(archive)

    rels = _parse_relationships(archive, "ppt/presentation.xml")
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
    if len(preview) > 30:
        return preview[:29] + "..."
    return preview


def _object_label(element: ElementTree.Element, fallback_index: int) -> tuple[str, str]:
    """図形や画像を比較するためのキーと表示名を作る関数。"""
    c_nv_pr = element.find(".//p:cNvPr", NS)
    object_id = c_nv_pr.attrib.get("id", "") if c_nv_pr is not None else ""
    object_name = c_nv_pr.attrib.get("name", "") if c_nv_pr is not None else ""
    preview = _first_text_preview(element)
    object_type = _local_name(element.tag)
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
    candidates = [element for element in root.iter() if _local_name(element.tag) in {"sp", "pic", "graphicFrame", "cxnSp"}]
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


def _is_media_path(path: str) -> bool:
    """OOXMLパッケージ内パスが画像または動画らしいかを判定する関数。"""
    suffix = Path(path).suffix.lower()
    return path.startswith("ppt/media/") and suffix in IMAGE_EXTENSIONS | VIDEO_EXTENSIONS


def _collect_slide_media(archive: zipfile.ZipFile, slide: SlideInfo) -> set[str]:
    """スライドのRelationshipから画像や動画のパスを収集する関数。"""
    media_paths: set[str] = set()
    for relationship in _parse_relationships(archive, slide.path).values():
        rel_type = relationship.rel_type.lower()
        if "image" in rel_type or "video" in rel_type or "media" in rel_type or _is_media_path(relationship.target):
            media_paths.add(relationship.target)
    return media_paths


def _summarize_package(data: bytes) -> PackageSummary:
    """OOXML/PPTXバイト列から差分要約に必要な情報を抽出する関数。"""
    slides: dict[str, SlideInfo] = {}
    slide_order: list[str] = []
    texts: list[TextItem] = []
    transforms: dict[tuple[str, str], TransformItem] = {}
    media_paths: set[str] = set()

    with zipfile.ZipFile(BytesIO(data)) as archive:
        for name in archive.namelist():
            if _is_media_path(name):
                media_paths.add(name)

        for slide in _parse_slides(archive):
            slide_key = _slide_key(slide)
            slides[slide_key] = slide
            slide_order.append(slide_key)
            slide_xml = _read_zip_member(archive, slide.path)
            if slide_xml is None:
                continue
            root = _parse_xml(slide_xml)
            if root is None:
                continue
            texts.extend(_collect_texts(root, slide, slide_key))
            transforms.update(_collect_transforms(root, slide, slide_key))
            media_paths.update(_collect_slide_media(archive, slide))

    return PackageSummary(slides=slides, slide_order=slide_order, texts=texts, transforms=transforms, media_paths=media_paths)


def _format_slide(slide: SlideInfo) -> str:
    """スライド情報をページ番号とID付きの文字列へ変換する関数。"""
    return f"p.{slide.page} id={slide.slide_id} ({slide.path})"


def _format_text_item(item: TextItem) -> str:
    """テキスト差分の1項目を表示用文字列へ変換する関数。"""
    return f"p.{item.page} id={item.slide_id}: {item.text}"


def _counter_items(items: list[TextItem]) -> Counter[str]:
    """テキスト項目の出現回数を数える関数。"""
    return Counter(item.text for item in items)


def _index_text_items(items: list[TextItem]) -> dict[str, list[TextItem]]:
    """テキスト文字列から出現位置一覧へ引ける辞書を作る関数。"""
    indexed: dict[str, list[TextItem]] = defaultdict(list)
    for item in items:
        indexed[item.text].append(item)
    return indexed


def _text_diff_lines(old: PackageSummary, new: PackageSummary) -> list[str]:
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
        item = new_index[text].pop(0)
        lines.append(f"+ {_format_text_item(item)}")
    for text in deletions:
        item = old_index[text].pop(0)
        lines.append(f"- {_format_text_item(item)}")
    return lines


def _slide_diff_lines(old: PackageSummary, new: PackageSummary) -> list[str]:
    """スライドの追加と削除を要約する行リストを作る関数。"""
    lines = ["## スライド差分"]
    old_keys = set(old.slides)
    new_keys = set(new.slides)
    added = [new.slides[key] for key in new.slide_order if key in new_keys - old_keys]
    deleted = [old.slides[key] for key in old.slide_order if key in old_keys - new_keys]

    if not added and not deleted:
        return lines + ["変更なし"]

    for slide in added:
        lines.append(f"+ {_format_slide(slide)}")
    for slide in deleted:
        lines.append(f"- {_format_slide(slide)}")
    return lines


def _cm(value: int) -> str:
    """EMU単位の値をcm表記へ変換する関数。"""
    return f"{value / EMU_PER_CM:.2f}cm"


def _ratio(old_value: int, new_value: int) -> str:
    """旧値と新値から拡大率の倍率表記を作る関数。"""
    if old_value == 0:
        return "n/a"
    return f"{new_value / old_value:.3f}x"


def _transform_diff_lines(old: PackageSummary, new: PackageSummary) -> list[str]:
    """位置と拡大率の差分を要約する行リストを作る関数。"""
    lines = ["## 位置・拡大率差分"]
    common_keys = sorted(set(old.transforms) & set(new.transforms))
    changed: list[str] = []

    for key in common_keys:
        old_item = old.transforms[key]
        new_item = new.transforms[key]
        if (old_item.x, old_item.y, old_item.cx, old_item.cy) == (new_item.x, new_item.y, new_item.cx, new_item.cy):
            continue
        changed.append(
            f"* p.{new_item.page} id={new_item.slide_id} {new_item.label}: "
            f"位置 ({_cm(old_item.x)}, {_cm(old_item.y)}) -> ({_cm(new_item.x)}, {_cm(new_item.y)}), "
            f"サイズ ({_cm(old_item.cx)}, {_cm(old_item.cy)}) -> ({_cm(new_item.cx)}, {_cm(new_item.cy)}), "
            f"拡大率 ({_ratio(old_item.cx, new_item.cx)}, {_ratio(old_item.cy, new_item.cy)})"
        )

    return lines + (changed if changed else ["変更なし"])


def _media_diff_lines(old: PackageSummary, new: PackageSummary) -> list[str]:
    """画像や動画の追加と削除を要約する行リストを作る関数。"""
    lines = ["## 画像・動画差分"]
    added = sorted(new.media_paths - old.media_paths)
    deleted = sorted(old.media_paths - new.media_paths)

    if not added and not deleted:
        return lines + ["変更なし"]

    lines.extend(f"+ {path}" for path in added)
    lines.extend(f"- {path}" for path in deleted)
    return lines


def summarize_ooxml_diff(old_data: bytes, new_data: bytes) -> str:
    """2つのOOXML/PPTXバイト列を比較し、人が読みやすい差分要約を返す関数。"""
    old = _summarize_package(old_data)
    new = _summarize_package(new_data)
    sections = [
        _slide_diff_lines(old, new),
        _text_diff_lines(old, new),
        _transform_diff_lines(old, new),
        _media_diff_lines(old, new),
    ]
    return "\n\n".join("\n".join(section) for section in sections)


def summarize_ooxml_files(old_path: Path, new_path: Path) -> str:
    """2つのOOXML/PPTXファイルを読み込み、人が読みやすい差分要約を返す関数。"""
    return summarize_ooxml_diff(old_path.read_bytes(), new_path.read_bytes())


def _git_relative_path(path: Path) -> str:
    """作業ツリー上のパスをGitで扱いやすいリポジトリ相対パスへ変換する関数。"""
    if not path.is_absolute():
        return path.as_posix()

    root_data = subprocess.check_output(["git", "rev-parse", "--show-toplevel"])
    root = Path(root_data.decode("utf-8", errors="replace").strip()).resolve()
    try:
        return path.resolve().relative_to(root).as_posix()
    except ValueError:
        return path.as_posix()


def summarize_git_worktree_file(path: Path, ref: str = "HEAD") -> str:
    """Gitの指定ref上のファイルと作業ツリー上のファイルを比較し、差分要約を返す関数。"""
    git_path = _git_relative_path(path)
    old_data = subprocess.check_output(["git", "show", f"{ref}:{git_path}"])
    new_data = path.read_bytes()
    return summarize_ooxml_diff(old_data, new_data)


def build_parser() -> argparse.ArgumentParser:
    """コマンドライン引数を解釈するパーサーを作る関数。"""
    parser = argparse.ArgumentParser(description="Summarize OOXML/PPTX differences for humans.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    files_parser = subparsers.add_parser("files", help="summarize two OOXML/PPTX files")
    files_parser.add_argument("old")
    files_parser.add_argument("new")

    git_parser = subparsers.add_parser("git", help="summarize ref vs worktree for one file")
    git_parser.add_argument("path")
    git_parser.add_argument("--ref", default="HEAD")

    stdin_parser = subparsers.add_parser("stdin", help="summarize two files, reading the new file from stdin")
    stdin_parser.add_argument("old")
    return parser


def main(argv: list[str] | None = None) -> int:
    """サブコマンドに応じてOOXML差分要約を実行する関数。"""
    args = build_parser().parse_args(argv)
    if args.command == "files":
        summary = summarize_ooxml_files(Path(args.old), Path(args.new))
    elif args.command == "git":
        summary = summarize_git_worktree_file(Path(args.path), args.ref)
    elif args.command == "stdin":
        summary = summarize_ooxml_diff(Path(args.old).read_bytes(), read_all_stdin())
    else:
        raise ValueError(f"unknown command: {args.command}")

    write_text_stdout_utf8(summary + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
