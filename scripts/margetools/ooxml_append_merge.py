from __future__ import annotations

import argparse
import posixpath
import subprocess
import sys
import zipfile
from copy import deepcopy
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET


SCRIPT_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = SCRIPT_DIR.parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from ooxml_storage import container_to_zip_data
from ooxml_utils import write_zip_entry


NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

REL_SLIDE = f"{NS_OFFICE_REL}/slide"
REL_WORKSHEET = f"{NS_OFFICE_REL}/worksheet"
MERGE_WARNING = "OOXML自動マージ: incoming側のXML差分を末尾へ追加しました。内容を手動確認してください。"
REPORT_DIR = "mergereport"

OOXML_EXTENSIONS = {
    ".pptx",
    ".pptm",
    ".docx",
    ".docm",
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm",
    ".xlmx",
}


# この merge driver は「incomingを末尾に追加する」ことに特化している。
# 補えない例:
# - 同じスライド・段落・セルを両者が編集した場合、意味的な統合や優先順位判断はできない。
# - PowerPointのスライドマスター、Wordのスタイル/番号定義、Excelのスタイル/数式参照の完全統合はできない。
# - マクロ、外部リンク、コメント、変更履歴、署名、埋め込みOLEなど、Office固有の複雑な関連部品は手動確認が必要になる。


@dataclass
class MergeReportEntry:
    """外部レポートに残す、自動マージで追加した部品の情報を表すクラス。"""

    kind: str
    added_part: str
    base_part: str | None
    current_part: str | None
    incoming_part: str
    note: str


def register_namespaces() -> None:
    """ElementTreeでXMLを書き戻す際に、OOXMLで一般的な接頭辞を保ちやすくする関数。"""
    namespaces = {
        "p": NS_P,
        "r": NS_R,
        "w": NS_W,
        "s": NS_SS,
        "rel": NS_REL,
        "ct": NS_CT,
    }
    for prefix, namespace in namespaces.items():
        ET.register_namespace(prefix, namespace)


def q(namespace: str, name: str) -> str:
    """ElementTreeで名前空間付きタグや属性名を作る関数。"""
    return f"{{{namespace}}}{name}"


def log(message: str) -> None:
    """日本語パスを含むログでも文字化けやエンコード例外を避けるためUTF-8で出力する関数。"""
    sys.stderr.buffer.write((message + "\n").encode("utf-8", errors="replace"))


def read_package(path: Path) -> dict[str, bytes]:
    """textconvと同じ正規化入口を通して、zip内メンバー辞書へ変換する関数。"""
    data = container_to_zip_data(path.read_bytes())
    members: dict[str, bytes] = {}
    with zipfile.ZipFile(BytesIO(data)) as archive:
        for name in archive.namelist():
            if not name.endswith("/"):
                members[name.replace("\\", "/")] = archive.read(name)
    return members


def read_package_or_empty(path: Path) -> dict[str, bytes]:
    """baseなどが空またはOOXMLでない場合に、空パッケージとして扱う関数。"""
    try:
        if not path.exists() or path.stat().st_size == 0:
            return {}
        return read_package(path)
    except (OSError, ValueError, zipfile.BadZipFile):
        return {}


def write_package(path: Path, members: dict[str, bytes]) -> None:
    """OOXMLメンバー辞書を固定順序のzipとしてファイルへ書き戻す関数。"""
    output = BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for name in sorted(members):
            write_zip_entry(archive, name, members[name])
    path.write_bytes(output.getvalue())


def parse_xml(members: dict[str, bytes], name: str) -> ET.Element:
    """OOXMLメンバー内のXMLをElementTreeとして読み込む関数。"""
    return ET.fromstring(members[name])


def xml_bytes(root: ET.Element) -> bytes:
    """ElementTreeをUTF-8 XML宣言付きのバイト列へ変換する関数。"""
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def comparison_note(base_part: str | None, current_part: str | None, incoming_part: str) -> str:
    """外部レポートへ書き込む比較対象メモを作る関数。"""
    return f"base={base_part or '(none)'} current={current_part or '(none)'} incoming={incoming_part}"


def markdown_cell(text: str | None) -> str:
    """Markdown表のセルに入れる文字列を安全な表記へ変換する関数。"""
    if not text:
        return "(none)"
    return text.replace("|", "\\|").replace("\n", "<br>")


def report_filename(worktree_path: str | None, kind: str) -> str:
    """マージ対象パスと時刻からレポートファイル名を作る関数。"""
    raw_name = Path(worktree_path or f"ooxml-{kind}").name
    safe_name = "".join(ch if ch.isalnum() or ch in {".", "-", "_"} else "_" for ch in raw_name)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return f"{safe_name}-{timestamp}.md"


def write_merge_report(entries: list[MergeReportEntry], worktree_path: str | None, kind: str) -> Path | None:
    """自動マージで追加した内容と警告をmergereport配下のMarkdownへ保存する関数。"""
    if not entries:
        return None

    report_dir = Path(REPORT_DIR)
    report_dir.mkdir(parents=True, exist_ok=True)
    report_path = report_dir / report_filename(worktree_path, kind)
    lines = [
        "# OOXML Merge Report",
        "",
        f"- target: `{worktree_path or '(unknown)'}`",
        f"- kind: `{kind}`",
        f"- warning: {MERGE_WARNING}",
        "",
        "## Added Parts",
        "",
        "| kind | added | base | current | incoming | note |",
        "|---|---|---|---|---|---|",
    ]
    for entry in entries:
        lines.append(
            "| "
            + " | ".join(
                [
                    markdown_cell(entry.kind),
                    markdown_cell(entry.added_part),
                    markdown_cell(entry.base_part),
                    markdown_cell(entry.current_part),
                    markdown_cell(entry.incoming_part),
                    markdown_cell(entry.note),
                ]
            )
            + " |"
        )
    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return report_path


def xml_member_changed(
    base: dict[str, bytes],
    current: dict[str, bytes],
    incoming: dict[str, bytes],
    base_part: str | None,
    current_part: str | None,
    incoming_part: str,
) -> bool:
    """textconv相当の正規化後メンバーを.xml単位で比較し、incoming差分の有無を判定する関数。"""
    incoming_data = incoming.get(incoming_part)
    if incoming_data is None:
        return False
    if base_part and base.get(base_part) == incoming_data:
        return False
    if current_part and current.get(current_part) == incoming_data:
        return False
    return True


def rels_path_for_part(part_name: str) -> str:
    """OOXML部品名から対応する.rels部品名を返す関数。"""
    directory, filename = posixpath.split(part_name)
    if directory:
        return f"{directory}/_rels/{filename}.rels"
    return f"_rels/{filename}.rels"


def part_from_relationship_target(source_part: str, target: str) -> str:
    """relationshipのTargetを、zip内の正規化された部品パスへ変換する関数。"""
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def target_from_part(source_part: str, target_part: str) -> str:
    """zip内の部品パスを、relationshipの相対Targetへ変換する関数。"""
    return posixpath.relpath(target_part, posixpath.dirname(source_part))


def next_numeric_suffix(existing: set[int], start: int = 1) -> int:
    """使用済み番号集合を避けた次の番号を返す関数。"""
    value = start
    while value in existing:
        value += 1
    return value


def numbers_from_paths(paths: list[str], prefix: str, suffix: str) -> set[int]:
    """sheet1.xmlやslide2.xmlのようなパスから数値部分を集める関数。"""
    numbers: set[int] = set()
    for path in paths:
        filename = posixpath.basename(path)
        if filename.startswith(prefix) and filename.endswith(suffix):
            raw = filename[len(prefix) : -len(suffix)]
            if raw.isdigit():
                numbers.add(int(raw))
    return numbers


def unique_part_name(members: dict[str, bytes], desired: str) -> str:
    """既存部品と衝突しないzip内パスを作る関数。"""
    if desired not in members:
        return desired
    directory, filename = posixpath.split(desired)
    stem, suffix = posixpath.splitext(filename)
    index = 1
    while True:
        candidate = posixpath.join(directory, f"{stem}_incoming{index}{suffix}")
        if candidate not in members:
            return candidate
        index += 1


def next_rid(rels_root: ET.Element) -> str:
    """relationship XMLで未使用のrIdを返す関数。"""
    used: set[int] = set()
    for rel in rels_root.findall(q(NS_REL, "Relationship")):
        rel_id = rel.get("Id", "")
        if rel_id.startswith("rId") and rel_id[3:].isdigit():
            used.add(int(rel_id[3:]))
    return f"rId{next_numeric_suffix(used)}"


def add_relationship(
    rels_root: ET.Element,
    rel_type: str,
    target: str,
    target_mode: str | None = None,
) -> str:
    """relationship XMLへ新しいRelationship要素を追加し、そのIdを返す関数。"""
    rel_id = next_rid(rels_root)
    attrs = {"Id": rel_id, "Type": rel_type, "Target": target}
    if target_mode:
        attrs["TargetMode"] = target_mode
    ET.SubElement(rels_root, q(NS_REL, "Relationship"), attrs)
    return rel_id


def merge_content_types_for_part(
    source: dict[str, bytes],
    target: dict[str, bytes],
    source_part: str,
    target_part: str,
) -> None:
    """コピーした部品に必要なContentType定義をtarget側へ追加する関数。"""
    if "[Content_Types].xml" not in source or "[Content_Types].xml" not in target:
        return

    source_root = parse_xml(source, "[Content_Types].xml")
    target_root = parse_xml(target, "[Content_Types].xml")
    existing_overrides = {
        override.get("PartName")
        for override in target_root.findall(q(NS_CT, "Override"))
    }
    existing_defaults = {
        default.get("Extension")
        for default in target_root.findall(q(NS_CT, "Default"))
    }

    for default in source_root.findall(q(NS_CT, "Default")):
        extension = default.get("Extension")
        if extension and extension not in existing_defaults:
            target_root.append(deepcopy(default))
            existing_defaults.add(extension)

    source_part_name = "/" + source_part
    target_part_name = "/" + target_part
    for override in source_root.findall(q(NS_CT, "Override")):
        if override.get("PartName") == source_part_name and target_part_name not in existing_overrides:
            copied = deepcopy(override)
            copied.set("PartName", target_part_name)
            target_root.append(copied)
            existing_overrides.add(target_part_name)
            break

    target["[Content_Types].xml"] = xml_bytes(target_root)


def copy_related_parts(
    source: dict[str, bytes],
    target: dict[str, bytes],
    source_part: str,
    target_part: str,
    mapping: dict[str, str],
) -> None:
    """コピー済み部品のrelationshipと参照先部品を再帰的にコピーする関数。"""
    source_rels = rels_path_for_part(source_part)
    if source_rels not in source:
        return

    rels_root = parse_xml(source, source_rels)
    for rel in rels_root.findall(q(NS_REL, "Relationship")):
        if rel.get("TargetMode") == "External":
            continue

        old_target = part_from_relationship_target(source_part, rel.get("Target", ""))
        if old_target not in source:
            continue

        if old_target not in mapping:
            mapping[old_target] = unique_part_name(target, old_target)
            copy_part_recursive(source, target, old_target, mapping[old_target], mapping)

        rel.set("Target", target_from_part(target_part, mapping[old_target]))

    target[rels_path_for_part(target_part)] = xml_bytes(rels_root)


def copy_part_recursive(
    source: dict[str, bytes],
    target: dict[str, bytes],
    source_part: str,
    target_part: str,
    mapping: dict[str, str],
) -> None:
    """部品本体・ContentType・関連部品をtarget側へ再帰的にコピーする関数。"""
    target[target_part] = source[source_part]
    merge_content_types_for_part(source, target, source_part, target_part)
    copy_related_parts(source, target, source_part, target_part, mapping)


def relationship_targets_by_type(members: dict[str, bytes], part_name: str, rel_type: str) -> list[str]:
    """指定部品のrelationshipから、指定Typeの参照先部品を取得する関数。"""
    rels_name = rels_path_for_part(part_name)
    if rels_name not in members:
        return []
    rels_root = parse_xml(members, rels_name)
    targets: list[str] = []
    for rel in rels_root.findall(q(NS_REL, "Relationship")):
        if rel.get("Type") == rel_type and rel.get("TargetMode") != "External":
            targets.append(part_from_relationship_target(part_name, rel.get("Target", "")))
    return targets


def slide_layout_for_slide(members: dict[str, bytes], slide_part: str | None) -> str | None:
    """PowerPointスライドが参照しているslideLayout部品を返す関数。"""
    if not slide_part:
        return None
    targets = relationship_targets_by_type(members, slide_part, f"{NS_OFFICE_REL}/slideLayout")
    return targets[0] if targets else None


def presentation_master_parts(members: dict[str, bytes]) -> list[str]:
    """PowerPointのpresentation.xmlからslideMaster部品を取得する関数。"""
    if "ppt/_rels/presentation.xml.rels" not in members:
        return []
    return relationship_targets_by_type(members, "ppt/presentation.xml", f"{NS_OFFICE_REL}/slideMaster")


def next_numbered_part_name(members: dict[str, bytes], directory: str, prefix: str, suffix: str) -> str:
    """slideLayout12.xmlのような連番部品名の次の未使用パスを作る関数。"""
    used = numbers_from_paths([name for name in members if name.startswith(directory + "/")], prefix, suffix)
    number = next_numeric_suffix(used)
    return f"{directory}/{prefix}{number}{suffix}"


def add_layout_to_master(members: dict[str, bytes], master_part: str, layout_part: str) -> None:
    """slideMaster.xmlとそのrelsへ、新しいslideLayout参照を登録する関数。"""
    master_root = parse_xml(members, master_part)
    master_rels_path = rels_path_for_part(master_part)
    master_rels = parse_xml(members, master_rels_path)
    rel_id = add_relationship(master_rels, f"{NS_OFFICE_REL}/slideLayout", target_from_part(master_part, layout_part))

    layout_list = master_root.find(q(NS_P, "sldLayoutIdLst"))
    if layout_list is None:
        layout_list = ET.SubElement(master_root, q(NS_P, "sldLayoutIdLst"))
    used_ids = {
        int(item.get("id", "0"))
        for item in layout_list.findall(q(NS_P, "sldLayoutId"))
        if item.get("id", "0").isdigit()
    }
    layout_id = max(used_ids | {2147483648}) + 1
    ET.SubElement(layout_list, q(NS_P, "sldLayoutId"), {"id": str(layout_id), q(NS_R, "id"): rel_id})

    members[master_part] = xml_bytes(master_root)
    members[master_rels_path] = xml_bytes(master_rels)


def copy_pptx_layout_as_standard_part(
    source: dict[str, bytes],
    target: dict[str, bytes],
    source_layout: str,
) -> str | None:
    """incomingのslideLayoutを_incoming名ではなく標準連番名でコピーする関数。"""
    if source_layout not in source:
        return None

    target_layout = next_numbered_part_name(target, "ppt/slideLayouts", "slideLayout", ".xml")
    target[target_layout] = source[source_layout]
    merge_content_types_for_part(source, target, source_layout, target_layout)

    target_masters = presentation_master_parts(target)
    target_master = target_masters[0] if target_masters else None
    source_rels_path = rels_path_for_part(source_layout)
    if source_rels_path in source:
        rels_root = parse_xml(source, source_rels_path)
        for rel in rels_root.findall(q(NS_REL, "Relationship")):
            if rel.get("TargetMode") == "External":
                continue
            if rel.get("Type") == f"{NS_OFFICE_REL}/slideMaster" and target_master:
                rel.set("Target", target_from_part(target_layout, target_master))
        target[rels_path_for_part(target_layout)] = xml_bytes(rels_root)
    elif target_master:
        rels_root = ET.Element(q(NS_REL, "Relationships"))
        add_relationship(rels_root, f"{NS_OFFICE_REL}/slideMaster", target_from_part(target_layout, target_master))
        target[rels_path_for_part(target_layout)] = xml_bytes(rels_root)

    if target_master:
        add_layout_to_master(target, target_master, target_layout)
    return target_layout


def choose_pptx_layout(
    source: dict[str, bytes],
    target: dict[str, bytes],
    source_slide: str,
    current_slide: str | None,
) -> tuple[str | None, str]:
    """追加スライドに使う安全なslideLayoutを選ぶ関数。"""
    current_layout = slide_layout_for_slide(target, current_slide)
    if current_layout and current_layout in target:
        return current_layout, "current counterpart slide layout reused"

    source_layout = slide_layout_for_slide(source, source_slide)
    if source_layout and source_layout in target:
        return source_layout, "same named current layout reused"
    if source_layout:
        copied_layout = copy_pptx_layout_as_standard_part(source, target, source_layout)
        if copied_layout:
            return copied_layout, "incoming layout copied as standard numbered layout"
    return None, "slide layout relationship omitted because no safe layout was found"


def copy_pptx_slide_without_structural_duplicates(
    source: dict[str, bytes],
    target: dict[str, bytes],
    source_slide: str,
    target_slide: str,
    layout_part: str | None,
) -> None:
    """slideLayout/slideMaster/themeを_incoming複製せず、スライド本体と直接関連部品をコピーする関数。"""
    target[target_slide] = source[source_slide]
    merge_content_types_for_part(source, target, source_slide, target_slide)

    source_rels = rels_path_for_part(source_slide)
    if source_rels not in source:
        return

    rels_root = parse_xml(source, source_rels)
    for rel in list(rels_root.findall(q(NS_REL, "Relationship"))):
        if rel.get("TargetMode") == "External":
            continue
        rel_type = rel.get("Type", "")
        old_target = part_from_relationship_target(source_slide, rel.get("Target", ""))
        if rel_type == f"{NS_OFFICE_REL}/slideLayout":
            if layout_part:
                rel.set("Target", target_from_part(target_slide, layout_part))
            else:
                rels_root.remove(rel)
            continue
        if old_target not in source:
            continue
        new_target = unique_part_name(target, old_target)
        copy_part_recursive(source, target, old_target, new_target, {old_target: new_target})
        rel.set("Target", target_from_part(target_slide, new_target))

    target[rels_path_for_part(target_slide)] = xml_bytes(rels_root)


def presentation_slide_parts(members: dict[str, bytes]) -> list[str]:
    """PowerPointのpresentation.xmlからスライド部品を表示順に取得する関数。"""
    presentation = parse_xml(members, "ppt/presentation.xml")
    rels = parse_xml(members, "ppt/_rels/presentation.xml.rels")
    rel_targets = {
        rel.get("Id"): part_from_relationship_target("ppt/presentation.xml", rel.get("Target", ""))
        for rel in rels.findall(q(NS_REL, "Relationship"))
        if rel.get("Type") == REL_SLIDE
    }

    slide_parts: list[str] = []
    slide_id_list = presentation.find(q(NS_P, "sldIdLst"))
    if slide_id_list is None:
        return slide_parts
    for slide_id in slide_id_list.findall(q(NS_P, "sldId")):
        rel_id = slide_id.get(q(NS_R, "id"))
        if rel_id in rel_targets:
            slide_parts.append(rel_targets[rel_id])
    return slide_parts


def append_pptx_incoming(
    base: dict[str, bytes],
    current: dict[str, bytes],
    incoming: dict[str, bytes],
    report_entries: list[MergeReportEntry],
) -> None:
    """incoming側で差分のあるPowerPointスライドだけをcurrent側の末尾へ追加する関数。"""
    presentation = parse_xml(current, "ppt/presentation.xml")
    presentation_rels = parse_xml(current, "ppt/_rels/presentation.xml.rels")
    slide_id_list = presentation.find(q(NS_P, "sldIdLst"))
    if slide_id_list is None:
        slide_id_list = ET.SubElement(presentation, q(NS_P, "sldIdLst"))

    used_slide_numbers = numbers_from_paths(list(current), "slide", ".xml")
    used_slide_ids = {
        int(slide.get("id", "0"))
        for slide in slide_id_list.findall(q(NS_P, "sldId"))
        if slide.get("id", "0").isdigit()
    }
    next_slide_id = max(used_slide_ids | {255}) + 1
    base_slides = presentation_slide_parts(base) if base else []
    current_slides = presentation_slide_parts(current)
    incoming_slides = presentation_slide_parts(incoming)
    appended = False

    for index, source_slide in enumerate(incoming_slides):
        base_slide = base_slides[index] if index < len(base_slides) else None
        current_slide = current_slides[index] if index < len(current_slides) else None
        if not xml_member_changed(base, current, incoming, base_slide, current_slide, source_slide):
            continue

        slide_number = next_numeric_suffix(used_slide_numbers)
        used_slide_numbers.add(slide_number)
        target_slide = f"ppt/slides/slide{slide_number}.xml"
        layout_part, layout_note = choose_pptx_layout(incoming, current, source_slide, current_slide)
        copy_pptx_slide_without_structural_duplicates(incoming, current, source_slide, target_slide, layout_part)
        rel_id = add_relationship(presentation_rels, REL_SLIDE, f"slides/slide{slide_number}.xml")
        ET.SubElement(slide_id_list, q(NS_P, "sldId"), {"id": str(next_slide_id), q(NS_R, "id"): rel_id})
        report_entries.append(
            MergeReportEntry(
                kind="pptx-slide",
                added_part=target_slide,
                base_part=base_slide,
                current_part=current_slide,
                incoming_part=source_slide,
                note=f"{comparison_note(base_slide, current_slide, source_slide)}; layout={layout_part or '(none)'}; {layout_note}",
            )
        )
        next_slide_id += 1
        appended = True

    if not appended:
        return
    current["ppt/presentation.xml"] = xml_bytes(presentation)
    current["ppt/_rels/presentation.xml.rels"] = xml_bytes(presentation_rels)


def document_relationship_map(current: dict[str, bytes], incoming: dict[str, bytes]) -> dict[str, str]:
    """Word本文が参照するincoming側relationshipをcurrent側へコピーし、rId対応表を返す関数。"""
    incoming_rels_path = "word/_rels/document.xml.rels"
    current_rels_path = "word/_rels/document.xml.rels"
    if incoming_rels_path not in incoming:
        return {}
    if current_rels_path in current:
        current_rels = parse_xml(current, current_rels_path)
    else:
        current_rels = ET.Element(q(NS_REL, "Relationships"))

    rid_map: dict[str, str] = {}
    incoming_rels = parse_xml(incoming, incoming_rels_path)
    for rel in incoming_rels.findall(q(NS_REL, "Relationship")):
        old_rid = rel.get("Id")
        if not old_rid:
            continue
        if rel.get("TargetMode") == "External":
            new_rid = add_relationship(current_rels, rel.get("Type", ""), rel.get("Target", ""), "External")
            rid_map[old_rid] = new_rid
            continue

        old_target = part_from_relationship_target("word/document.xml", rel.get("Target", ""))
        if old_target not in incoming:
            continue
        new_target = unique_part_name(current, old_target)
        copy_part_recursive(incoming, current, old_target, new_target, {old_target: new_target})
        new_rid = add_relationship(current_rels, rel.get("Type", ""), target_from_part("word/document.xml", new_target))
        rid_map[old_rid] = new_rid

    current[current_rels_path] = xml_bytes(current_rels)
    return rid_map


def rewrite_relationship_ids(element: ET.Element, rid_map: dict[str, str]) -> None:
    """コピーしたXML要素内のrId参照をcurrent側のrIdへ置き換える関数。"""
    for node in element.iter():
        for attr_name, attr_value in list(node.attrib.items()):
            if attr_value in rid_map and attr_name.startswith(f"{{{NS_R}}}"):
                node.set(attr_name, rid_map[attr_value])


def append_docx_incoming(
    base: dict[str, bytes],
    current: dict[str, bytes],
    incoming: dict[str, bytes],
    report_entries: list[MergeReportEntry],
) -> None:
    """incoming側で差分のあるWord本文だけを改ページ付きでcurrent側の末尾へ追加する関数。"""
    if not xml_member_changed(base, current, incoming, "word/document.xml", "word/document.xml", "word/document.xml"):
        return

    current_doc = parse_xml(current, "word/document.xml")
    incoming_doc = parse_xml(incoming, "word/document.xml")
    current_body = current_doc.find(q(NS_W, "body"))
    incoming_body = incoming_doc.find(q(NS_W, "body"))
    if current_body is None or incoming_body is None:
        raise ValueError("word/document.xml body is missing")

    rid_map = document_relationship_map(current, incoming)
    current_sect = current_body.find(q(NS_W, "sectPr"))
    if current_sect is not None:
        current_body.remove(current_sect)

    page_break_paragraph = ET.Element(q(NS_W, "p"))
    run = ET.SubElement(page_break_paragraph, q(NS_W, "r"))
    ET.SubElement(run, q(NS_W, "br"), {q(NS_W, "type"): "page"})
    current_body.append(page_break_paragraph)

    for child in list(incoming_body):
        if child.tag == q(NS_W, "sectPr"):
            continue
        copied = deepcopy(child)
        rewrite_relationship_ids(copied, rid_map)
        current_body.append(copied)

    if current_sect is not None:
        current_body.append(current_sect)
    current["word/document.xml"] = xml_bytes(current_doc)
    report_entries.append(
        MergeReportEntry(
            kind="docx-body",
            added_part="word/document.xml",
            base_part="word/document.xml",
            current_part="word/document.xml",
            incoming_part="word/document.xml",
            note=comparison_note("word/document.xml", "word/document.xml", "word/document.xml"),
        )
    )


def shared_string_count(root: ET.Element | None) -> int:
    """Excel sharedStrings.xml内の文字列数を返す関数。"""
    if root is None:
        return 0
    return len(root.findall(q(NS_SS, "si")))


def read_optional_xml(members: dict[str, bytes], name: str) -> ET.Element | None:
    """存在すればXMLを読み込み、なければNoneを返す関数。"""
    if name not in members:
        return None
    return parse_xml(members, name)


def ensure_shared_strings(current: dict[str, bytes], incoming: dict[str, bytes]) -> int:
    """incoming側のExcel共有文字列をcurrent側へ追記し、元の文字列数を返す関数。"""
    incoming_root = read_optional_xml(incoming, "xl/sharedStrings.xml")
    if incoming_root is None:
        return 0

    current_root = read_optional_xml(current, "xl/sharedStrings.xml")
    if current_root is None:
        current_root = ET.Element(q(NS_SS, "sst"))
    offset = shared_string_count(current_root)
    for item in incoming_root.findall(q(NS_SS, "si")):
        current_root.append(deepcopy(item))

    current_root.set("count", str(shared_string_count(current_root)))
    current_root.set("uniqueCount", str(shared_string_count(current_root)))
    current["xl/sharedStrings.xml"] = xml_bytes(current_root)
    merge_content_types_for_part(incoming, current, "xl/sharedStrings.xml", "xl/sharedStrings.xml")
    return offset


def remap_sheet_shared_strings(sheet_root: ET.Element, offset: int) -> None:
    """コピーするExcelシート内の共有文字列インデックスを追記後の位置へずらす関数。"""
    if offset == 0:
        return
    for cell in sheet_root.findall(f".//{q(NS_SS, 'c')}"):
        if cell.get("t") != "s":
            continue
        value = cell.find(q(NS_SS, "v"))
        if value is not None and value.text and value.text.isdigit():
            value.text = str(int(value.text) + offset)


def workbook_sheet_parts(members: dict[str, bytes]) -> list[tuple[str, str]]:
    """Excelのworkbook.xmlからシート名とワークシート部品を表示順に取得する関数。"""
    workbook = parse_xml(members, "xl/workbook.xml")
    rels = parse_xml(members, "xl/_rels/workbook.xml.rels")
    rel_targets = {
        rel.get("Id"): part_from_relationship_target("xl/workbook.xml", rel.get("Target", ""))
        for rel in rels.findall(q(NS_REL, "Relationship"))
        if rel.get("Type") == REL_WORKSHEET
    }
    sheets = workbook.find(q(NS_SS, "sheets"))
    if sheets is None:
        return []

    result: list[tuple[str, str]] = []
    for sheet in sheets.findall(q(NS_SS, "sheet")):
        rel_id = sheet.get(q(NS_R, "id"))
        if rel_id in rel_targets:
            result.append((sheet.get("name", "incoming"), rel_targets[rel_id]))
    return result


def unique_sheet_name(existing: set[str], base_name: str) -> str:
    """Excelの31文字制限と重複を避けたシート名を作る関数。"""
    clean_base = (base_name or "incoming").replace(":", "_").replace("\\", "_").replace("/", "_")
    clean_base = clean_base.replace("?", "_").replace("*", "_").replace("[", "_").replace("]", "_")
    prefix = f"incoming_{clean_base}"[:31]
    candidate = prefix
    index = 1
    while candidate in existing:
        suffix = f"_{index}"
        candidate = (prefix[: 31 - len(suffix)] + suffix)[:31]
        index += 1
    existing.add(candidate)
    return candidate


def shared_strings_changed(base: dict[str, bytes], current: dict[str, bytes], incoming: dict[str, bytes]) -> bool:
    """Excelの共有文字列XMLにincoming差分があるかを判定する関数。"""
    return xml_member_changed(
        base,
        current,
        incoming,
        "xl/sharedStrings.xml",
        "xl/sharedStrings.xml",
        "xl/sharedStrings.xml",
    )


def append_xlsx_incoming(
    base: dict[str, bytes],
    current: dict[str, bytes],
    incoming: dict[str, bytes],
    report_entries: list[MergeReportEntry],
) -> None:
    """incoming側で差分のあるExcelシートだけをcurrent側の末尾へ追加する関数。"""
    workbook = parse_xml(current, "xl/workbook.xml")
    workbook_rels = parse_xml(current, "xl/_rels/workbook.xml.rels")
    sheets = workbook.find(q(NS_SS, "sheets"))
    if sheets is None:
        sheets = ET.SubElement(workbook, q(NS_SS, "sheets"))

    existing_names = {sheet.get("name", "") for sheet in sheets.findall(q(NS_SS, "sheet"))}
    used_sheet_ids = {
        int(sheet.get("sheetId", "0"))
        for sheet in sheets.findall(q(NS_SS, "sheet"))
        if sheet.get("sheetId", "0").isdigit()
    }
    used_sheet_numbers = numbers_from_paths(list(current), "sheet", ".xml")
    shared_string_diff = shared_strings_changed(base, current, incoming)
    base_sheets = workbook_sheet_parts(base) if base else []
    current_sheets = workbook_sheet_parts(current)
    incoming_sheets = workbook_sheet_parts(incoming)
    sheets_to_append: list[tuple[str, str, str | None, str | None, str]] = []

    for index, (incoming_name, source_sheet) in enumerate(incoming_sheets):
        base_sheet = base_sheets[index][1] if index < len(base_sheets) else None
        current_sheet = current_sheets[index][1] if index < len(current_sheets) else None
        sheet_diff = xml_member_changed(base, current, incoming, base_sheet, current_sheet, source_sheet)
        if sheet_diff:
            sheets_to_append.append((incoming_name, source_sheet, base_sheet, current_sheet, source_sheet))
        elif shared_string_diff:
            sheets_to_append.append(
                (
                    incoming_name,
                    source_sheet,
                    "xl/sharedStrings.xml",
                    "xl/sharedStrings.xml",
                    "xl/sharedStrings.xml",
                )
            )

    if not sheets_to_append:
        return

    shared_string_offset = ensure_shared_strings(current, incoming)

    next_sheet_id = max(used_sheet_ids | {0}) + 1
    for incoming_name, source_sheet, compare_base, compare_current, compare_incoming in sheets_to_append:
        sheet_number = next_numeric_suffix(used_sheet_numbers)
        used_sheet_numbers.add(sheet_number)
        target_sheet = f"xl/worksheets/sheet{sheet_number}.xml"

        sheet_root = parse_xml(incoming, source_sheet)
        remap_sheet_shared_strings(sheet_root, shared_string_offset)
        current[target_sheet] = xml_bytes(sheet_root)
        merge_content_types_for_part(incoming, current, source_sheet, target_sheet)
        copy_related_parts(incoming, current, source_sheet, target_sheet, {source_sheet: target_sheet})

        rel_id = add_relationship(workbook_rels, REL_WORKSHEET, f"worksheets/sheet{sheet_number}.xml")
        ET.SubElement(
            sheets,
            q(NS_SS, "sheet"),
            {
                "name": unique_sheet_name(existing_names, incoming_name),
                "sheetId": str(next_sheet_id),
                q(NS_R, "id"): rel_id,
            },
        )
        report_entries.append(
            MergeReportEntry(
                kind="xlsx-sheet",
                added_part=target_sheet,
                base_part=compare_base,
                current_part=compare_current,
                incoming_part=compare_incoming,
                note=comparison_note(compare_base, compare_current, compare_incoming),
            )
        )
        next_sheet_id += 1

    current["xl/workbook.xml"] = xml_bytes(workbook)
    current["xl/_rels/workbook.xml.rels"] = xml_bytes(workbook_rels)


def detect_kind(path: Path, members: dict[str, bytes]) -> str:
    """拡張子と主要部品の有無からOOXML種別を判定する関数。"""
    suffix = path.suffix.lower()
    if suffix == ".tar":
        suffix = Path(path.stem).suffix.lower()
    if suffix in {".pptx", ".pptm"} or "ppt/presentation.xml" in members:
        return "pptx"
    if suffix in {".docx", ".docm"} or "word/document.xml" in members:
        return "docx"
    if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm", ".xlmx"} or "xl/workbook.xml" in members:
        return "xlsx"
    raise ValueError(f"unsupported OOXML type: {path}")


def merge_ooxml(base_path: Path, current_path: Path, incoming_path: Path, worktree_path: str | None = None) -> None:
    """Git merge driverから呼ばれ、incoming側のXML差分をcurrent側へ追加して保存する関数。"""
    register_namespaces()
    base = read_package_or_empty(base_path)
    current = read_package(current_path)
    incoming = read_package(incoming_path)
    kind = detect_kind(Path(worktree_path or current_path), current)
    report_entries: list[MergeReportEntry] = []

    if kind == "pptx":
        append_pptx_incoming(base, current, incoming, report_entries)
    elif kind == "docx":
        append_docx_incoming(base, current, incoming, report_entries)
    elif kind == "xlsx":
        append_xlsx_incoming(base, current, incoming, report_entries)
    else:
        raise ValueError(f"unsupported OOXML type: {kind}")

    write_package(current_path, current)
    report_path = write_merge_report(report_entries, worktree_path, kind)
    if report_path:
        log(f"ooxml append merge report: {report_path}")
    log(f"ooxml append merge success: {worktree_path or current_path}")


def install_git_config() -> None:
    """このmerge driverをローカルGit設定へ登録する関数。"""
    commands = [
        ["git", "config", "--local", "merge.ooxml-append.name", "OOXML append incoming merge"],
        [
            "git",
            "config",
            "--local",
            "merge.ooxml-append.driver",
            "python -B scripts/margetools/ooxml_append_merge.py merge %O %A %B %P",
        ],
    ]
    for command in commands:
        result = subprocess.run(command, capture_output=True, text=True)
        if result.returncode == 0:
            log(f"{' '.join(command)} success")
        else:
            log(f"{' '.join(command)} failed")
            if result.stderr:
                log(result.stderr.strip())
            result.check_returncode()


def build_parser() -> argparse.ArgumentParser:
    """コマンドライン引数を解釈するパーサーを作る関数。"""
    parser = argparse.ArgumentParser(description="Append incoming OOXML changes during Git merge.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    merge_parser = subparsers.add_parser("merge", help="merge OOXML by appending incoming content")
    merge_parser.add_argument("base")
    merge_parser.add_argument("current")
    merge_parser.add_argument("incoming")
    merge_parser.add_argument("path", nargs="?")

    subparsers.add_parser("install", help="install local git merge driver settings")
    return parser


def main(argv: list[str] | None = None) -> int:
    """サブコマンドに応じてOOXML merge toolの処理を実行する関数。"""
    args = build_parser().parse_args(argv)
    if args.command == "merge":
        merge_ooxml(Path(args.base), Path(args.current), Path(args.incoming), args.path)
    elif args.command == "install":
        install_git_config()
    else:
        raise ValueError(f"unknown command: {args.command}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
