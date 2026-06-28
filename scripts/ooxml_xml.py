from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree


FIXED_ISO_TIME = "1980-01-01T00:00:00Z"

#文字を短縮
NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
}

# Replace()する辞書
VOLATILE_CORE_VALUES = {
    (NS["dc"], "creator"): "",
    (NS["cp"], "lastModifiedBy"): "",
    (NS["cp"], "revision"): "1",
    (NS["dcterms"], "created"): FIXED_ISO_TIME,
    (NS["dcterms"], "modified"): FIXED_ISO_TIME,
}
# Replace()する辞書
VOLATILE_APP_VALUES = {
    (NS["ep"], "TotalTime"): "0",
    (NS["ep"], "AppVersion"): "",
}


def _canonicalize_xml(data: bytes) -> bytes:
    """XMLの属性順や名前空間表現を正規化し、差分が安定する形へ変換する関数。"""
    text = data.decode("utf-8-sig")
    normalized = ElementTree.canonicalize(xml_data=text, with_comments=True, strip_text=False)
    return normalized.encode("utf-8")


def _set_existing_child_text(root: ElementTree.Element, values: dict[tuple[str, str], str]) -> None:
    """既に存在するXMLのメタデータ要素の値を固定値へ置き換える関数。"""
    for child in list(root):
        namespace = ""
        local_name = child.tag
        if child.tag.startswith("{"):
            namespace, local_name = child.tag[1:].split("}", 1)
        value = values.get((namespace, local_name))
        if value is not None:
            child.text = value
            child.tail = child.tail or ""


def _normalize_metadata_xml(path: Path, data: bytes) -> bytes:
    """OOXMLの更新日時や編集者など、Officeが書き換えやすいXMLのメタデータを正規化する関数。"""
    try:
        root = ElementTree.fromstring(data)
    except ElementTree.ParseError:
        return data

    normalized_path = path.as_posix()
    if normalized_path == "docProps/core.xml":
        _set_existing_child_text(root, VOLATILE_CORE_VALUES)
    elif normalized_path == "docProps/app.xml":
        _set_existing_child_text(root, VOLATILE_APP_VALUES)
    else:
        return data

    return ElementTree.tostring(root, encoding="utf-8", xml_declaration=False)


def _normalize_xml_data(relative_path: Path, data: bytes) -> bytes:
    """OOXML内のXMLデータを、相対パスに応じたメタデータ処理込みで正規化する関数。"""
    if relative_path.as_posix().startswith("docProps/"):
        data = _normalize_metadata_xml(relative_path, data)
    try:
        return _canonicalize_xml(data)
    except (ElementTree.ParseError, UnicodeDecodeError):
        return data


def _normalize_xml_file(path: Path, relative_path: Path) -> None:
    """XMLファイルを読み込み、メタデータ固定化とXML正規化を行って保存する関数。"""
    path.write_bytes(_normalize_xml_data(relative_path, path.read_bytes()))


def _normalize_ooxml_tree(root: Path) -> None:
    """展開済みOOXMLディレクトリ内のXML群を正規化する関数。"""
    for path in sorted(root.rglob("*")):
        if not path.is_file():
            continue
        suffix = path.suffix.lower()
        if suffix in {".xml", ".rels"}:
            _normalize_xml_file(path, path.relative_to(root))

# これだけメイン
def maybe_normalize_member(arcname: str, data: bytes) -> bytes:
    """zip内ファイルがXMLまたはrelsなら正規化し、それ以外はそのまま返す関数。"""
    suffix = Path(arcname).suffix.lower()
    if suffix in {".xml", ".rels"}:
        return _normalize_xml_data(Path(*arcname.split("/")), data)
    return data
