from __future__ import annotations

from pathlib import Path
import re
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
    #(NS["dc"], "creator"): "",
    #(NS["cp"], "lastModifiedBy"): "",
    #(NS["cp"], "revision"): "1",
    #(NS["dcterms"], "created"): FIXED_ISO_TIME,
    #(NS["dcterms"], "modified"): FIXED_ISO_TIME,
}
# Replace()する辞書
VOLATILE_APP_VALUES = {
    #(NS["ep"], "TotalTime"): "0",
    #(NS["ep"], "AppVersion"): "",
}


def _decode_xml(data: bytes) -> tuple[str, str]:
    """XML宣言やBOMから主要な文字コードを推定し、文字列へ変換する関数。"""
    if data.startswith(b"\xff\xfe") or data.startswith(b"\xfe\xff"):
        return data.decode("utf-16"), "utf-16"
    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig"), "utf-8"

    head = data[:200].decode("ascii", errors="ignore")
    match = re.search(r'encoding=["\']([^"\']+)["\']', head, flags=re.IGNORECASE)
    encoding = match.group(1) if match else "utf-8"
    return data.decode(encoding), encoding


def _encode_xml(text: str, encoding: str) -> bytes:
    """XML文字列を元の文字コードへ戻す関数。"""
    return text.encode(encoding)


def _set_existing_element_text(xml_text: str, tag_name: str, value: str) -> str:
    """既に存在するXML要素の本文だけを、接頭辞を保ったまま固定値へ置き換える関数。"""
    escaped_value = (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    pattern = re.compile(
        rf"(<(?P<prefix>[A-Za-z_][\w.-]*:)?{re.escape(tag_name)}\b[^>]*>)(.*?)(</(?P=prefix)?{re.escape(tag_name)}>)",
        flags=re.DOTALL,
    )
    return pattern.sub(lambda match: f"{match.group(1)}{escaped_value}{match.group(4)}", xml_text, count=1)


def _normalize_metadata_xml(path: Path, data: bytes) -> bytes:
    """OOXMLの更新日時や編集者など、Officeが書き換えやすいXMLのメタデータを正規化する関数。"""
    normalized_path = path.as_posix()
    values = {}
    if normalized_path == "docProps/core.xml":
        values = VOLATILE_CORE_VALUES
    elif normalized_path == "docProps/app.xml":
        values = VOLATILE_APP_VALUES
    if not values:
        return data

    try:
        ElementTree.fromstring(data)
        xml_text, encoding = _decode_xml(data)
    except (ElementTree.ParseError, UnicodeDecodeError, LookupError):
        return data

    for (_, local_name), value in values.items():
        xml_text = _set_existing_element_text(xml_text, local_name, value)
    return _encode_xml(xml_text, encoding)


def _normalize_xml_data(relative_path: Path, data: bytes) -> bytes:
    """OOXML内のXMLデータを、相対パスに応じたメタデータ処理込みで正規化する関数。"""
    if relative_path.as_posix().startswith("docProps/"):
        return _normalize_metadata_xml(relative_path, data)
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
