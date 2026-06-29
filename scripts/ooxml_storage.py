from __future__ import annotations

import hashlib
import json
import tarfile
import zipfile
from io import BytesIO
from pathlib import Path

from ooxml_utils import is_zip_data, safe_archive_name, unpack_zip_data, write_zip_entry
from ooxml_xml import maybe_normalize_member


OOXML_TAR_FORMAT = "ooxml-filter-tar"
OOXML_TAR_VERSION = 1
OOXML_TAR_PREFIX = "._ooxml_filter/"
OOXML_TAR_MANIFEST = f"{OOXML_TAR_PREFIX}manifest.json"
OOXML_TAR_LARGE_DIR = f"{OOXML_TAR_PREFIX}large"
DEFAULT_LARGE_FILE_THRESHOLD = 100 * 1024 * 1024


def _sha256(data: bytes) -> str:
    """バイト列のSHA-256ハッシュを16進文字列で返す関数。"""
    return hashlib.sha256(data).hexdigest()


def _write_tar_entry(archive: tarfile.TarFile, name: str, data: bytes) -> None:
    """固定されたtar属性で1ファイル分のエントリを書き込む関数。"""
    info = tarfile.TarInfo(name)
    info.size = len(data)
    info.mtime = 0
    info.mode = 0o644
    info.uid = 0
    info.gid = 0
    info.uname = ""
    info.gname = ""
    archive.addfile(info, BytesIO(data))


def _make_large_member_zip(arcname: str, data: bytes) -> bytes:
    """100MB超などの巨大メンバーを単体zipへ退避する関数。"""
    output = BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        write_zip_entry(archive, arcname, data)
    return output.getvalue()


def _read_large_member_zip(archive_data: bytes, arcname: str) -> bytes:
    """退避用zipから元のOOXMLメンバー内容を取り出す関数。"""
    with zipfile.ZipFile(BytesIO(archive_data)) as archive:
        if arcname in archive.namelist():
            return archive.read(arcname)
        names = [name for name in archive.namelist() if not name.endswith("/")]
        if len(names) != 1:
            raise ValueError(f"large member zip does not contain expected file: {arcname}")
        return archive.read(names[0])


def _normalise_zip_members(data: bytes) -> dict[str, bytes]:
    """OOXML zipを読み込み、各メンバー名の検証とXML正規化を行う関数。"""
    members: dict[str, bytes] = {}
    with zipfile.ZipFile(BytesIO(data)) as archive:
        for name in sorted(name for name in archive.namelist() if not name.endswith("/")):
            arcname = safe_archive_name(name)
            members[arcname] = maybe_normalize_member(arcname, archive.read(name))
    return members


def _members_to_zip_data(members: dict[str, bytes]) -> bytes:
    """正規化済みOOXMLメンバー群を固定条件のzipへ戻す関数。"""
    output = BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for arcname in sorted(members):
            write_zip_entry(archive, arcname, members[arcname])
    return output.getvalue()


def _manifest_data(large_files: list[dict[str, object]], large_threshold: int) -> bytes:
    """TAR保存形式のメタ情報をUTF-8 JSONとして作る関数。"""
    manifest = {
        "format": OOXML_TAR_FORMAT,
        "version": OOXML_TAR_VERSION,
        "large_threshold": large_threshold,
        "large_files": large_files,
    }
    text = json.dumps(manifest, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return text.encode("utf-8")


def _read_manifest(data: bytes) -> dict[str, object]:
    """TAR保存形式のメタ情報JSONを読み込んで検証する関数。"""
    manifest = json.loads(data.decode("utf-8"))
    if manifest.get("format") != OOXML_TAR_FORMAT:
        raise ValueError("not an OOXML filter tar manifest")
    if manifest.get("version") != OOXML_TAR_VERSION:
        raise ValueError(f"unsupported OOXML filter tar version: {manifest.get('version')}")
    return manifest


def members_to_tar_data(
    members: dict[str, bytes],
    large_threshold: int = DEFAULT_LARGE_FILE_THRESHOLD,
) -> bytes:
    """正規化済みOOXMLメンバー群をGit保存用tarへ変換する関数。"""
    output = BytesIO()
    large_files: list[dict[str, object]] = []

    with tarfile.open(fileobj=output, mode="w", format=tarfile.PAX_FORMAT) as archive:
        for arcname in sorted(members):
            payload = members[arcname]
            if arcname.startswith(OOXML_TAR_PREFIX):
                raise ValueError(f"OOXML member uses reserved filter path: {arcname}")

            if len(payload) > large_threshold:
                digest = _sha256(payload)
                zip_path = f"{OOXML_TAR_LARGE_DIR}/{digest}.zip"
                _write_tar_entry(archive, zip_path, _make_large_member_zip(arcname, payload))
                large_files.append(
                    {
                        "path": arcname,
                        "zip_path": zip_path,
                        "size": len(payload),
                        "sha256": digest,
                    }
                )
                continue

            _write_tar_entry(archive, arcname, payload)

        _write_tar_entry(archive, OOXML_TAR_MANIFEST, _manifest_data(large_files, large_threshold))

    return output.getvalue()


def is_ooxml_tar_data(data: bytes) -> bool:
    """受け取ったバイト列がこのツールのOOXML保存用tarかを判定する関数。"""
    try:
        with tarfile.open(fileobj=BytesIO(data), mode="r:*") as archive:
            return OOXML_TAR_MANIFEST in archive.getnames()
    except (tarfile.TarError, EOFError, OSError):
        return False


def tar_to_members(data: bytes) -> dict[str, bytes]:
    """Git保存用tarから正規化済みOOXMLメンバー群を復元する関数。"""
    members: dict[str, bytes] = {}
    large_payloads: dict[str, bytes] = {}
    manifest: dict[str, object] | None = None

    with tarfile.open(fileobj=BytesIO(data), mode="r:*") as archive:
        for member in archive.getmembers():
            if not member.isfile():
                continue
            extracted = archive.extractfile(member)
            if extracted is None:
                continue
            payload = extracted.read()
            arcname = safe_archive_name(member.name)

            if arcname == OOXML_TAR_MANIFEST:
                manifest = _read_manifest(payload)
            elif arcname.startswith(OOXML_TAR_LARGE_DIR + "/"):
                large_payloads[arcname] = payload
            elif not arcname.startswith(OOXML_TAR_PREFIX):
                members[arcname] = payload

    if manifest is None:
        raise ValueError("OOXML filter tar manifest is missing")

    for item in manifest.get("large_files", []):
        if not isinstance(item, dict):
            raise ValueError("invalid large file manifest entry")
        path = safe_archive_name(str(item["path"]))
        zip_path = safe_archive_name(str(item["zip_path"]))
        payload = _read_large_member_zip(large_payloads[zip_path], path)
        digest = _sha256(payload)
        if digest != item.get("sha256"):
            raise ValueError(f"large OOXML member checksum mismatch: {path}")
        members[path] = payload

    return members


def normalize_zip_data(data: bytes) -> bytes:
    """OOXML zipを正規化し、固定条件のzipとして返す関数。"""
    if not is_zip_data(data):
        return data
    try:
        return _members_to_zip_data(_normalise_zip_members(data))
    except (OSError, ValueError, zipfile.BadZipFile):
        return data


def zip_to_tar_data(data: bytes, large_threshold: int = DEFAULT_LARGE_FILE_THRESHOLD) -> bytes:
    """git add時にOOXML zipを正規化済みtarへ変換する関数。"""
    if not is_zip_data(data):
        return data
    try:
        return members_to_tar_data(_normalise_zip_members(data), large_threshold=large_threshold)
    except (OSError, ValueError, zipfile.BadZipFile, tarfile.TarError):
        return data


def tar_to_zip_data(data: bytes) -> bytes:
    """git checkout時にGit保存用tarをOOXML zipへ戻す関数。"""
    return _members_to_zip_data(tar_to_members(data))


def container_to_zip_data(data: bytes) -> bytes:
    """ZIP/TARどちらのOOXML保存形式でも、差分処理用のzipへそろえる関数。"""
    if is_ooxml_tar_data(data):
        try:
            return tar_to_zip_data(data)
        except (OSError, ValueError, zipfile.BadZipFile, tarfile.TarError, KeyError):
            return data
    return normalize_zip_data(data)


def directory_to_zip_data(src: Path) -> bytes:
    """展開済みOOXMLディレクトリを正規化済みzipデータへ変換する関数。"""
    members: dict[str, bytes] = {}
    for path in sorted(path for path in src.rglob("*") if path.is_file()):
        arcname = safe_archive_name(path.relative_to(src).as_posix())
        members[arcname] = maybe_normalize_member(arcname, path.read_bytes())
    return _members_to_zip_data(members)


def unpack_container_data(data: bytes, dst: Path) -> None:
    """ZIP/TARどちらのOOXML保存形式でも、指定ディレクトリへ展開する関数。"""
    unpack_zip_data(container_to_zip_data(data), dst)
