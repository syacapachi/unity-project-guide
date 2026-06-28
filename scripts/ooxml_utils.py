from __future__ import annotations

import posixpath
import sys
import zipfile
from io import BytesIO
from pathlib import Path


FIXED_TIME = (1980, 1, 1, 0, 0, 0)


def read_all_stdin() -> bytes:
    """標準入力からGit filterに渡されたファイル内容をすべて読み込む関数。"""
    return sys.stdin.buffer.read()


def write_all_stdout(data: bytes) -> None:
    """標準出力へGit filterの変換結果をバイナリのまま書き出す関数。"""
    sys.stdout.buffer.write(data)


def is_zip_data(data: bytes) -> bool:
    """受け取ったバイト列がzipとして読めるかを判定する関数。"""
    try:
        with zipfile.ZipFile(BytesIO(data)) as archive:
            return archive.testzip() is None
    except zipfile.BadZipFile:
        return False


def _safe_member_path(root: Path, name: str) -> Path:
    """zip内のパスを検証し、展開先ディレクトリ外へ出ない安全なパスへ変換する関数。"""
    raw_parts = name.replace("\\", "/").split("/")
    normalized = posixpath.normpath("/".join(raw_parts))
    if ".." in raw_parts or normalized.startswith("../") or normalized == ".." or posixpath.isabs(normalized):
        raise ValueError(f"unsafe zip member path: {name}")
    return root / Path(*normalized.split("/"))


def safe_archive_name(name: str) -> str:
    """zip内のパスを検証し、安全なアーカイブ名として返す関数。"""
    raw_parts = name.replace("\\", "/").split("/")
    normalized = posixpath.normpath("/".join(raw_parts))
    if ".." in raw_parts or normalized in {"", "."} or normalized.startswith("../") or normalized == ".." or posixpath.isabs(normalized):
        raise ValueError(f"unsafe zip member path: {name}")
    return normalized


def unpack_zip_data(data: bytes, dst: Path) -> None:
    """zipデータを指定ディレクトリへ安全に展開する関数。"""
    with zipfile.ZipFile(BytesIO(data)) as archive:
        for name in sorted(archive.namelist()):
            if name.endswith("/"):
                continue
            out_path = _safe_member_path(dst, name)
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_bytes(archive.read(name))


def write_zip_entry(archive: zipfile.ZipFile, arcname: str, data: bytes) -> None:
    """固定されたzip属性で1ファイル分のエントリを書き込む関数。"""
    info = zipfile.ZipInfo(arcname)
    info.date_time = FIXED_TIME
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o644 << 16
    archive.writestr(info, data)
