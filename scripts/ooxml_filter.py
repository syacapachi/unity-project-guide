from __future__ import annotations

import argparse
import hashlib
import shutil
import subprocess
import zipfile
from io import BytesIO
from pathlib import Path

from ooxml_utils import (
    is_zip_data,
    read_all_stdin,
    safe_archive_name,
    unpack_zip_data,
    write_all_stdout,
    write_text_stdout_utf8,
    write_zip_entry,
)
from ooxml_xml import maybe_normalize_member


def _write_normalized_zip(src: Path, dst: Path) -> None:
    """展開済みディレクトリを固定順序・固定時刻のzipへ再圧縮する関数。"""
    files = sorted(path for path in src.rglob("*") if path.is_file())
    with zipfile.ZipFile(dst, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path in files:
            arcname = path.relative_to(src).as_posix()
            write_zip_entry(archive, arcname, maybe_normalize_member(arcname, path.read_bytes()))


def _normalize_ooxml_zip_data(data: bytes) -> bytes:
    """OOXML zipをメモリ上でXML正規化し、固定条件で再圧縮して返す関数。"""
    if not is_zip_data(data):
        return data

    try:
        input_zip = zipfile.ZipFile(BytesIO(data))
        output = BytesIO()
        with input_zip, zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as output_zip:
            names = sorted(name for name in input_zip.namelist() if not name.endswith("/"))
            for name in names:
                arcname = safe_archive_name(name)
                payload = maybe_normalize_member(arcname, input_zip.read(name))
                write_zip_entry(output_zip, arcname, payload)
        return output.getvalue()
    except (OSError, ValueError, zipfile.BadZipFile):
        return data


def clean_stdin() -> None:
    """git add時に呼ばれ、OOXMLファイルを正規化zipへ変換する関数。"""
    write_all_stdout(_normalize_ooxml_zip_data(read_all_stdin()))


def smudge_stdin() -> None:
    """git checkout時に呼ばれ、保存済みOOXML zipをそのまま作業ツリーへ戻す関数。"""
    write_all_stdout(read_all_stdin())


def pack_directory(src: Path, dst: Path) -> None:
    """指定ディレクトリをOOXML向けに正規化してzipファイルへまとめる関数。"""
    _write_normalized_zip(src, dst)


def unpack_file(src: Path, dst: Path) -> None:
    """zipファイルを指定ディレクトリへ安全に展開する関数。"""
    if dst.exists():
        shutil.rmtree(dst)
    dst.mkdir(parents=True)
    unpack_zip_data(src.read_bytes(), dst)


def textconv_data(data: bytes) -> str:
    """git diff用にOOXML zipの中身をテキスト表現へ変換する関数。"""
    normalized = _normalize_ooxml_zip_data(data)
    if not is_zip_data(normalized):
        return normalized.decode("utf-8", errors="replace")

    lines: list[str] = []
    with zipfile.ZipFile(BytesIO(normalized)) as archive:
        for name in sorted(item for item in archive.namelist() if not item.endswith("/")):
            payload = archive.read(name)
            lines.append(f"--- {name}")
            if name.lower().endswith((".xml", ".rels")):
                lines.append(payload.decode("utf-8", errors="replace"))
            else:
                digest = hashlib.sha256(payload).hexdigest()
                lines.append(f"<binary size={len(payload)} sha256={digest}>")
            lines.append("")
    return "\n".join(lines)


def textconv(path: str | None) -> None:
    """git diffから呼ばれ、ファイルまたは標準入力のOOXMLをテキスト化して出力する関数。"""
    data = Path(path).read_bytes() if path else read_all_stdin()
    write_text_stdout_utf8(textconv_data(data))


def install_git_config() -> None:
    """このリポジトリのローカルGit設定へOOXML filterとdiff driverを登録する関数。"""
    commands = [
        ["git", "config", "--local", "filter.ooxml.clean", "python -B scripts/ooxml_filter.py clean"],
        ["git", "config", "--local", "filter.ooxml.smudge", "python -B scripts/ooxml_filter.py smudge"],
        ["git", "config", "--local", "filter.ooxml.required", "true"],
        ["git", "config", "--local", "diff.ooxml.textconv", "python -B scripts/ooxml_filter.py textconv"],
        ["git", "config", "--local", "diff.ooxml.cachetextconv", "true"],
    ]
    for command in commands:
        subprocess.run(command, check=True)


def build_parser() -> argparse.ArgumentParser:
    """コマンドライン引数を解釈するパーサーを作る関数。"""
    # -h, --helpを設定
    parser = argparse.ArgumentParser(description="Normalize Office Open XML files for Git.")
    # 第一引数チェック
    subparsers = parser.add_subparsers(dest="command", required=True)

    # help-> clean  read OOXML from stdin and write normalized zip to stdout
    subparsers.add_parser("clean", help="read OOXML from stdin and write normalized zip to stdout")
    subparsers.add_parser("smudge", help="pass OOXML from stdin to stdout")

    pack_parser = subparsers.add_parser("pack", help="pack an unpacked OOXML directory")
    #python this.py pack [src] [dst] のように引数を名前としてアクセスできるようにする
    pack_parser.add_argument("src")
    pack_parser.add_argument("dst")

    unpack_parser = subparsers.add_parser("unpack", help="unpack an OOXML zip file")
    unpack_parser.add_argument("src")
    unpack_parser.add_argument("dst")

    textconv_parser = subparsers.add_parser("textconv", help="convert OOXML zip to text for git diff")
    textconv_parser.add_argument("path", nargs="?")

    subparsers.add_parser("install", help="install local git filter and diff settings")
    return parser


def main(argv: list[str] | None = None) -> int:
    """サブコマンドに応じてOOXML filterツールの処理を実行する関数。"""
    args = build_parser().parse_args(argv)
    if args.command == "clean":
        clean_stdin()
    elif args.command == "smudge":
        smudge_stdin()
    elif args.command == "pack":
        pack_directory(Path(args.src), Path(args.dst))
    elif args.command == "unpack":
        unpack_file(Path(args.src), Path(args.dst))
    elif args.command == "textconv":
        textconv(args.path)
    elif args.command == "install":
        install_git_config()
    else:
        raise ValueError(f"unknown command: {args.command}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
