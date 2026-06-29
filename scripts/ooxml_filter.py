from __future__ import annotations

import argparse
import shutil
import subprocess
from pathlib import Path

from ooxml_utils import (
    read_all_stdin,
    write_all_stdout,
    write_text_stdout_utf8,
)
from ooxml_storage import (
    container_to_zip_data,
    directory_to_zip_data,
    is_ooxml_tar_data,
    unpack_container_data,
    zip_to_tar_data,
)
from ooxml_summary import summarize_ooxml_for_textconv


def clean_stdin() -> None:
    """git add時に呼ばれ、OOXML zipを正規化済みtarへ変換する関数。"""
    write_all_stdout(zip_to_tar_data(read_all_stdin()))


def smudge_stdin() -> None:
    """git checkout時に呼ばれ、Git保存用tarを作業ツリー用OOXML zipへ戻す関数。"""
    data = read_all_stdin()
    write_all_stdout(container_to_zip_data(data) if is_ooxml_tar_data(data) else data)


def pack_directory(src: Path, dst: Path) -> None:
    """指定ディレクトリをOOXML向けに正規化してzipファイルへまとめる関数。"""
    dst.write_bytes(directory_to_zip_data(src))


def unpack_file(src: Path, dst: Path) -> None:
    """ZIP/TAR形式のOOXML保存データを指定ディレクトリへ安全に展開する関数。"""
    if dst.exists():
        shutil.rmtree(dst)
    dst.mkdir(parents=True)
    unpack_container_data(src.read_bytes(), dst)


def textconv_data(data: bytes) -> str:
    """git diff用にZIP/TAR形式のOOXMLをテキスト表現へ変換する関数。"""
    normalized = container_to_zip_data(data)
    return summarize_ooxml_for_textconv(normalized)


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
        ["git", "config", "--local", "diff.ooxml.cachetextconv", "false"],
    ]
    for command in commands:
        subprocess.run(command, check=True)


def build_parser() -> argparse.ArgumentParser:
    """コマンドライン引数を解釈するパーサーを作る関数。"""
    # -h, --helpを設定
    parser = argparse.ArgumentParser(description="Normalize Office Open XML files for Git.")
    # 第一引数チェック
    subparsers = parser.add_subparsers(dest="command", required=True)

    # Git filterとして呼び出されるclean/smudgeサブコマンドを設定
    subparsers.add_parser("clean", help="read OOXML zip from stdin and write normalized tar to stdout")
    subparsers.add_parser("smudge", help="read normalized tar from stdin and write OOXML zip to stdout")

    pack_parser = subparsers.add_parser("pack", help="pack an unpacked OOXML directory")
    #python this.py pack [src] [dst] のように引数を名前としてアクセスできるようにする
    pack_parser.add_argument("src")
    pack_parser.add_argument("dst")

    unpack_parser = subparsers.add_parser("unpack", help="unpack an OOXML zip or filter tar file")
    unpack_parser.add_argument("src")
    unpack_parser.add_argument("dst")

    textconv_parser = subparsers.add_parser("textconv", help="convert OOXML zip or filter tar to text for git diff")
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
