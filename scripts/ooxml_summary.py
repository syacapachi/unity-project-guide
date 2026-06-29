from __future__ import annotations

import argparse
import subprocess
from pathlib import Path

from docx_summary import is_docx_package, summarize_docx_inventory
from ooxml_utils import read_all_stdin, write_text_stdout_utf8
from pptx_summary import is_pptx_package, summarize_pptx_diff, summarize_pptx_inventory
from summary_utils import generic_zip_textconv
from xlsx_summary import is_xlsx_package, summarize_xlsx_inventory


__all__ = ["summarize_ooxml_for_textconv"]


def summarize_ooxml_for_textconv(data: bytes) -> str:
    """OOXMLバイト列を形式判定し、git textconv向けの人が読みやすい要約を返す関数。"""
    if is_pptx_package(data):
        return summarize_pptx_inventory(data)
    if is_docx_package(data):
        return summarize_docx_inventory(data)
    if is_xlsx_package(data):
        return summarize_xlsx_inventory(data)
    return generic_zip_textconv(data)


def _summarize_ooxml_files(old_path: Path, new_path: Path) -> str:
    """2つのOOXMLファイルを読み込み、対応形式の差分要約を返す関数。"""
    old_data = old_path.read_bytes()
    new_data = new_path.read_bytes()
    if is_pptx_package(old_data) and is_pptx_package(new_data):
        return summarize_pptx_diff(old_data, new_data)
    return "\n".join(["## 旧ファイル要約", summarize_ooxml_for_textconv(old_data), "", "## 新ファイル要約", summarize_ooxml_for_textconv(new_data)])


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


def _summarize_git_worktree_file(path: Path, ref: str = "HEAD") -> str:
    """Gitの指定ref上のファイルと作業ツリー上のファイルを比較し、差分要約を返す関数。"""
    git_path = _git_relative_path(path)
    old_data = subprocess.check_output(["git", "show", f"{ref}:{git_path}"])
    new_data = path.read_bytes()
    if is_pptx_package(old_data) and is_pptx_package(new_data):
        return summarize_pptx_diff(old_data, new_data)
    return "\n".join(["## 旧ファイル要約", summarize_ooxml_for_textconv(old_data), "", "## 新ファイル要約", summarize_ooxml_for_textconv(new_data)])


def _build_parser() -> argparse.ArgumentParser:
    """コマンドライン引数を解釈するパーサーを作る関数。"""
    parser = argparse.ArgumentParser(description="Summarize OOXML differences for humans.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    files_parser = subparsers.add_parser("files", help="summarize two OOXML files")
    files_parser.add_argument("old")
    files_parser.add_argument("new")

    git_parser = subparsers.add_parser("git", help="summarize ref vs worktree for one file")
    git_parser.add_argument("path")
    git_parser.add_argument("--ref", default="HEAD")

    subparsers.add_parser("textconv", help="summarize one OOXML file from stdin")
    return parser


def _main(argv: list[str] | None = None) -> int:
    """サブコマンドに応じてOOXML差分要約を実行する関数。"""
    args = _build_parser().parse_args(argv)
    if args.command == "files":
        summary = _summarize_ooxml_files(Path(args.old), Path(args.new))
    elif args.command == "git":
        summary = _summarize_git_worktree_file(Path(args.path), args.ref)
    elif args.command == "textconv":
        summary = summarize_ooxml_for_textconv(read_all_stdin())
    else:
        raise ValueError(f"unknown command: {args.command}")
    write_text_stdout_utf8(summary + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
