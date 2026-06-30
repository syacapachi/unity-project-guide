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
    directory_to_tar_data,
    is_ooxml_tar_data,
    unpack_container_data,
    zip_to_tar_data,
)
from ooxml_summary import summarize_ooxml_for_textconv


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


def clean_stdin() -> None:
    """git add時に呼ばれ、OOXML zipを正規化済みtarへ変換する関数。"""
    write_all_stdout(zip_to_tar_data(read_all_stdin()))


def smudge_stdin() -> None:
    """git checkout時に呼ばれ、Git保存用tarを作業ツリー用OOXML zipへ戻す関数。"""
    data = read_all_stdin()
    write_all_stdout(container_to_zip_data(data) if is_ooxml_tar_data(data) else data)


def _tar_output_path(dst: Path) -> Path:
    """packの出力先を、分かりやすい.tar拡張子のパスへそろえる関数。"""
    if dst.suffix.lower() == ".tar":
        return dst
    return dst.with_name(dst.name + ".tar")


def pack_directory(src: Path, dst: Path) -> None:
    """指定ディレクトリをOOXML向けに正規化してGit保存用tarへまとめる関数。"""
    output_path = _tar_output_path(dst)
    output_path.write_bytes(directory_to_tar_data(src))
    _log(f"pack {src} -> {output_path} success")


def unpack_file(src: Path, dst: Path) -> None:
    """TARをZIP再圧縮してから、指定ディレクトリへ安全に展開する関数。"""
    if dst.exists():
        shutil.rmtree(dst)
    dst.mkdir(parents=True)
    unpack_container_data(src.read_bytes(), dst)
    _log(f"unpack {src} -> {dst} success")


def textconv_data(data: bytes) -> str:
    """git diff用にZIP/TAR形式のOOXMLをテキスト表現へ変換する関数。"""
    normalized = container_to_zip_data(data)
    return summarize_ooxml_for_textconv(normalized)


def textconv(path: str | None) -> None:
    """git diffから呼ばれ、ファイルまたは標準入力のOOXMLをテキスト化して出力する関数。"""
    data = Path(path).read_bytes() if path else read_all_stdin()
    write_text_stdout_utf8(textconv_data(data))


def _log(message: str) -> None:
    """日本語パスを含むログでも文字コードエラーにしないためUTF-8で出力する関数。"""
    write_text_stdout_utf8(message + "\n")


def _command_to_text(command: list[str]) -> str:
    """ログ表示用にsubprocessのコマンド配列を読みやすい文字列へ変換する関数。"""
    return " ".join(command)


def _run_logged(command: list[str]) -> None:
    """install時に実行したGit設定コマンドの成功・失敗をログへ出す関数。"""
    result = subprocess.run(command, capture_output=True, text=True)
    command_text = _command_to_text(command)
    if result.returncode == 0:
        _log(f"{command_text} success")
        return

    _log(f"{command_text} failed")
    if result.stderr:
        _log(result.stderr.strip())
    result.check_returncode()


def _is_ooxml_path(path: Path) -> bool:
    """パスがOOXML対象拡張子かどうかを判定する関数。"""
    return path.suffix.lower() in OOXML_EXTENSIONS


def _is_ooxml_tar_path(path: Path) -> bool:
    """パスが.pptx.tarなどのOOXML用tar名かどうかを判定する関数。"""
    return path.suffix.lower() == ".tar" and Path(path.stem).suffix.lower() in OOXML_EXTENSIONS


def _smudge_target_path(path: Path) -> Path:
    """smudge後に書き戻す元のOOXML拡張子のパスを返す関数。"""
    if _is_ooxml_tar_path(path):
        return path.with_suffix("")
    return path


def _iter_smudge_candidates(root: Path):
    """install後の復元対象になりうるOOXMLファイル群を列挙する関数。"""
    for path in sorted(root.rglob("*")):
        if ".git" in path.parts or not path.is_file():
            continue
        if _is_ooxml_path(path) or _is_ooxml_tar_path(path):
            yield path


def _smudge_worktree_file(path: Path) -> bool:
    """作業ツリー上のTAR内容ファイルをOOXML zipへ復元する関数。"""
    data = path.read_bytes()
    if not is_ooxml_tar_data(data):
        return False

    target = _smudge_target_path(path)
    restored = container_to_zip_data(data)
    if target != path and target.exists():
        existing = target.read_bytes()
        if existing != data and existing != restored:
            _log(f"smudge {path} -> {target} skipped: target exists")
            return False

    target.write_bytes(restored)
    if target != path:
        try:
            path.unlink()
        except OSError as exc:
            _log(f"smudge {path} source kept: {exc}")
    _log(f"smudge {path} -> {target} success")
    return True


def smudge_worktree_files(root: Path) -> int:
    """install直後に、作業ツリーへ残ったTAR内容のOOXMLをまとめて復元する関数。"""
    restored_count = 0
    for path in _iter_smudge_candidates(root):
        try:
            if _smudge_worktree_file(path):
                restored_count += 1
        except (OSError, ValueError) as exc:
            _log(f"smudge {path} failed: {exc}")
    _log(f"worktree smudge success: {restored_count} file(s)")
    return restored_count


def install_git_config() -> None:
    """ローカルGit設定へOOXML filterとdiff driverを登録し、作業ツリーを復元する関数。"""
    commands = [
        ["git", "config", "--local", "filter.ooxml.clean", "python -B scripts/ooxml_filter.py clean"],
        ["git", "config", "--local", "filter.ooxml.smudge", "python -B scripts/ooxml_filter.py smudge"],
        ["git", "config", "--local", "filter.ooxml.required", "true"],
        ["git", "config", "--local", "diff.ooxml.textconv", "python -B scripts/ooxml_filter.py textconv"],
        ["git", "config", "--local", "diff.ooxml.cachetextconv", "false"],
    ]
    for command in commands:
        _run_logged(command)
    smudge_worktree_files(Path("."))


def build_parser() -> argparse.ArgumentParser:
    """コマンドライン引数を解釈するパーサーを作る関数。"""
    # -h, --helpを設定
    parser = argparse.ArgumentParser(description="Normalize Office Open XML files for Git.")
    # 第一引数チェック
    subparsers = parser.add_subparsers(dest="command", required=True)

    # Git filterとして呼び出されるclean/smudgeサブコマンドを設定
    subparsers.add_parser("clean", help="read OOXML zip from stdin and write normalized tar to stdout")
    subparsers.add_parser("smudge", help="read normalized tar from stdin and write OOXML zip to stdout")

    pack_parser = subparsers.add_parser("pack", help="pack an unpacked OOXML directory into filter tar")
    #python this.py pack [src] [dst] のように引数を名前としてアクセスできるようにする
    pack_parser.add_argument("src")
    pack_parser.add_argument("dst")

    unpack_parser = subparsers.add_parser("unpack", help="restore an OOXML zip from filter tar and unpack it")
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
