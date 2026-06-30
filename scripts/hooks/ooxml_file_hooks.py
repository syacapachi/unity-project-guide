from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tarfile
from io import BytesIO
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = SCRIPT_DIR.parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from ooxml_storage import OOXML_TAR_MANIFEST, container_to_zip_data, zip_to_tar_data
from ooxml_utils import safe_archive_name


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
TREE_SUFFIX = ".ooxml"


def log(message: str) -> None:
    """日本語パスを含むログでも文字コードエラーにしないためUTF-8で出力する関数。"""
    sys.stderr.buffer.write((message + "\n").encode("utf-8", errors="replace"))


def run_git(args: list[str], check: bool = True) -> subprocess.CompletedProcess[bytes]:
    """Gitコマンドをshellを介さず実行し、必要なら失敗時に例外を投げる関数。"""
    command = ["git", "-c", "core.quotepath=false", *args]
    result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if check and result.returncode != 0:
        stderr = result.stderr.decode("utf-8", errors="replace").strip()
        raise RuntimeError(f"{' '.join(command)} failed: {stderr}")
    return result


def is_ooxml_office_path(path: Path) -> bool:
    """パスがhook管理対象のOOXML Officeファイルかを判定する関数。"""
    return path.suffix.lower() in OOXML_EXTENSIONS


def office_path_to_tree_path(path: Path) -> Path:
    """OOXML Officeファイルのパスから、Git管理用展開ディレクトリのパスを作る関数。"""
    return Path(str(path) + TREE_SUFFIX)


def tree_path_to_office_path(path: Path) -> Path:
    """Git管理用展開ディレクトリのパスから、復元先Officeファイルのパスを作る関数。"""
    text = str(path)
    if not text.endswith(TREE_SUFFIX):
        raise ValueError(f"not an OOXML tree path: {path}")
    return Path(text[: -len(TREE_SUFFIX)])


def is_ooxml_tree_dir(path: Path) -> bool:
    """ディレクトリがOOXML展開ツリーかを判定する関数。"""
    return path.is_dir() and path.name.endswith(TREE_SUFFIX) and (path / OOXML_TAR_MANIFEST).is_file()


def safe_tree_member_path(root: Path, arcname: str) -> Path:
    """TAR内のメンバー名を検証し、展開先ディレクトリ内の安全なパスへ変換する関数。"""
    safe_name = safe_archive_name(arcname)
    return root / Path(*safe_name.split("/"))


def remove_generated_tree(path: Path) -> None:
    """生成済みOOXML展開ディレクトリを安全に削除する関数。"""
    if path.exists() and not path.name.endswith(TREE_SUFFIX):
        raise ValueError(f"refuse to remove non-OOXML tree: {path}")
    if path.exists():
        shutil.rmtree(path)


def extract_tar_to_tree(tar_data: bytes, dst: Path) -> None:
    """既存storage形式のTARデータを、Git管理用のファイル群へ展開する関数。"""
    remove_generated_tree(dst)
    dst.mkdir(parents=True, exist_ok=True)
    with tarfile.open(fileobj=BytesIO(tar_data), mode="r:*") as archive:
        for member in archive.getmembers():
            if not member.isfile():
                continue
            extracted = archive.extractfile(member)
            if extracted is None:
                continue
            out_path = safe_tree_member_path(dst, member.name)
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_bytes(extracted.read())


def write_tar_entry(archive: tarfile.TarFile, name: str, data: bytes) -> None:
    """固定されたTAR属性で1ファイル分のエントリを書き込む関数。"""
    info = tarfile.TarInfo(name)
    info.size = len(data)
    info.mtime = 0
    info.mode = 0o644
    info.uid = 0
    info.gid = 0
    info.uname = ""
    info.gname = ""
    archive.addfile(info, BytesIO(data))


def tree_to_tar_data(src: Path) -> bytes:
    """Git管理用のファイル群を、既存storage形式のTARデータへまとめる関数。"""
    if not is_ooxml_tree_dir(src):
        raise ValueError(f"not an OOXML tree directory: {src}")
    output = BytesIO()
    with tarfile.open(fileobj=output, mode="w", format=tarfile.PAX_FORMAT) as archive:
        for path in sorted(item for item in src.rglob("*") if item.is_file()):
            arcname = safe_archive_name(path.relative_to(src).as_posix())
            write_tar_entry(archive, arcname, path.read_bytes())
    return output.getvalue()


def staged_paths() -> list[Path]:
    """pre-commit時にステージされているパス一覧をNUL区切りで取得する関数。"""
    result = run_git(["diff", "--cached", "--name-only", "-z", "--diff-filter=ACMRD"])
    raw_paths = [part for part in result.stdout.split(b"\0") if part]
    return [Path(part.decode("utf-8", errors="replace")) for part in raw_paths]


def stage_path(path: Path) -> None:
    """指定パスをGit indexへ追加する関数。"""
    run_git(["add", "--", path.as_posix()])


def unstage_path(path: Path) -> None:
    """指定パスをGit indexから外す関数。"""
    run_git(["restore", "--staged", "--", path.as_posix()])


def stage_removed_tree(path: Path) -> None:
    """削除されたOOXML Officeファイルに対応する展開ツリーの削除をステージする関数。"""
    if path.exists():
        remove_generated_tree(path)
    run_git(["add", "-u", "--", path.as_posix()], check=False)


def convert_office_to_tree(path: Path) -> bool:
    """OOXML OfficeファイルをTAR形式経由で展開ツリーへ変換し、Git管理対象にする関数。"""
    if not path.exists():
        tree_path = office_path_to_tree_path(path)
        stage_removed_tree(tree_path)
        unstage_path(path)
        log(f"ooxml hook removed staged Office file and updated tree: {path}")
        return True

    data = path.read_bytes()
    tar_data = zip_to_tar_data(data)
    if tar_data == data:
        log(f"ooxml hook skipped non-zip Office file: {path}")
        return False

    tree_path = office_path_to_tree_path(path)
    extract_tar_to_tree(tar_data, tree_path)
    stage_path(tree_path)
    unstage_path(path)
    log(f"ooxml hook staged tree and unstaged Office file: {path} -> {tree_path}")
    return True


def pre_commit() -> int:
    """pre-commit hookから呼ばれ、ステージ済みOfficeファイルを展開ツリーへ置き換える関数。"""
    converted = 0
    for path in staged_paths():
        if is_ooxml_office_path(path):
            if convert_office_to_tree(path):
                converted += 1
    if converted:
        log(f"ooxml hook pre-commit converted {converted} file(s)")
    return 0


def iter_ooxml_tree_dirs(root: Path) -> list[Path]:
    """作業ツリー内のOOXML展開ディレクトリを列挙する関数。"""
    results: list[Path] = []
    for path in sorted(root.rglob("*")):
        if ".git" in path.parts:
            continue
        if path.is_dir() and is_ooxml_tree_dir(path):
            results.append(path)
    return results


def restore_office_from_tree(tree_path: Path) -> bool:
    """OOXML展開ツリーをTAR形式経由でOfficeファイルへ復元する関数。"""
    office_path = tree_path_to_office_path(tree_path)
    zip_data = container_to_zip_data(tree_to_tar_data(tree_path))
    office_path.parent.mkdir(parents=True, exist_ok=True)
    office_path.write_bytes(zip_data)
    log(f"ooxml hook restored Office file: {tree_path} -> {office_path}")
    return True


def post_checkout() -> int:
    """post-checkout hookから呼ばれ、展開ツリーからOfficeファイルを復元する関数。"""
    restored = 0
    for tree_path in iter_ooxml_tree_dirs(Path(".")):
        if restore_office_from_tree(tree_path):
            restored += 1
    if restored:
        log(f"ooxml hook post-checkout restored {restored} file(s)")
    return 0


def hook_wrapper_text(command: str) -> str:
    """Git hookに配置する薄いshラッパーの内容を作る関数。"""
    return "\n".join(
        [
            "#!/bin/sh",
            f'python -B scripts/hooks/ooxml_file_hooks.py {command} "$@"',
            "",
        ]
    )


def install_hooks() -> int:
    """pre-commitとpost-checkoutのGit hookラッパーを.git/hooksへ配置する関数。"""
    hooks_dir_data = run_git(["rev-parse", "--git-path", "hooks"]).stdout
    hooks_dir = Path(hooks_dir_data.decode("utf-8", errors="replace").strip())
    hooks_dir.mkdir(parents=True, exist_ok=True)
    hooks = {
        "pre-commit": "pre-commit",
        "post-checkout": "post-checkout",
    }
    for hook_name, command in hooks.items():
        hook_path = hooks_dir / hook_name
        hook_path.write_text(hook_wrapper_text(command), encoding="utf-8", newline="\n")
        try:
            # ファイルに実行権限を付与
            hook_path.chmod(0o755)
        except OSError:
            pass
        log(f"ooxml hook installed: {hook_path}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    """コマンドライン引数を解釈するパーサーを作る関数。"""
    parser = argparse.ArgumentParser(description="Manage OOXML files as extracted Git trees with hooks.")
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("pre-commit", help="convert staged Office files to extracted OOXML trees")
    post_checkout_parser = subparsers.add_parser("post-checkout", help="restore Office files from extracted OOXML trees")
    post_checkout_parser.add_argument("hook_args", nargs="*")
    subparsers.add_parser("install", help="install pre-commit and post-checkout hooks")
    return parser


def main(argv: list[str] | None = None) -> int:
    """サブコマンドに応じてOOXML hook処理を実行する関数。"""
    args = build_parser().parse_args(argv)
    if args.command == "pre-commit":
        return pre_commit()
    if args.command == "post-checkout":
        return post_checkout()
    if args.command == "install":
        return install_hooks()
    raise ValueError(f"unknown command: {args.command}")


if __name__ == "__main__":
    raise SystemExit(main())
