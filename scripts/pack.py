# git add した時にOOXMLファイルを正規化zipへ変換する互換用入口
from pathlib import Path

from ooxml_filter import clean_stdin, pack_directory


def pack(src, dst):
    """展開済みOOXMLディレクトリを正規化してzipファイルへ再圧縮する関数。"""
    pack_directory(Path(src), Path(dst))


def main():
    """引数がなければGit clean filter、引数があればディレクトリ再圧縮を実行する関数。"""
    import sys

    if len(sys.argv) == 1:
        clean_stdin()
        return 0
    if len(sys.argv) == 3:
        pack(sys.argv[1], sys.argv[2])
        return 0
    print("usage: python scripts/pack.py [src-dir dst-zip]", file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
