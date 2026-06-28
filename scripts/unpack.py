# git checkout した時にOOXMLファイルを作業ツリーへ戻す互換用入口
from pathlib import Path

from ooxml_filter import smudge_stdin, unpack_file


def unpack(src, dst):
    """OOXML zipファイルを指定ディレクトリへ安全に展開する関数。"""
    unpack_file(Path(src), Path(dst))


def main():
    """引数がなければGit smudge filter、引数があればzip展開を実行する関数。"""
    import sys

    if len(sys.argv) == 1:
        smudge_stdin()
        return 0
    if len(sys.argv) == 3:
        unpack(sys.argv[1], sys.argv[2])
        return 0
    print("usage: python scripts/unpack.py [src-zip dst-dir]", file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
