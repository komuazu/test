# -*- coding: utf-8 -*-
"""配布用ZIPを作る

    python make_zip.py [出力先.zip]

各PC固有のファイル（config.json / memo.json / web/data.json）は入れない。
日本語のファイル名が開けない環境向けに、起動.bat と同じ内容の start.bat も同梱する。
"""
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKIP_DIRS = {'__pycache__', '.git'}
SKIP_FILES = {'config.json', 'memo.json', 'memo.json.tmp', 'data.json',
              'make_zip.py', '.gitignore'}


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / '営業有る無しWEBAPP.zip'
    files = []
    for p in sorted(HERE.rglob('*')):
        if not p.is_file():
            continue
        rel = p.relative_to(HERE)
        if any(part in SKIP_DIRS for part in rel.parts) or p.name in SKIP_FILES:
            continue
        if p.resolve() == out.resolve():
            continue
        files.append((p, str(rel).replace('\\', '/')))

    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for p, name in files:
            z.write(p, name)
        # 日本語ファイル名が開けない環境向けの別名
        z.writestr('start.bat', (HERE / '起動.bat').read_bytes())

    print('作成: %s (%s bytes)' % (out, format(out.stat().st_size, ',')))
    for _, name in files:
        print('   ', name)
    print('    start.bat')


if __name__ == '__main__':
    main()
