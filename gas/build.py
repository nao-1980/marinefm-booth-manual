#!/usr/bin/env python3
# index.html から Apps Script 用の gas/Index.html を作る
# 使い方: python3 gas/build.py
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
# PDFは GitHub Pages に置いたものを参照する
PDF_BASE = 'https://nao-1980.github.io/marinefm-booth-manual/'

html = (ROOT / 'index.html').read_text(encoding='utf-8')
# config.js は使わない（google.script.run で直接呼ぶ）
html = html.replace('<script src="config.js"></script>\n', '')
# viewport は Code.gs の addMetaTag で指定する
html = re.sub(r'<meta name="(viewport|claude:viewport-fit-sides)"[^>]*>\n', '', html)
# 相対パスのPDFを絶対URLにする
html, n = re.subn(r'href="([\w.-]+\.pdf)"', lambda m: f'href="{PDF_BASE}{m.group(1)}"', html)
(ROOT / 'gas' / 'Index.html').write_text(html, encoding='utf-8')
print(f'gas/Index.html を作成（PDFリンク {n} 件を変換）')
