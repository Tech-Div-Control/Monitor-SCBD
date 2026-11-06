#!/usr/bin/env python3
"""Tạo docs/index.html (bản chạy trên Netlify) từ apps-script/Index.html (nguồn duy nhất của giao diện)."""
import pathlib
root = pathlib.Path(__file__).resolve().parent.parent
h = (root / 'apps-script' / 'Index.html').read_text(encoding='utf8')
h = h.replace('<base target="_top">', '')
h = h.replace('<head>', '<head>\n<link rel="preconnect" href="https://script.google.com">\n'
              '<link rel="preconnect" href="https://script.googleusercontent.com" crossorigin>', 1)
j = h.rindex('<script>')
h = h[:j] + '<script src="config.js"></script>\n' + h[j:]
(root / 'docs' / 'index.html').write_text(h, encoding='utf8')
print('Đã tạo docs/index.html')
