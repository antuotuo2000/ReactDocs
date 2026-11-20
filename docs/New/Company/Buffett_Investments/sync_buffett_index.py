#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Idempotent SITE_RELATED_NAV synchronizer for Buffett_Investments_MUI_Suite.
Refreshes <!-- SITE_RELATED_NAV_START --> ... <!-- SITE_RELATED_NAV_END --> across all 48 company HTML files.
"""
import glob, os, re

HERE = os.path.dirname(os.path.abspath(__file__))
PATTERN = re.compile(r"<!-- SITE_RELATED_NAV_START -->.*?<!-- SITE_RELATED_NAV_END -->", re.DOTALL)

def main():
    files = sorted(f for f in os.listdir(HERE) if f.endswith(".html") and f != "index.html")
    for i, fname in enumerate(files):
        prev_f = files[(i - 1) % len(files)]
        next_f = files[(i + 1) % len(files)]
        block = (
            "<!-- SITE_RELATED_NAV_START -->\n"
            "      <ul class=\"nav-list\">\n"
            "        <li><a href=\"./index.html\">🏛️ 巴菲特投资版图总览门户</a></li>\n"
            "        <li><a href=\"../../index.html\">📚 数据中心 · Research &amp; Docs</a></li>\n"
            f"        <li><a href=\"./{prev_f}\">← {prev_f[:-5]}</a></li>\n"
            f"        <li><a href=\"./{next_f}\">→ {next_f[:-5]}</a></li>\n"
            "      </ul>\n"
            "      <!-- SITE_RELATED_NAV_END -->"
        )
        fpath = os.path.join(HERE, fname)
        with open(fpath, "r", encoding="utf-8") as fp:
            txt = fp.read()
        new_txt = PATTERN.sub(block, txt)
        if new_txt != txt:
            with open(fpath, "w", encoding="utf-8") as fp:
                fp.write(new_txt)
    print(f"Synced {len(files)} standalone Buffett company monograph pages.")

if __name__ == "__main__":
    main()
