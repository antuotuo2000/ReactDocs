#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sync_docs_index.py — 数据中心 自动化增量全文索引、详情页关联导航与主题/双语同步脚本 (零第三方依赖 · 自愈增强版)

功能特性：
1. 全目录递归检索：支持任意深度的公司/主题文件夹与多级子目录，新增或替换文件自动识别入库并按最新时间置顶。
2. 增量缓存加速：基于文件大小与修改时间 (.index_cache.json)，仅解析新增或修改过的文件，毫秒级完成增量同步。
3. 详情页全自动联动同步（每次运行脚本自动对齐全站详情页）：
   - 自动更新每份详情页侧边栏的「站点与关联导航」（生成准确的本地相对路径返回 docs/index.html，并自动列出同公司目录下的全部关联深度研报）。
   - 自动统一详情页与数据中心首页的「浅色 / 深色外观切换」（默认 MUI 浅色，共享 localStorage['mui_portal_theme_v4'] 并支持跨标签页实时同步）。
   - 自动统一详情页与数据中心首页右上角的「中文 | EN 双语切换按钮」（共享 localStorage['mui_portal_lang_v1'] 并支持跨标签页实时同步）。
4. 模板自愈修复：
   - 即使 docs/index.html 是从画布预览版 (Canvas) 或浏览器“另存为”保存下来的（缺失注释锚点、残留内联旧索引或丢失 search_index.js 引用），脚本也会自动清理旧内联缓存、恢复标准锚点并注入最新防缓存版本号 (?v=...)。
   - 若误将门户 index.html 放在外层根目录（与 docs/ 同级），脚本会自动识别并迁移合并至 docs/index.html，防止相对路径失效。
5. 多格式正文提取（纯 Python 3 标准库，兼容 Python 3.7–3.13+，无需 pip install）：
   - 网页与研报：.html / .htm（提取标题、H1-H4 章节大纲、核心摘要与 100% 纯文本正文）
   - 文本与数据：.md / .txt / .csv / .tsv / .json / .yaml / .py / .js / .sql
   - Office 文档：.docx / .xlsx / .pptx（标准库 zipfile + XML 解析提取正文与工作表名）
   - PDF 文档：自动检测系统中已有的 pypdf / pdftotext 提取正文，无依赖时回退元数据索引
6. 导航栏配置同步：自动读取 docs/nav_links.json 并注入 docs/index.html，同时兼容 --menu 与 --group 参数快捷添加链接。

用法示例：
  python3 sync_docs_index.py                       # 增量扫描新文件、更新首页索引并同步全站详情页导航/主题/双语
  python3 sync_docs_index.py --open                # 同步完成后自动在浏览器打开 docs/index.html
  python3 sync_docs_index.py --watch               # 持续监听文件夹，拖入或修改文件自动更新索引与关联导航
  python3 sync_docs_index.py --rebuild             # 忽略缓存，全量重新解析所有文件并刷新全部详情页
  python3 sync_docs_index.py --add-link "雪球" "https://xueqiu.com" --group "精选导航"
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import html as html_mod
from html.parser import HTMLParser
import json
import os
import stat
import zipfile
from pathlib import Path
import re
import shutil
import subprocess
import sys
import time
from urllib.parse import quote, unquote
import webbrowser
import xml.etree.ElementTree as ET
import zipfile

IGNORED_NAMES = {
    ".ds_store",
    ".nojekyll",
    "thumbs.db",
    "desktop.ini",
    "search_index.js",
    "components-snapshot.js",
    "components-version.js",
    "portal.css",
    "nav_links.json",
    ".index_cache.json",
    ".drive_metadata.json",
    ".stray_root_index_backup.html",
    "index.root-backup.html",
    "sync_docs_index.py",
    "build_snapshot.py",
    "serve.py",
    "双击更新索引.command",
    "双击更新索引.bat",
    "启动财报主页.command",
}

IGNORED_DIRS = {
    "__MACOSX",
    "node_modules",
    "venv",
    ".venv",
    "__pycache__",
}

TEXT_EXTENSIONS = {
    ".md", ".markdown", ".txt", ".csv", ".tsv", ".json",
    ".yaml", ".yml", ".xml", ".py", ".js", ".ts", ".jsx",
    ".tsx", ".sql", ".sh", ".css", ".log", ".ini", ".toml",
}

CATEGORY_KNOWN = {
    "Company": "公司深度",
    "SeniorExecutive": "高管人物",
    "CorporateCulture": "企业文化",
    "software": "软件生态",
    "hardware": "硬件产品",
    "patent": "核心专利",
    "PhysicalStores": "实体零售",
}

CATEGORY_KNOWN_EN = {
    "Company": "Deep Dive",
    "SeniorExecutive": "Executive",
    "CorporateCulture": "Culture",
    "software": "Software",
    "hardware": "Hardware",
    "patent": "Patent",
    "PhysicalStores": "Retail",
    "RootFile": "Document",
}

CATEGORY_SORT_ORDER = {
    "Company": 1,
    "CorporateCulture": 2,
    "SeniorExecutive": 3,
    "hardware": 4,
    "software": 5,
    "patent": 6,
    "PhysicalStores": 7,
}

STANDARD_DIMS = [
    "Company",
    "SeniorExecutive",
    "CorporateCulture",
    "hardware",
    "software",
    "patent",
    "PhysicalStores",
]

DEFAULT_ROOT_DRIVE_URL = "https://drive.google.com/drive/folders/1GEm-cxIC-tAvyiMlAs82b8QYzjUqhNBc"
SKIP_TAGS = {"script", "style", "svg", "noscript", "template", "header", "aside"}

INJECTED_BLOCK_PATTERNS = (
    r"<!--\s*SITE_RELATED_NAV_START\s*-->[\s\S]*?<!--\s*SITE_RELATED_NAV_END\s*-->",
    r"<!--\s*PORTAL_DETAIL_SYNC_CSS_START\s*-->[\s\S]*?<!--\s*PORTAL_DETAIL_SYNC_CSS_END\s*-->",
    r"<!--\s*PORTAL_DETAIL_TOPBAR_START\s*-->[\s\S]*?<!--\s*PORTAL_DETAIL_TOPBAR_END\s*-->",
    r"<!--\s*PORTAL_DETAIL_SYNC_JS_START\s*-->[\s\S]*?<!--\s*PORTAL_DETAIL_SYNC_JS_END\s*-->",
    r"<!--\s*PORTAL_MUI_OPT_CSS_START\s*-->[\s\S]*?<!--\s*PORTAL_MUI_OPT_CSS_END\s*-->",
    r"<!--\s*COMPANY_SUBFOLDER_HUB_START\s*-->[\s\S]*?<!--\s*COMPANY_SUBFOLDER_HUB_END\s*-->",
)



PORTAL_MUI_OPT_CSS_BLOCK = """<!-- PORTAL_MUI_OPT_CSS_START -->
  <script>
    (function () {
      try {
        if (window.localStorage.getItem('mui_portal_sidebar_collapsed_v1') === '1') {
          document.documentElement.setAttribute('data-sidebar-collapsed', '1');
        }
      } catch (e) {}
    })();
  </script>
  <style id="portalMuiOptStyle">
    /* Sidebar Height & Nav Optimization + Cross-Page Collapse */
    @media (min-width: 1024px) {
      .mui-sidebar {
        top: 72px !important;
        max-height: none !important;
        overflow-y: visible !important;
        padding-bottom: 24px !important;
        gap: 10px !important;
      }
      html[data-sidebar-collapsed="1"] .mui-sidebar {
        display: none !important;
      }
      html[data-sidebar-collapsed="1"] .layout-shell {
        grid-template-columns: minmax(0, 1fr) !important;
      }
    }
    .sidebar-card {
      padding: 12px 14px !important;
      flex-shrink: 0 !important;
    }
    .sidebar-subtitle {
      margin: 4px 0 8px !important;
      font-size: 0.77rem !important;
      line-height: 1.4 !important;
    }
    .sidebar-mini-kpis {
      gap: 4px !important;
      padding-top: 8px !important;
    }
    .mui-sidebar .mui-nav-list {
      gap: 2px !important;
    }
    .mui-sidebar .mui-nav-list:first-of-type {
      max-height: 175px;
      overflow-y: auto;
      scrollbar-width: thin;
      padding-right: 3px;
    }
    .mui-ext-nav-link {
      padding: 4px 8px !important;
      font-size: 0.77rem !important;
      line-height: 1.32 !important;
    }
    .mui-ext-nav-left > span:last-child {
      display: -webkit-box;
      -webkit-line-clamp: 2;
      -webkit-box-orient: vertical;
      overflow: hidden;
    }
    .mui-nav-link {
      padding: 5px 10px !important;
      font-size: 0.80rem !important;
    }
    main > section, .section-block, .mui-section { content-visibility: auto; contain-intrinsic-size: auto 520px; }
    /* Detail Page Tag/Chip Wrapping Fix (Sidebar, Mobile Drawer, Hero, Cards) */
    .mui-chip {
      box-sizing: border-box !important;
      display: inline-flex !important;
      flex-wrap: wrap !important;
      align-items: center !important;
      max-width: 100% !important;
      min-width: 0 !important;
      white-space: normal !important;
      height: auto !important;
      min-height: 24px;
      line-height: 1.35 !important;
      padding: 3px 10px !important;
      overflow-wrap: anywhere !important;
      word-break: break-word !important;
      text-align: left !important;
    }
    .appbar-left .mui-chip {
      white-space: nowrap !important;
      flex-wrap: nowrap !important;
      max-width: 190px !important;
      overflow: hidden !important;
      text-overflow: ellipsis !important;
      display: inline-block !important;
    }
    .sidebar-brand-top, .hero-overline, .mui-hero-badges, .kpi-top, .matrix-card-head {
      flex-wrap: wrap !important;
      gap: 6px !important;
      max-width: 100% !important;
      min-width: 0 !important;
    }
    @media (max-width: 1023px) {
      .mui-sidebar {
        width: min(300px, 86vw) !important;
        max-width: 86vw !important;
        box-sizing: border-box !important;
        overflow-x: hidden !important;
      }
      .mui-sidebar .sidebar-card {
        width: 100% !important;
        max-width: 100% !important;
        min-width: 0 !important;
        box-sizing: border-box !important;
        overflow-x: hidden !important;
      }
    }
    /* Company Subfolder Hub on Detail Pages */
    .co-hub-card {
      margin: 0 0 18px; padding: 14px 16px; border-radius: 14px;
      border: 1px solid var(--mui-divider, #e2e8f0);
      background: linear-gradient(rgba(25, 118, 210, 0.03) 1px, transparent 1px),
                  linear-gradient(90deg, rgba(25, 118, 210, 0.03) 1px, transparent 1px),
                  var(--mui-bg-paper, #fff);
      background-size: 20px 20px, 20px 20px, auto;
    }
    [data-theme="dark"] .co-hub-card, html.dark .co-hub-card {
      background: var(--mui-bg-paper, #0f172a);
      border-color: rgba(255,255,255,0.10);
    }
    .co-hub-head { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; margin-bottom: 10px; }
    .co-hub-heading { font-size: 0.90rem; color: var(--mui-text-primary, #0f172a); }
    [data-theme="dark"] .co-hub-heading, html.dark .co-hub-heading { color: #f1f5f9; }
    .co-hub-actions { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; margin-bottom: 10px; }
    .co-hub-btn {
      display: inline-flex; align-items: center; justify-content: center; gap: 5px;
      height: 34px; padding: 0 13px; border-radius: 999px; font-size: 0.80rem; font-weight: 700;
      border: 1px solid var(--mui-divider, #cbd5e1); background: var(--mui-bg-paper, #fff);
      color: var(--mui-text-primary, #0f172a); text-decoration: none !important; cursor: pointer;
      transition: all 0.14s ease;
    }
    [data-theme="dark"] .co-hub-btn, html.dark .co-hub-btn { background: rgba(15,23,42,0.8); color: #e2e8f0; border-color: rgba(255,255,255,0.14); }
    .co-hub-btn:hover { border-color: #1976d2; color: #1976d2; }
    .co-hub-btn-primary { background: #0073e6 !important; color: #fff !important; border-color: #0073e6 !important; }
    .co-hub-btn-primary:hover { background: #005bb5 !important; color: #fff !important; }
    .co-hub-pills {
      display: inline-flex; flex-wrap: wrap; gap: 3px; padding: 3px; border-radius: 10px;
      background: var(--mui-bg-subtle, #f1f5f9); border: 1px solid var(--mui-divider, #e2e8f0); margin-bottom: 10px;
    }
    [data-theme="dark"] .co-hub-pills, html.dark .co-hub-pills { background: rgba(15,23,42,0.65); border-color: rgba(255,255,255,0.10); }
    .co-hub-pill {
      padding: 5px 12px; border-radius: 8px; font-size: 0.77rem; font-weight: 700; cursor: pointer;
      border: none; background: transparent; color: var(--mui-text-secondary, #475569); transition: all 0.14s ease;
    }
    [data-theme="dark"] .co-hub-pill, html.dark .co-hub-pill { color: #94a3b8; }
    .co-hub-pill.active { background: #fff; color: #0073e6; box-shadow: 0 1px 3px rgba(15,23,42,0.10); }
    [data-theme="dark"] .co-hub-pill.active, html.dark .co-hub-pill.active { background: #1e293b; color: #60a5fa; }
    .co-hub-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 8px; }
    @media (max-width: 1120px) { .co-hub-grid { grid-template-columns: repeat(3, minmax(0, 1fr)); } }
    @media (max-width: 720px) { .co-hub-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); } }
    @media (max-width: 440px) { .co-hub-grid { grid-template-columns: minmax(0, 1fr); } }
    .co-hub-item {
      display: flex; align-items: center; gap: 8px; padding: 8px 10px; border-radius: 10px;
      border: 1px solid var(--mui-divider, #e2e8f0); background: var(--mui-bg-paper, #fff);
      text-decoration: none !important; min-width: 0; transition: all 0.14s ease;
    }
    [data-theme="dark"] .co-hub-item, html.dark .co-hub-item { background: rgba(15,23,42,0.7); border-color: rgba(255,255,255,0.10); }
    .co-hub-item:hover, .co-hub-item.active {
      border-color: #1976d2; box-shadow: 0 0 0 1px #1976d2;
    }
    .co-hub-ticker {
      flex-shrink: 0; min-width: 38px; padding: 4px 5px; border-radius: 6px; text-align: center;
      font-size: 0.68rem; font-weight: 800; background: rgba(25,118,210,0.12); color: #1976d2;
    }
    .co-hub-info { display: flex; flex-direction: column; min-width: 0; flex: 1; }
    .co-hub-title {
      font-size: 0.77rem; font-weight: 700; color: var(--mui-text-primary, #0f172a);
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
    }
    [data-theme="dark"] .co-hub-title, html.dark .co-hub-title { color: #e2e8f0; }
    .co-hub-meta {
      font-size: 0.68rem; color: var(--mui-text-muted, #64748b);
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
    }
    /* Mobile Horizontal Overflow Fix */
    @media (max-width: 720px) {
      html, body { overflow-x: hidden; max-width: 100vw; }
      .mui-chip, .mui-segmented button, .cite-link { white-space: normal !important; max-width: 100%; overflow-wrap: anywhere; }
      .mui-input-outlined { max-width: 100%; min-width: 0; }
    }
  </style>
  <script>
    (function () {
      var KEY = 'mui_portal_sidebar_collapsed_v1';
      function syncFromStorage() {
        try {
          var collapsed = window.localStorage.getItem(KEY) === '1';
          if (collapsed) document.documentElement.setAttribute('data-sidebar-collapsed', '1');
          else document.documentElement.removeAttribute('data-sidebar-collapsed');
        } catch (e) {}
      }
      document.addEventListener('click', function (ev) {
        var btn = ev.target && ev.target.closest ? ev.target.closest('#sidebarToggleBtn') : null;
        if (!btn) return;
        if (window.innerWidth >= 1024) {
          ev.preventDefault();
          ev.stopImmediatePropagation();
          var next = document.documentElement.getAttribute('data-sidebar-collapsed') !== '1';
          if (next) document.documentElement.setAttribute('data-sidebar-collapsed', '1');
          else document.documentElement.removeAttribute('data-sidebar-collapsed');
          try { window.localStorage.setItem(KEY, next ? '1' : '0'); } catch (e) {}
          window.dispatchEvent(new Event('resize'));
        }
      }, true);
      window.addEventListener('storage', function (ev) {
        if (ev.key === KEY) syncFromStorage();
      });
    })();
  </script>
<!-- PORTAL_MUI_OPT_CSS_END -->"""



def build_subfolder_matrix_data(docs_dir: str, files_by_rel: dict[str, dict]) -> list[dict]:
    """扫描每个公司目录下的首个子文件夹作为分类：有 index.html 则仅收录 index.html 并剪枝，否则递归扁平收录全部层级。"""
    matrix: list[dict] = []
    if not os.path.isdir(docs_dir):
        return matrix
    for comp in sorted(os.listdir(docs_dir)):
        cpath = os.path.join(docs_dir, comp)
        if not os.path.isdir(cpath) or comp.startswith("."):
            continue
        for sub in sorted(os.listdir(cpath)):
            spath = os.path.join(cpath, sub)
            if not os.path.isdir(spath) or sub.startswith("."):
                continue
            for dp, dns, fns in os.walk(spath):
                dns[:] = sorted(d for d in dns if not d.startswith("."))
                if "index.html" in fns:
                    rel_p = os.path.relpath(os.path.join(dp, "index.html"), docs_dir).replace(os.sep, "/")
                    meta = files_by_rel.get(rel_p, {})
                    sub_rel = os.path.relpath(dp, cpath).replace(os.sep, "/")
                    matrix.append({
                        "company": comp,
                        "companyLabel": meta.get("companyLabel", comp),
                        "ticker": meta.get("ticker", comp[:6].upper()),
                        "category": sub,
                        "subPath": sub_rel,
                        "relPath": rel_p,
                        "title": meta.get("title") or f"{comp} · {sub_rel}",
                        "titleEn": meta.get("titleEn") or f"{comp} · {sub_rel}",
                        "isIndex": True,
                    })
                    dns[:] = []
                else:
                    for fn in sorted(fns):
                        if not fn.lower().endswith(".html"):
                            continue
                        rel_p = os.path.relpath(os.path.join(dp, fn), docs_dir).replace(os.sep, "/")
                        meta = files_by_rel.get(rel_p, {})
                        sub_rel = os.path.relpath(dp, cpath).replace(os.sep, "/")
                        matrix.append({
                            "company": comp,
                            "companyLabel": meta.get("companyLabel", comp),
                            "ticker": meta.get("ticker", comp[:6].upper()),
                            "category": sub,
                            "subPath": sub_rel,
                            "relPath": rel_p,
                            "title": meta.get("title") or fn.replace(".html", ""),
                            "titleEn": meta.get("titleEn") or fn.replace(".html", ""),
                            "isIndex": False,
                        })
    return matrix

def ensure_index_preview_and_bilingual(idx_html: str) -> str:
    """确保首页 index.html 具备完整的英文 titleEn/summaryEn/headingsEn 切换与无报错速览弹窗。"""
    if "function getTitle(it)" not in idx_html:
        helpers = (
            "      function getTitle(it) { return (isEn() && it.titleEn) ? it.titleEn : it.title; }\n"
            "      function getSummary(it) { return (isEn() && it.summaryEn) ? it.summaryEn : (it.summary || ''); }\n"
            "      function getHeadings(it) { return (isEn() && it.headingsEn && it.headingsEn.length) ? it.headingsEn : (it.headings || []); }\n"
        )
        idx_html = idx_html.replace("      function refreshItemSearchFields(it) {", helpers + "\n      function refreshItemSearchFields(it) {", 1)
    idx_html = idx_html.replace("it._titleHay = [it.title, it.name,", "it._titleHay = [it.title, it.titleEn || '', it.name,")
    idx_html = idx_html.replace("it._headingHay = (it.headings || []).join(' ').toLowerCase();", "it._headingHay = (it.headings || []).concat(it.headingsEn || []).join(' ').toLowerCase();")
    idx_html = idx_html.replace("var heads = (it.headings || []).slice(0, 8);", "var allHeads = getHeadings(it); var heads = allHeads.slice(0, 8);")
    idx_html = idx_html.replace("highlight(it.title, tokens)", "highlight(getTitle(it), tokens)")
    idx_html = idx_html.replace("highlight(it.summary || '', tokens)", "highlight(getSummary(it), tokens)")
    idx_html = idx_html.replace("document.getElementById('readerTitle').textContent = it.title;", "document.getElementById('readerTitle').textContent = getTitle(it);")
    # Fix rHeads.length inside openReaderDrawer (prevent ReferenceError: allHeads is not defined)
    pos = idx_html.find("document.getElementById('readerBody').innerHTML =")
    if pos != -1:
        head, tail = idx_html[:pos], idx_html[pos:]
        tail = tail.replace("allHeads.length", "rHeads.length", 1)
        idx_html = head + tail
    # Restore new-tab opening for homepage Open buttons so index stays warm and navigation is instant
    idx_html = idx_html.replace('esc(href) + \'">\' + (isEn() ? \'Open\' : \'打开\')', 'esc(href) + \'" target="_blank" rel="noopener noreferrer">\' + (isEn() ? \'Open ↗\' : \'打开 ↗\')')
    idx_html = idx_html.replace('esc(href) + \'">\' + (isEn() ? \'Open Page\' : \'打开页面\')', 'esc(href) + \'" target="_blank" rel="noopener noreferrer">\' + (isEn() ? \'Open in New Tab ↗\' : \'新窗口打开文件 ↗\')')
    # Speed optimization: load 1.9MB search_index.js async, skip off-screen card layout, and defer FULLTEXT lowercasing to idle
    idx_html = re.sub(r'(<script\s+id="docsSearchIndexScript"[^>]*?)\s+defer(?=>)', r'\1 async', idx_html)
    if "contain-intrinsic-size: auto 260px" not in idx_html and "</style>" in idx_html:
        idx_html = idx_html.replace("</style>", "    .MuiCard { content-visibility: auto; contain-intrinsic-size: auto 260px; }\n  </style>", 1)
    idx_html = idx_html.replace(
        "contentItems.forEach(refreshItemSearchFields);",
        "var _runIdx = function () { contentItems.forEach(refreshItemSearchFields); if (state.query.trim()) renderAll(); }; if (window.requestIdleCallback) requestIdleCallback(_runIdx, { timeout: 600 }); else setTimeout(_runIdx, 120);",
        1,
    )
    # Clicking card/list title opens the webpage directly instead of Preview drawer
    idx_html = idx_html.replace('\'<div class="MuiCard-title" data-open-reader="\' + esc(it.id) + \'">\' + highlight(getTitle(it), tokens) + \'</div>\'', '\'<a class="MuiCard-title" href="\' + esc(href) + \'" target="_blank" rel="noopener noreferrer" style="display:block;text-decoration:none;color:inherit;">\' + highlight(getTitle(it), tokens) + \'</a>\'')
    idx_html = idx_html.replace('\'<div style="font-weight:700;cursor:pointer;color:var(--mui-text-primary);overflow-wrap:anywhere;" data-open-reader="\' + esc(it.id) + \'">\' + highlight(getTitle(it), tokens) + \'</div>\'', '\'<a href="\' + esc(href) + \'" target="_blank" rel="noopener noreferrer" style="display:block;font-weight:700;text-decoration:none;color:var(--mui-text-primary);overflow-wrap:anywhere;">\' + highlight(getTitle(it), tokens) + \'</a>\'')
    # Hide file relPath on homepage cards & list table
    idx_html = idx_html.replace('\'<span style="flex:1 1 100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;opacity:0.85;">\' + highlight(it.relPath, tokens) + \'</span>\' +', "")
    idx_html = idx_html.replace('\'<div class="mono" style="font-size:0.73rem;color:var(--mui-text-muted);margin-top:2px;overflow-wrap:anywhere;">\' + highlight(it.relPath, tokens) + \'</div>\' +', "")
    # Restore original homepage card & AppBar CSS (remove temporary header/card-banner overrides)
    idx_html = re.sub(r"\s*/\* Taller AppBar & Larger Header/Tab Fonts \*/[\s\S]*?(?=\s*</style>)", "", idx_html)
    idx_html = re.sub(r"\s*/\* Enhanced File Card Header \(\.MuiCard-banner\)[\s\S]*?(?=\s*</style>)", "", idx_html)

    # Keep homepage clean (strip any temporary homepage subfolder matrix panel)
    idx_html = re.sub(r"\s*/\* Company Subfolder Matrix Panel \*/[\s\S]*?(?=\s*</style>)", "", idx_html)
    idx_html = re.sub(r"\s*<!--\s*SUBFOLDER_MATRIX_PANEL_START\s*-->[\s\S]*?<!--\s*SUBFOLDER_MATRIX_PANEL_END\s*-->", "", idx_html)
    idx_html = re.sub(r"\s*<!--\s*SUBFOLDER_MATRIX_JS_START\s*-->[\s\S]*?<!--\s*SUBFOLDER_MATRIX_JS_END\s*-->", "", idx_html)
    return idx_html



def audit_detail_chip_wrapping(docs_dir: str) -> dict[str, int]:
    """每次打包/同步时检查所有详情页 .mui-chip 标签换行与移动端抽屉防溢出规则是否到位。"""
    scanned = 0
    fixed_inline = 0
    missing_opt = 0
    for dp, dns, fns in os.walk(docs_dir):
        dns[:] = [d for d in dns if not d.startswith(".")]
        for fn in fns:
            if not fn.lower().endswith(".html"):
                continue
            fp = os.path.join(dp, fn)
            if os.path.abspath(fp) == os.path.abspath(os.path.join(docs_dir, "index.html")):
                continue
            scanned += 1
            with open(fp, "r", encoding="utf-8", errors="ignore") as f:
                txt = f.read()
            if "mui-chip" in txt and "Detail Page Tag/Chip Wrapping Fix" not in txt:
                missing_opt += 1
            # Strip any inline white-space:nowrap on .mui-chip elements
            patched = re.sub(
                r'(<[^>]*class="[^"]*\bmui-chip\b[^"]*"[^>]*style="[^"]*?)white-space\s*:\s*nowrap\s*;?',
                r'\1white-space:normal;',
                txt,
                flags=re.IGNORECASE,
            )
            if patched != txt:
                fixed_inline += 1
                with open(fp, "w", encoding="utf-8") as f:
                    f.write(patched)
    return {"scanned": scanned, "fixed_inline": fixed_inline, "missing_opt": missing_opt}

def package_portal_zip(portal_root: Path, zip_path: Path) -> Path:
    """将整个 an/ 门户目录（含 docs/、sync_docs_index.py 及启动脚本）打包为可独立分发的 ZIP 归档。"""
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    stage_parent = portal_root.parent
    with zipfile.ZipFile(str(zip_path), "w", zipfile.ZIP_DEFLATED) as zf:
        for dp, dns, fns in os.walk(str(portal_root)):
            dns[:] = sorted(d for d in dns if d not in ("__MACOSX", "__pycache__", ".git"))
            for fn in sorted(fns):
                if fn.startswith("._") or fn == ".DS_Store" or fn.endswith(".tmp"):
                    continue
                fp = os.path.join(dp, fn)
                if os.path.abspath(fp) == os.path.abspath(str(zip_path)):
                    continue
                rel = os.path.relpath(fp, str(stage_parent)).replace(os.sep, "/")
                zi = zipfile.ZipInfo.from_file(fp, rel)
                zi.compress_type = zipfile.ZIP_DEFLATED
                zi.external_attr = ((stat.S_IFREG | 0o755) if fn.endswith((".command", ".py", ".sh")) else (stat.S_IFREG | 0o644)) << 16
                with open(fp, "rb") as f:
                    zf.writestr(zi, f.read())
    return zip_path

def strip_injected_detail_blocks(raw_html: str) -> str:
    """移除脚本注入的导航与顶栏块，保证正文摘要提取干净且重复执行幂等。"""
    cleaned = raw_html
    for pat in INJECTED_BLOCK_PATTERNS:
        cleaned = re.sub(pat, "", cleaned)
    return cleaned


class HTMLContentParser(HTMLParser):
    """零依赖容错 HTML 解析器：提取 <title>、H1-H4 标题大纲、Meta 描述与完整可见正文。"""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._skip_stack: list[str] = []
        self._in_title = False
        self._heading_tag: str | None = None
        self.title_parts: list[str] = []
        self.meta_desc: str = ""
        self.headings: list[str] = []
        self._cur_heading: list[str] = []
        self.text_chunks: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        t = tag.lower()
        if t in SKIP_TAGS:
            self._skip_stack.append(t)
            return
        if self._skip_stack:
            if t in ("main", "article", "section", "h1", "h2") and all(
                s not in ("script", "style") for s in self._skip_stack
            ):
                self._skip_stack.clear()
            else:
                return
        if t == "title":
            self._in_title = True
        elif t in ("h1", "h2", "h3", "h4"):
            self._heading_tag = t
            self._cur_heading = []
        elif t == "meta":
            attr_map = {(k or "").lower(): (v or "") for k, v in attrs}
            if attr_map.get("name", "").lower() == "description" and attr_map.get("content"):
                self.meta_desc = attr_map["content"].strip()

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        t = tag.lower()
        if t in SKIP_TAGS:
            return
        if not self._skip_stack and t == "meta":
            attr_map = {(k or "").lower(): (v or "") for k, v in attrs}
            if attr_map.get("name", "").lower() == "description" and attr_map.get("content"):
                self.meta_desc = attr_map["content"].strip()

    def handle_endtag(self, tag: str) -> None:
        t = tag.lower()
        if self._skip_stack:
            if t in self._skip_stack:
                while self._skip_stack:
                    popped = self._skip_stack.pop()
                    if popped == t:
                        break
            return
        if t == "title":
            self._in_title = False
        elif t == self._heading_tag:
            heading_text = re.sub(r"\s+", " ", "".join(self._cur_heading)).strip()
            if len(heading_text) >= 2 and heading_text not in self.headings:
                self.headings.append(heading_text)
            self._heading_tag = None

    def handle_data(self, data: str) -> None:
        if self._skip_stack:
            return
        if self._in_title:
            self.title_parts.append(data)
            return
        if self._heading_tag is not None:
            self._cur_heading.append(data)
        cleaned = data.strip()
        if cleaned:
            self.text_chunks.append(cleaned)



ZH_EN_PHRASES = [
    ("企业文化与组织架构深度研究报告", "Corporate Culture & Org Architecture Deep Dive"),
    ("企业文化与治理架构深度研究报告", "Corporate Culture & Governance Deep Dive"),
    ("企业文化与工程治理深度研究报告", "Corporate Culture & Engineering Governance Deep Dive"),
    ("企业文化与组织演进深度研究报告", "Corporate Culture & Organizational Evolution Deep Dive"),
    ("深度商业与产品战略研究报告", "Business & Product Strategy Deep Dive"),
    ("深度商业与领导力研究报告", "Business & Leadership Deep Dive"),
    ("深度商业与价值投资研究报告", "Business & Value Investing Deep Dive"),
    ("深度商业与资本配置研究报告", "Business & Capital Allocation Deep Dive"),
    ("深度商业与投资研究报告", "Business & Investment Deep Dive"),
    ("深度商业与战略研究报告", "Business & Strategy Deep Dive"),
    ("深度商业、财务与估值研究报告", "Business, Financials & Valuation Deep Dive"),
    ("深度投资人物与战略研究报告", "Investor & Strategy Deep Dive"),
    ("深度商业与领袖战略研究报告", "Business & Leadership Strategy Deep Dive"),
    ("深度商业与技术战略研究报告", "Business & Technology Strategy Deep Dive"),
    ("深度商业与全栈算力战略研究报告", "Business & Full-Stack Compute Strategy Deep Dive"),
    ("深度商业与平台战略研究报告", "Business & Platform Strategy Deep Dive"),
    ("深度商业与航天战略研究报告", "Business & Aerospace Strategy Deep Dive"),
    ("深度商业与转型战略研究报告", "Business & Transformation Strategy Deep Dive"),
    ("深度商业与算力架构研究报告", "Business & Compute Architecture Deep Dive"),
    ("深度商业与供应链战略研究报告", "Business & Supply Chain Strategy Deep Dive"),
    ("深度商业与工程战略研究报告", "Business & Engineering Strategy Deep Dive"),
    ("深度领导力与战略研究报告", "Leadership & Strategy Deep Dive"),
    ("深度持仓研究报告", "13F Portfolio Deep Dive"),
    ("深度研究报告", "Deep Dive Research Report"),
    ("深度分析报告", "Deep Dive Analysis"),
    ("企业文化深度研究报告", "Corporate Culture Deep Dive"),
    ("企业文化研报", "Corporate Culture Report"),
    ("执行摘要与核心指标", "Executive Summary & KPIs"),
    ("《股东手册》", "Owner's Manual "),
    ("所有权原则", "Owner-Related Principles"),
    ("头版测试", "Front-Page Test"),
    ("声誉铁律", "Reputation Ironclad"),
    ("美债堡垒", "T-Bill Fortress"),
    ("零成本浮存金", "Zero-Cost Float"),
    ("浮存金", "Float"),
    ("现金堡垒", "Cash Fortress"),
    ("净现金与美债储备", "Net Cash & T-Bills Reserve"),
    ("净现金", "Net Cash"),
    ("营业利润率", "Operating Margin"),
    ("营业利润", "Operating Earnings"),
    ("自由现金流", "Free Cash Flow"),
    ("毛利率", "Gross Margin"),
    ("净利率", "Net Margin"),
    ("同比增长率", "YoY Growth"),
    ("同比增速", "YoY Growth"),
    ("市场份额", "Market Share"),
    ("总市值", "Market Cap"),
    ("市值", "Market Cap"),
    ("股价", "Stock Price"),
    ("市盈率", "P/E"),
    ("市销率", "P/S"),
    ("资本配置", "Capital Allocation"),
    ("资本开支", "CapEx"),
    ("股票回购", "Share Buybacks"),
    ("股息增长", "Dividend Growth"),
    ("护城河", "Moat"),
    ("雷达图", "Radar"),
    ("横向对比", "Peer Comparison"),
    ("组织架构", "Org Architecture"),
    ("企业文化", "Corporate Culture"),
    ("治理架构", "Governance Architecture"),
    ("领导力哲学", "Leadership Philosophy"),
    ("商业版图", "Business Architecture"),
    ("产品矩阵", "Product Matrix"),
    ("收入架构", "Revenue Architecture"),
    ("业务与收入架构", "Business & Revenue Architecture"),
    ("行业地位与四重核心护城河", "Industry Position & Four Core Moats"),
    ("四重核心护城河", "Four Core Moats"),
    ("四重文化护城河", "Four Cultural Moats"),
    ("核心护城河", "Core Moats"),
    ("核心竞争力", "Core Competitiveness"),
    ("五大核心", "Five Core"),
    ("四大核心", "Four Core"),
    ("三大核心", "Three Core"),
    ("前十大重仓", "Top 10 Holdings"),
    ("前五大公司", "Top 5 Companies"),
    ("组合集中度", "Portfolio Concentration"),
    ("完整持仓矩阵", "Full Holdings Matrix"),
    ("持仓全景", "Portfolio Overview"),
    ("调仓脉络", "Rebalancing Moves"),
    ("阿贝尔时代", "Greg Abel Era"),
    ("后巴菲特时代", "Post-Buffett Era"),
    ("巴菲特", "Buffett"),
    ("芒格", "Munger"),
    ("阿贝尔", "Greg Abel"),
    ("伯克希尔·哈撒韦", "Berkshire Hathaway"),
    ("伯克希尔哈撒韦", "Berkshire Hathaway"),
    ("伯克希尔", "Berkshire"),
    ("苹果公司", "Apple Inc."),
    ("苹果", "Apple"),
    ("奥驰亚集团", "Altria Group"),
    ("奥驰亚", "Altria"),
    ("亚马逊", "Amazon"),
    ("谷歌", "Google"),
    ("英伟达", "NVIDIA"),
    ("特斯拉", "Tesla"),
    ("奈飞", "Netflix"),
    ("拼多多控股", "PDD Holdings"),
    ("拼多多", "PDD"),
    ("段永平", "Duan Yongping"),
    ("李录", "Li Lu"),
    ("喜马拉雅资本", "Himalaya Capital"),
    ("格雷格·阿贝尔", "Greg Abel"),
    ("史蒂夫·乔布斯", "Steve Jobs"),
    ("蒂姆·库克", "Tim Cook"),
    ("黄仁勋", "Jensen Huang"),
    ("埃隆·马斯克", "Elon Musk"),
    ("马克·扎克伯格", "Mark Zuckerberg"),
    ("桑达尔·皮查伊", "Sundar Pichai"),
    ("查克·罗宾斯", "Chuck Robbins"),
    ("萨尔·曼库索", "Sal Mancuso"),
    ("博通", "Broadcom"),
    ("新思科技", "Synopsys"),
    ("闪迪", "SanDisk"),
    ("西部数据", "Western Digital"),
    ("安谋控股", "Arm Holdings"),
    ("火箭实验室", "Rocket Lab"),
    ("莱纳房产", "Lennar Corp"),
    ("昂跑", "On Holding"),
    ("罗布乐思", "Roblox"),
    ("资生堂集团", "Shiseido"),
    ("腾讯音乐", "Tencent Music"),
    ("华特迪士尼", "Walt Disney"),
    ("思科系统", "Cisco Systems"),
    ("卡骆驰", "Crocs"),
    ("优倍快", "Ubiquiti"),
    ("纳斯达克100ETF", "Nasdaq-100 ETF"),
    ("标普500ETF", "S&P 500 ETF"),
    ("标普500", "S&P 500"),
    ("无创健康传感与生物光学监测技术", "Non-Invasive Health Sensing & Bio-Optical Tech"),
    ("第一性原理", "First Principles"),
    ("人均创收", "Revenue per Employee"),
    ("员工", "Employees"),
    ("总部", "HQ"),
    ("极端分权", "Extreme Decentralization"),
    ("分权联邦", "Decentralized Federation"),
    ("复利机器", "Compounding Machine"),
    ("流媒体", "Streaming"),
    ("全栈", "Full-Stack"),
    ("算力", "Compute"),
    ("人工智能", "AI"),
    ("数据中心", "Data Center"),
    ("半导体", "Semiconductor"),
    ("供应链", "Supply Chain"),
    ("商业模式", "Business Model"),
    ("估值", "Valuation"),
    ("营收", "Revenue"),
    ("净利润", "Net Income"),
    ("净利", "Net Income"),
    ("组合", "Portfolio"),
    ("现金", "Cash"),
    ("控股", "Holding"),
    ("年化", "Annualized"),
    ("研究报告", "Research Report"),
    ("研报", "Report"),
    ("深度", "Deep Dive"),
    ("全景", "Overview"),
    ("对比", "Comparison"),
    ("演进", "Trajectory"),
    ("拆解", "Breakdown"),
    ("分析", "Analysis"),
    ("报告", "Report"),
    ("与", " & "),
    ("和", " & "),
    ("、", ", "),
    ("：", ": "),
    ("（", " ("),
    ("）", ")"),
    ("—", " - "),
    ("万亿", "T"),
    ("十亿", "B"),
    ("亿", "00M"),
    ("万", "0k"),
    ("人", " People"),
    ("条", ""),
    ("项", ""),
    ("只", ""),
    ("家", ""),
    ("年", "Y"),
    ("月", "M"),
]

def zh_to_en_portal_text(text: str) -> str:
    if not text or not re.search(r"[\u4e00-\u9fff]", text):
        return text
    out = text
    # Convert Chinese numeric units (万亿 -> T, 亿 -> B/M, 万 -> k) cleanly
    def _conv_tril(m):
        return f"${float(m.group(1).replace(',', '')):.2f}T " if m.group(0).startswith("$") else f"{float(m.group(1).replace(',', '')):.2f}T "
    out = re.sub(r"(\$?)([\d,]+(?:\.\d+)?)\s*万亿", lambda m: f"{m.group(1)}{float(m.group(2).replace(',', '')):.2f}T ", out)
    out = re.sub(r"(\$?)([\d,]+(?:\.\d+)?)\s*亿", lambda m: f"{m.group(1)}{float(m.group(2).replace(',', ''))/10:.2f}B " if float(m.group(2).replace(',', '')) >= 10 else f"{m.group(1)}{float(m.group(2).replace(',', ''))*100:.0f}M ", out)
    out = re.sub(r"([\d,]+(?:\.\d+)?)\s*万", lambda m: f"{float(m.group(1).replace(',', ''))*10:.1f}k ", out)
    for zh, en in ZH_EN_PHRASES:
        out = out.replace(zh, en)
    out = out.replace("。", ". ").replace("，", ", ").replace("；", "; ")
    out = re.sub(r"[\u4e00-\u9fff]+", " ", out)
    out = re.sub(r"[《》「」『』【】]", "", out)
    # Deduplicate consecutive repeated English words/phrases (e.g. "Berkshire Hathaway Berkshire Hathaway")
    out = re.sub(r"\b([A-Za-z0-9.&'\-]+(?:\s+[A-Za-z0-9.&'\-]+){0,2})\s+\1\b", r"\1", out, flags=re.I)
    out = re.sub(r"\s+([.,;:)])", r"\1", out)
    out = re.sub(r"\(\s*\)", "", out)
    out = re.sub(r"\s{2,}", " ", out).strip(" ,:;-·/")
    return out or text


def stable_id(prefix: str, key: str) -> str:
    digest = hashlib.sha1(key.encode("utf-8")).hexdigest()[:16]
    return f"{prefix}_{digest}"


def get_mtime_ns(st: os.stat_result) -> int:
    return getattr(st, "st_mtime_ns", int(st.st_mtime * 1_000_000_000))


def read_text_with_fallback(path: str) -> str:
    for enc in ("utf-8", "utf-8-sig", "gb18030", "latin-1"):
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
        except OSError:
            return ""
    return ""


def is_portal_html(path: str) -> bool:
    if not os.path.isfile(path):
        return False
    head_chunk = read_text_with_fallback(path)
    return (
        "DOCS_CATALOG" in head_chunk
        or "docsCatalogData" in head_chunk
        or "globalSearchInput" in head_chunk
    )


def extract_docx_content(path: str) -> tuple[str, list[str], str]:
    headings: list[str] = []
    paragraphs: list[str] = []
    try:
        with zipfile.ZipFile(path, "r") as zf:
            if "word/document.xml" not in zf.namelist():
                return "", [], ""
            xml_bytes = zf.read("word/document.xml")
        root = ET.fromstring(xml_bytes)
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        for p in root.findall(".//w:p", ns):
            texts = [t.text for t in p.findall(".//w:t", ns) if t.text]
            line = re.sub(r"\s+", " ", "".join(texts)).strip()
            if not line:
                continue
            paragraphs.append(line)
            p_style = p.find(".//w:pStyle", ns)
            if p_style is not None:
                val = p_style.attrib.get(f"{{{ns['w']}}}val", "")
                if "Heading" in val or "标题" in val or val in ("1", "2", "3"):
                    if line not in headings:
                        headings.append(line)
    except Exception:
        return "", [], ""
    title = headings[0] if headings else (paragraphs[0][:80] if paragraphs else "")
    return title, headings[:20], " ".join(paragraphs)


def extract_xlsx_content(path: str) -> tuple[str, list[str], str]:
    headings: list[str] = []
    strings: list[str] = []
    try:
        with zipfile.ZipFile(path, "r") as zf:
            if "xl/workbook.xml" in zf.namelist():
                wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
                for elem in wb_root.iter():
                    if elem.tag.endswith("sheet") and elem.attrib.get("name"):
                        headings.append(f"工作表: {elem.attrib['name']}")
            if "xl/sharedStrings.xml" in zf.namelist():
                ss_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
                for elem in ss_root.iter():
                    if elem.tag.endswith("t") and elem.text and elem.text.strip():
                        strings.append(elem.text.strip())
    except Exception:
        return "", [], ""
    return "", headings[:20], " ".join(headings + strings)


def extract_pptx_content(path: str) -> tuple[str, list[str], str]:
    headings: list[str] = []
    slides_text: list[str] = []
    try:
        with zipfile.ZipFile(path, "r") as zf:
            slide_files = sorted(
                n for n in zf.namelist()
                if n.startswith("ppt/slides/slide") and n.endswith(".xml")
            )
            for idx, sf in enumerate(slide_files, 1):
                s_root = ET.fromstring(zf.read(sf))
                parts = [
                    elem.text.strip()
                    for elem in s_root.iter()
                    if elem.tag.endswith("t") and elem.text and elem.text.strip()
                ]
                if parts:
                    headings.append(f"Slide {idx}: {parts[0][:80]}")
                    slides_text.append(" ".join(parts))
    except Exception:
        return "", [], ""
    title = headings[0].replace("Slide 1: ", "") if headings else ""
    return title, headings[:20], " ".join(slides_text)


def extract_pdf_content(path: str) -> tuple[str, list[str], str]:
    for mod_name in ("pypdf", "PyPDF2"):
        try:
            mod = __import__(mod_name)
            reader = mod.PdfReader(path)
            pages_txt: list[str] = []
            headings: list[str] = []
            meta_title = ""
            if getattr(reader, "metadata", None) and getattr(reader.metadata, "title", None):
                meta_title = str(reader.metadata.title).strip()
            for i, page in enumerate(reader.pages):
                txt = (page.extract_text() or "").strip()
                if txt:
                    pages_txt.append(re.sub(r"\s+", " ", txt))
                    first_line = txt.splitlines()[0].strip()[:80]
                    if first_line and len(headings) < 18:
                        headings.append(f"P{i + 1}: {first_line}")
            return meta_title, headings, " ".join(pages_txt)
        except Exception:
            continue

    pdftotext_bin = shutil.which("pdftotext")
    if pdftotext_bin:
        try:
            res = subprocess.run(
                [pdftotext_bin, "-q", path, "-"],
                capture_output=True,
                text=True,
                timeout=20,
                check=False,
            )
            if res.returncode == 0 and res.stdout.strip():
                lines = [ln.strip() for ln in res.stdout.splitlines() if ln.strip()]
                headings = [ln[:80] for ln in lines[:12] if len(ln) >= 3]
                body = re.sub(r"\s+", " ", " ".join(lines))
                return (lines[0][:90] if lines else ""), headings, body
        except Exception:
            pass
    return "", [], ""


def extract_file_payload(full_path: str, file_name: str, max_chars: int) -> dict[str, object]:
    """根据扩展名提取文件标题、章节大纲、核心摘要与全文检索文本。"""
    stem, ext = os.path.splitext(file_name)
    ext_low = ext.lower()
    title = stem
    headings: list[str] = []
    body = ""
    meta_desc = ""

    if ext_low in (".html", ".htm"):
        raw = strip_injected_detail_blocks(read_text_with_fallback(full_path))
        parser = HTMLContentParser()
        try:
            parser.feed(raw)
            parser.close()
        except Exception:
            pass
        parsed_title = re.sub(r"\s+", " ", "".join(parser.title_parts)).strip()
        parsed_title = re.sub(r"\s*\|\s*(?:数据中心|Research Hub).*$", "", parsed_title).strip()
        if parsed_title:
            title = parsed_title
        elif parser.headings:
            title = parser.headings[0]
        headings = parser.headings[:16]
        meta_desc = parser.meta_desc
        body = re.sub(r"\s+", " ", " ".join(parser.text_chunks)).strip()

    elif ext_low in (".md", ".markdown"):
        raw = read_text_with_fallback(full_path)
        for line in raw.splitlines():
            m = re.match(r"^(#{1,4})\s+(.+)$", line.strip())
            if m:
                h_txt = re.sub(r"[*_`#]+", "", m.group(2)).strip()
                if h_txt and h_txt not in headings:
                    headings.append(h_txt)
        if headings:
            title = headings[0]
        clean_md = re.sub(r"```[\s\S]*?```", " ", raw)
        clean_md = re.sub(r"[#>*_`~\[\]()!-]+", " ", clean_md)
        body = re.sub(r"\s+", " ", clean_md).strip()
        headings = headings[:16]

    elif ext_low in TEXT_EXTENSIONS:
        raw = read_text_with_fallback(full_path)
        if ext_low == ".json":
            try:
                obj = json.loads(raw)
                if isinstance(obj, dict):
                    headings = [f"字段: {k}" for k in list(obj.keys())[:14]]
            except Exception:
                pass
        else:
            lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
            if lines and ext_low in (".csv", ".tsv"):
                headings = [f"表头: {lines[0][:120]}"]
            elif lines:
                headings = [ln[:90] for ln in lines[:8]]
        body = re.sub(r"\s+", " ", raw).strip()

    elif ext_low == ".docx":
        t_doc, h_doc, b_doc = extract_docx_content(full_path)
        if t_doc:
            title = t_doc
        headings, body = h_doc, b_doc

    elif ext_low == ".xlsx":
        t_xls, h_xls, b_xls = extract_xlsx_content(full_path)
        if t_xls:
            title = t_xls
        headings, body = h_xls, b_xls

    elif ext_low == ".pptx":
        t_ppt, h_ppt, b_ppt = extract_pptx_content(full_path)
        if t_ppt:
            title = t_ppt
        headings, body = h_ppt, b_ppt

    elif ext_low == ".pdf":
        t_pdf, h_pdf, b_pdf = extract_pdf_content(full_path)
        if t_pdf:
            title = t_pdf
        headings, body = h_pdf, b_pdf

    if max_chars > 0 and len(body) > max_chars:
        body = body[:max_chars]

    summary_source = meta_desc if meta_desc else body
    if summary_source.startswith(title):
        summary_source = summary_source[len(title):].strip(" ：:·-—|")
    summary_source = re.sub(r"^html_code\s*=\s*r?'''\s*", "", summary_source)
    summary_source = re.sub(r"^已(?:成功)?导出独立\s*(?:HTML|网页)\s*(?:报告|研报)?文件[：:\s]*\S*\s*", "", summary_source)
    summary_source = re.sub(r"\b(导出HTML|导出\.html|下载\s*\.html)\b\s*", "", summary_source)
    summary = summary_source[:190].strip()
    if len(summary_source) > 190:
        summary += "…"
    if not summary:
        ext_badge = ext_low.lstrip(".").upper() or "FILE"
        summary = f"{ext_badge} 文件 · 位于 {file_name}，可通过路径、文件名与目录分类直接检索打开。"

    file_type = ext_low.lstrip(".").upper() if ext_low else "FILE"

    return {
        "title": title,
        "headings": headings,
        "headingsEn": [zh_to_en_portal_text(h) for h in headings],
        "titleEn": zh_to_en_portal_text(title),
        "summary": summary,
        "summaryEn": zh_to_en_portal_text(summary),
        "fileType": file_type,
        "wordCount": len(body),
        "fullText": body,
    }


def infer_category(parts: list[str], ext_low: str) -> tuple[str, str]:
    if len(parts) == 1:
        if parts[0].lower() == "index.html":
            return "Portal", "Root"
        return "RootFile", "Root"

    sub_folder = parts[1] if len(parts) >= 3 else "Company"
    if sub_folder in CATEGORY_KNOWN:
        return sub_folder, sub_folder

    low = sub_folder.lower()
    if any(k in low for k in ("exec", "ceo", "founder", "leader", "高管", "人物")):
        return "SeniorExecutive", sub_folder
    if any(k in low for k in ("culture", "org", "文化", "组织")):
        return "CorporateCulture", sub_folder
    if any(k in low for k in ("soft", "app", "cloud", "ai", "软件")):
        return "software", sub_folder
    if any(k in low for k in ("hard", "chip", "device", "硬件", "芯片")):
        return "hardware", sub_folder
    if any(k in low for k in ("patent", "ip", "专利")):
        return "patent", sub_folder
    if any(k in low for k in ("store", "retail", "门店", "零售")):
        return "PhysicalStores", sub_folder
    if len(parts) == 2:
        if ext_low in (".html", ".htm", ".pdf", ".docx", ".md"):
            return "hardware" if "iphone" in parts[1].lower() or "mac" in parts[1].lower() else "Company", "Direct"
    return sub_folder, sub_folder


def heal_stray_root_index(parent_dir: str, docs_dir: str) -> None:
    """若用户误把门户 index.html 保存在外层目录 (与 docs/ 同级)，自动迁移合并至 docs/index.html。"""
    if os.path.abspath(parent_dir) == os.path.abspath(docs_dir):
        return
    stray_index = os.path.join(parent_dir, "index.html")
    docs_index = os.path.join(docs_dir, "index.html")
    if not os.path.isfile(stray_index) or not is_portal_html(stray_index):
        return

    backup_path = os.path.join(parent_dir, ".stray_root_index_backup.html")
    try:
        if not os.path.isfile(docs_index):
            shutil.move(stray_index, docs_index)
            print(f"⚠️  检测到门户首页被放在外层目录，已自动移入标准位置：{docs_index}")
        else:
            stray_mtime = os.path.getmtime(stray_index)
            docs_mtime = os.path.getmtime(docs_index)
            if stray_mtime > docs_mtime + 1.0:
                shutil.copy2(docs_index, backup_path)
                shutil.move(stray_index, docs_index)
                print(f"⚠️  检测到外层目录存在更新的 index.html，已自动合并至 {docs_index}（原文件已备份）")
            else:
                shutil.move(stray_index, backup_path)
                print(
                    f"⚠️  已清理外层目录重复的旧版 index.html（备份为 {os.path.basename(backup_path)}），"
                    f"请统一打开：{docs_index}"
                )
    except OSError:
        pass


def locate_docs_dir(cli_dir: str | None) -> str:
    """智能识别 docs 目录：优先识别含子目录的 docs/，兼容脚本置于仓库根目录或 docs/ 目录内。"""
    if cli_dir:
        target = os.path.abspath(cli_dir)
        sub_docs = os.path.join(target, "docs")
        if os.path.isdir(sub_docs):
            heal_stray_root_index(target, sub_docs)
            if os.path.isfile(os.path.join(sub_docs, "index.html")):
                return sub_docs
        return target

    script_dir = os.path.dirname(os.path.abspath(__file__))
    cwd = os.getcwd()

    for base in (script_dir, cwd):
        sub_docs = os.path.join(base, "docs")
        if os.path.isdir(sub_docs):
            heal_stray_root_index(base, sub_docs)
            if os.path.isfile(os.path.join(sub_docs, "index.html")):
                return sub_docs

    for cand in (script_dir, cwd):
        if os.path.isfile(os.path.join(cand, "index.html")):
            return cand

    return os.path.join(script_dir, "docs") if os.path.isdir(os.path.join(script_dir, "docs")) else script_dir


def load_json_safe(path: str, default: object) -> object:
    if not os.path.isfile(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def write_json_atomic(path: str, data: object, indent: int | None = 2) -> None:
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        if indent is None:
            json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
        else:
            json.dump(data, f, ensure_ascii=False, indent=indent)
    os.replace(tmp, path)


def encode_relative_url(from_dir: str, to_file: str) -> str:
    """生成经过 URL 编码的安全相对路径（兼容文件名或目录中的空格与特殊字符）。"""
    rel = os.path.relpath(to_file, from_dir).replace(os.sep, "/")
    parts = [quote(seg, safe="._-~()") for seg in rel.split("/")]
    joined = "/".join(parts)
    if not joined.startswith("."):
        joined = "./" + joined
    return joined


def build_concise_nav_labels(rec: dict) -> tuple[str, str]:
    """为详情页侧边栏『站点与关联导航』生成简洁清晰的中英文双语标题。"""
    cat = str(rec.get("category") or "Company")
    cat_zh = CATEGORY_KNOWN.get(cat, cat)
    cat_en = CATEGORY_KNOWN_EN.get(cat, cat)
    raw_title = str(rec.get("title") or rec.get("name") or "").strip()
    stem = os.path.splitext(str(rec.get("name") or ""))[0]

    # 压缩过长标题
    short_zh = re.split(r"[：:|—]", raw_title)[0].strip()
    if len(short_zh) > 22:
        short_zh = short_zh[:21].rstrip() + "…"
    if not short_zh:
        short_zh = stem

    # 提取英文可读名称
    ascii_matches = re.findall(r"[A-Za-z0-9][A-Za-z0-9\s.&'/-]{1,34}[A-Za-z0-9)]", raw_title)
    if ascii_matches:
        best_en = max(ascii_matches, key=len).strip()
    else:
        best_en = re.sub(r"[_-]+", " ", stem).strip()
    best_en = re.sub(r"\b(Analysis|Report|Deep Dive)\b", "", best_en, flags=re.IGNORECASE).strip(" ·-—")
    if not best_en:
        best_en = str(rec.get("company") or stem)
    if len(best_en) > 24:
        best_en = best_en[:23].rstrip() + "…"

    return f"[{cat_zh}] {short_zh}", f"[{cat_en}] {best_en}"


# ==============================================================================
# 详情页注入模板：与数据中心首页 100% 一致的浅色/深色主题 + 右上角中英文切换 + 关联导航
# ==============================================================================
DETAIL_SYNC_CSS_TEMPLATE = """<!-- PORTAL_DETAIL_SYNC_CSS_START -->
  <script>
    (function () {
      try {
        var rawT = window.localStorage.getItem('mui_portal_theme_v4');
        var t = rawT ? String(rawT).replace(/^"|"$/g, '') : 'light';
        var isDark = (t === 'dark');
        document.documentElement.classList.toggle('dark', isDark);
        document.documentElement.setAttribute('data-mui-theme', isDark ? 'dark' : 'light');
      } catch (e) {}
    })();
  </script>
  <style id="portalDetailSyncStyle">
    /* MUI Homepage-Consistent Light & Dark Surfaces for Classic Detail Pages */
    html:not(.dark) body {
      background-color: #F8FAFC !important;
      color: #0F172A !important;
    }
    html.dark body {
      background-color: #090D16 !important;
      color: #F8FAFC !important;
    }
    html:not(.dark) aside#sidebar {
      background-color: #FFFFFF !important;
      color: #334155 !important;
      border-right: 1px solid #E2E8F0 !important;
      box-shadow: 0 2px 10px rgba(15, 23, 42, 0.03) !important;
    }
    html.dark aside#sidebar {
      background-color: #0F1623 !important;
      color: #CBD5E1 !important;
      border-right: 1px solid rgba(255, 255, 255, 0.09) !important;
    }
    html:not(.dark) aside#sidebar h2,
    html:not(.dark) aside#sidebar .text-white {
      color: #0F172A !important;
    }
    html:not(.dark) aside#sidebar .text-slate-300,
    html:not(.dark) aside#sidebar .text-slate-400,
    html:not(.dark) aside#sidebar .text-slate-500 {
      color: #475569 !important;
    }
    html:not(.dark) aside#sidebar .border-slate-800,
    html:not(.dark) aside#sidebar .border-slate-800\\/80,
    html:not(.dark) aside#sidebar .border-slate-800\\/70,
    html:not(.dark) aside#sidebar .border-slate-700 {
      border-color: #E2E8F0 !important;
    }
    html:not(.dark) aside#sidebar .bg-navy-950\\/70,
    html:not(.dark) aside#sidebar .bg-navy-800 {
      background-color: #F1F5F9 !important;
      color: #0F172A !important;
    }
    html:not(.dark) aside#sidebar #desktop-nav .nav-link:hover {
      background-color: #F1F5F9 !important;
      color: #0073E6 !important;
    }
    html:not(.dark) header#mobile-header {
      background-color: rgba(255, 255, 255, 0.92) !important;
      color: #0F172A !important;
      border-bottom: 1px solid #E2E8F0 !important;
      backdrop-filter: blur(12px);
    }
    html:not(.dark) header#mobile-header .text-white {
      color: #0F172A !important;
    }

    /* Sticky Homepage-Consistent Top AppBar */
    .PortalDetailAppBar {
      position: sticky;
      top: 0;
      z-index: 35;
      height: 56px;
      padding: 0 20px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      background: rgba(255, 255, 255, 0.88);
      backdrop-filter: blur(14px);
      -webkit-backdrop-filter: blur(14px);
      border-bottom: 1px solid #E2E8F0;
    }
    html.dark .PortalDetailAppBar {
      background: rgba(9, 13, 22, 0.86);
      border-bottom-color: rgba(255, 255, 255, 0.09);
    }
    .PortalDetailAppBar-left {
      display: flex;
      align-items: center;
      gap: 10px;
      min-width: 0;
    }
    .PortalBackLink {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 5px 11px;
      border-radius: 9px;
      font-size: 12.5px;
      font-weight: 700;
      color: #0073E6;
      background: rgba(0, 115, 230, 0.08);
      border: 1px solid rgba(0, 115, 230, 0.24);
      text-decoration: none;
      white-space: nowrap;
      transition: all 0.15s ease;
    }
    .PortalBackLink:hover {
      background: rgba(0, 115, 230, 0.14);
    }
    html.dark .PortalBackLink {
      color: #3399FF;
      background: rgba(51, 153, 255, 0.14);
      border-color: rgba(51, 153, 255, 0.32);
    }
    .PortalCrumbMeta {
      font-size: 12px;
      font-weight: 600;
      color: #64748B;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    html.dark .PortalCrumbMeta {
      color: #94A3B8;
    }
    .PortalDetailAppBar-right {
      display: flex;
      align-items: center;
      gap: 8px;
      flex-shrink: 0;
    }
    .PortalLangSwitch {
      display: inline-flex;
      align-items: center;
      height: 36px;
      padding: 2px;
      border-radius: 10px;
      border: 1px solid #E2E8F0;
      background: #F1F5F9;
      gap: 2px;
    }
    html.dark .PortalLangSwitch {
      border-color: rgba(255, 255, 255, 0.10);
      background: rgba(19, 29, 47, 0.85);
    }
    .PortalLangBtn {
      height: 30px;
      padding: 0 10px;
      border: none;
      border-radius: 7px;
      background: transparent;
      color: #475569;
      font-size: 12px;
      font-weight: 700;
      cursor: pointer;
      transition: all 0.15s ease;
    }
    html.dark .PortalLangBtn {
      color: #94A3B8;
    }
    .PortalLangBtn.active {
      background: #FFFFFF;
      color: #0073E6;
      box-shadow: 0 1px 3px rgba(15, 23, 42, 0.10);
    }
    html.dark .PortalLangBtn.active {
      background: #0F1623;
      color: #3399FF;
      box-shadow: 0 1px 4px rgba(0, 0, 0, 0.45);
    }
    .PortalThemeBtn {
      width: 36px;
      height: 36px;
      border-radius: 10px;
      border: 1px solid #E2E8F0;
      background: #FFFFFF;
      color: #475569;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      transition: all 0.15s ease;
    }
    .PortalThemeBtn:hover {
      background: #F1F5F9;
      color: #0073E6;
      border-color: rgba(0, 115, 230, 0.32);
    }
    html.dark .PortalThemeBtn {
      background: #0F1623;
      border-color: rgba(255, 255, 255, 0.10);
      color: #CBD5E1;
    }
    html.dark .PortalThemeBtn:hover {
      background: rgba(255, 255, 255, 0.06);
      color: #3399FF;
    }

    /* Sidebar Related Navigation Block */
    .PortalSideRelatedWrap {
      padding: 10px 12px;
      border-bottom: 1px solid #E2E8F0;
    }
    html.dark .PortalSideRelatedWrap {
      border-bottom-color: rgba(255, 255, 255, 0.08);
    }
    .PortalSideRelatedLabel {
      padding: 0 8px 6px;
      font-size: 10.5px;
      font-weight: 700;
      letter-spacing: 0.06em;
      text-transform: uppercase;
      color: #64748B;
    }
    .PortalSideRelatedList {
      list-style: none;
      margin: 0;
      padding: 0;
      display: flex;
      flex-direction: column;
      gap: 3px;
      max-height: 158px;
      overflow-y: auto;
    }
    .PortalSideRelatedLink {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 6px;
      padding: 6px 9px;
      border-radius: 8px;
      font-size: 12px;
      font-weight: 600;
      color: #334155;
      text-decoration: none;
      transition: all 0.14s ease;
    }
    .PortalSideRelatedLink:hover {
      background: rgba(0, 115, 230, 0.08);
      color: #0073E6;
    }
    html.dark .PortalSideRelatedLink {
      color: #CBD5E1;
    }
    html.dark .PortalSideRelatedLink:hover {
      background: rgba(51, 153, 255, 0.14);
      color: #3399FF;
    }
    aside#sidebar svg, header#mobile-header svg { max-width: 20px; max-height: 20px; }
    .hidden { display: none; }
    @media print {
      .PortalDetailAppBar { display: none !important; }
    }
    /* Sidebar Height & Nav Optimization */
    @media (min-width: 1024px) {
      .mui-sidebar {
        top: 72px !important;
        max-height: none !important;
        overflow-y: visible !important;
        padding-bottom: 24px !important;
        gap: 10px !important;
      }
    }
    .sidebar-card {
      padding: 12px 14px !important;
      flex-shrink: 0 !important;
    }
    .sidebar-subtitle {
      margin: 4px 0 8px !important;
      font-size: 0.77rem !important;
      line-height: 1.4 !important;
    }
    .sidebar-mini-kpis {
      gap: 4px !important;
      padding-top: 8px !important;
    }
    .mui-sidebar .mui-nav-list {
      gap: 2px !important;
    }
    .mui-sidebar .mui-nav-list:first-of-type {
      max-height: 175px;
      overflow-y: auto;
      scrollbar-width: thin;
      padding-right: 3px;
    }
    .mui-ext-nav-link {
      padding: 4px 8px !important;
      font-size: 0.77rem !important;
      line-height: 1.32 !important;
    }
    .mui-ext-nav-left > span:last-child {
      display: -webkit-box;
      -webkit-line-clamp: 2;
      -webkit-box-orient: vertical;
      overflow: hidden;
    }
    .mui-nav-link {
      padding: 5px 10px !important;
      font-size: 0.80rem !important;
    }


  </style>
<!-- PORTAL_DETAIL_SYNC_CSS_END -->"""


DETAIL_SYNC_JS_TEMPLATE = """<!-- PORTAL_DETAIL_SYNC_JS_START -->
  <script id="portalDetailSyncScript">
    (function () {
      'use strict';
      var THEME_KEY = 'mui_portal_theme_v4';
      var LANG_KEY = 'mui_portal_lang_v1';

      var UI_PAIRS = [
        ['站点与关联导航', 'Site & Related Nav'],
        ['返回数据中心首页', 'Back to Research Hub'],
        ['阅读模式切换', 'Theme Appearance'],
        ['深色护眼模式', 'Dark Mode'],
        ['机构亮色模式', 'Light Mode'],
        ['打印 / 导出 PDF 研报', 'Print / Export PDF'],
        ['导出独立 HTML 文件 (.html)', 'Export Standalone HTML (.html)'],
        ['导出独立 HTML 研报文件', 'Export Standalone HTML (.html)'],
        ['一键导出独立 HTML 文件', 'Export Standalone HTML'],
        ['一键下载独立 HTML 文件', 'Download Standalone HTML'],
        ['导出HTML', 'Export HTML'],
        ['下载 .html', 'Download .html'],
        ['核心研究章节', 'Research Chapters'],
        ['重置清单', 'Reset Checklist'],
        ['00. 执行摘要与核心指标', '00. Executive Summary & KPIs'],
        ['执行摘要与核心指标', 'Executive Summary & KPIs'],
        ['01. 业务与收入架构', '01. Business & Revenue Architecture'],
        ['业务与收入架构', 'Business & Revenue Architecture'],
        ['02. 行业地位与四重护城河', '02. Industry Position & Four Moats'],
        ['行业地位与四重护城河', 'Industry Position & Four Moats'],
        ['02. 行业地位与四重核心护城河', '02. Industry Position & Four Moats'],
        ['行业地位与四重核心护城河', 'Industry Position & Four Moats'],
        ['05. 投资决策与风险配置', '05. Investment Thesis & Risk Allocation'],
        ['投资决策与风险配置', 'Investment Thesis & Risk Allocation'],
        ['05. 三幕演进与管理遗产复盘', '05. Three-Era Evolution & Legacy'],
        ['三幕演进与管理遗产复盘', 'Three-Era Evolution & Legacy'],
        ['06. 季度跟踪 KPI 核查清单', '06. Quarterly KPI Checklist'],
        ['季度跟踪 KPI 核查清单', 'Quarterly KPI Checklist'],
        ['六、 季度基本面跟踪 KPI 交互式核查清单（Checklist）', '06. Interactive Fundamental KPI Checklist'],
        ['核心复盘：', 'Key Takeaway:'],
        ['核心假设：', 'Key Assumptions:'],
        ['数据来源与研究说明：', 'Primary Sources & Notes:'],
        ['免责声明：', 'Disclaimer:'],
        ['研究基准', 'Research Basis'],
        ['牛市情景 (Bull Case)', 'Bull Case Scenario'],
        ['基准情景 (Base Case)', 'Base Case Scenario'],
        ['熊市情景 (Bear Case)', 'Bear Case Scenario'],
        ['多头核心论点 (Bull Thesis)', 'Bull Thesis'],
        ['空头担忧论点 (Bear Thesis)', 'Bear Thesis'],
        ['初始投入本金 (USD)', 'Initial Principal (USD)'],
        ['长期持有年限 (Years)', 'Holding Horizon (Years)'],
        ['年化复合回报率 (CAGR %)', 'Annualized Return (CAGR %)'],
        ['年化复合增长率 (CAGR %)', 'Compound Annual Growth (CAGR %)'],
        ['有效综合年化复合增长率', 'Effective Compound Annual Growth Rate'],
        ['期末预计总资产市值', 'Projected Terminal Value'],
        ['本金增值倍数 (MOIC Multiple)', 'Multiple on Invested Capital (MOIC)'],
        ['累计创造净资本利得', 'Cumulative Net Capital Gain'],
        ['扣除初始投入本金后的纯复利收益', 'Net compounding gain above initial principal'],
        ['四大核心潜在风险评估与跟踪边界', 'Four Key Risks & Monitoring Guardrails'],
        ['核心优劣势与竞合定位评估', 'Strengths, Weaknesses & Competitive Positioning'],
        ['客观局限与组织张力反思 (Critical Lens)', 'Critical Lens & Organizational Trade-offs'],
        ['总市值与最新股价', 'Market Cap & Share Price'],
        ['参考股价 (52周区间)', 'Share Price (52-Wk Range)'],
        ['关注度：极高', 'Priority: Critical'],
        ['关注度：高', 'Priority: High'],
        ['关注度：中高', 'Priority: Med-High'],
        ['关注度：中', 'Priority: Medium']
      ];

      var zh2en = {};
      var en2zh = {};
      for (var i = 0; i < UI_PAIRS.length; i++) {
        zh2en[UI_PAIRS[i][0]] = UI_PAIRS[i][1];
        en2zh[UI_PAIRS[i][1]] = UI_PAIRS[i][0];
      }

      function getStoreStr(k, fb) {
        try {
          var v = window.localStorage.getItem(k);
          return v ? String(v).replace(/^"|"$/g, '') : fb;
        } catch (e) {
          return fb;
        }
      }

      function setStoreStr(k, v) {
        try {
          window.localStorage.setItem(k, v);
        } catch (e) {}
      }

      function renderThemeIcon(isDark) {
        var ic = document.getElementById('portalDetailThemeIcon');
        if (!ic) return;
        ic.innerHTML = isDark
          ? '<svg width="17" height="17" viewBox="0 0 24 24" fill="currentColor"><path d="M12 7c-2.76 0-5 2.24-5 5s2.24 5 5 5 5-2.24 5-5-2.24-5-5-5zM2 13h2c.55 0 1-.45 1-1s-.45-1-1-1H2c-.55 0-1 .45-1 1s.45 1 1 1zm18 0h2c.55 0 1-.45 1-1s-.45-1-1-1h-2c-.55 0-1 .45-1 1s.45 1 1 1zM11 2v2c0 .55.45 1 1 1s1-.45 1-1V2c0-.55-.45-1-1-1s-1 .45-1 1zm0 18v2c0 .55.45 1 1 1s1-.45 1-1v-2c0-.55-.45-1-1-1s-1 .45-1 1zM5.99 4.58a.996.996 0 0 0-1.41 0 .996.996 0 0 0 0 1.41l1.06 1.06c.39.39 1.03.39 1.41 0s.39-1.03 0-1.41L5.99 4.58zm12.37 12.37a.996.996 0 0 0-1.41 0 .996.996 0 0 0 0 1.41l1.06 1.06c.39.39 1.03.39 1.41 0a.996.996 0 0 0 0-1.41l-1.06-1.06zm1.06-10.96a.996.996 0 0 0 0-1.41.996.996 0 0 0-1.41 0l-1.06 1.06c-.39.39-.39 1.03 0 1.41s1.03.39 1.41 0l1.06-1.06zM7.05 18.36a.996.996 0 0 0 0-1.41.996.996 0 0 0-1.41 0l-1.06 1.06c-.39.39-.39 1.03 0 1.41s1.03.39 1.41 0l1.06-1.06z"/></svg>'
          : '<svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M12 3a9 9 0 1 0 9 9c0-.46-.04-.92-.1-1.36a5.389 5.389 0 0 1-4.4 2.26 5.403 5.403 0 0 1-3.14-9.8c-.44-.06-.9-.1-1.36-.1z"/></svg>';
      }

      function applyDetailTheme(mode, persist) {
        var isDark = (mode === 'dark');
        document.documentElement.classList.toggle('dark', isDark);
        document.documentElement.setAttribute('data-mui-theme', isDark ? 'dark' : 'light');
        renderThemeIcon(isDark);
        if (persist) setStoreStr(THEME_KEY, isDark ? 'dark' : 'light');
        if (typeof window.updateChartsTheme === 'function') {
          try { window.updateChartsTheme(); } catch (e) {}
        }
      }

      /* Hook global toggleTheme() so sidebar button also syncs with homepage */
      var origToggleTheme = window.toggleTheme;
      window.toggleTheme = function () {
        var nextDark = !document.documentElement.classList.contains('dark');
        applyDetailTheme(nextDark ? 'dark' : 'light', true);
      };

      var origTextMap = new WeakMap();
      var netTransCache = {};
      var curLang = 'zh';

      function translateDynamicPattern(s, toEn) {
        if (toEn) {
          var m1 = s.match(/^已核查\\s*(\\d+)\\s*\\/\\s*(\\d+)\\s*项\\s*\\((\\d+%)\\)$/);
          if (m1) return 'Verified ' + m1[1] + ' / ' + m1[2] + ' (' + m1[3] + ')';
          var m2 = s.match(/^概率：(\\d+%)$/);
          if (m2) return 'Prob: ' + m2[1];
          var m3 = s.match(/^(\\d+)\\s*年$/);
          if (m3) return m3[1] + ' Yrs';
          var m4 = s.match(/^覆盖时间：(\\d{4})年(\\d{1,2})月$/);
          if (m4) return 'Coverage: ' + m4[1] + '-' + m4[2];
        } else {
          var e1 = s.match(/^Verified\\s*(\\d+)\\s*\\/\\s*(\\d+)\\s*\\((\\d+%)\\)$/);
          if (e1) return '已核查 ' + e1[1] + ' / ' + e1[2] + ' 项 (' + e1[3] + ')';
          var e2 = s.match(/^Prob:\\s*(\\d+%)$/);
          if (e2) return '概率：' + e2[1];
          var e3 = s.match(/^(\\d+)\\s*Yrs$/);
          if (e3) return e3[1] + ' 年';
          var e4 = s.match(/^Coverage:\\s*(\\d{4})-(\\d{1,2})$/);
          if (e4) return '覆盖时间：' + e4[1] + '年' + e4[2] + '月';
        }
        return null;
      }

      function maybeTranslateRemainingOnline(nodesToTranslate) {
        if (!nodesToTranslate.length || typeof window.fetch !== 'function') return;
        if (typeof navigator !== 'undefined' && navigator.onLine === false) return;
        var batchSize = 18;
        var idx = 0;
        function nextBatch() {
          if (curLang !== 'en' || idx >= nodesToTranslate.length) return;
          var slice = nodesToTranslate.slice(idx, idx + batchSize);
          idx += batchSize;
          var joined = slice.map(function (it) { return it.zh; }).join('\\n====\\n');
          var ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
          var timer = ctrl ? setTimeout(function () { ctrl.abort(); }, 3500) : 0;
          window.fetch(
            'https://translate.googleapis.com/translate_a/single?client=gtx&sl=zh-CN&tl=en&dt=t&q=' + encodeURIComponent(joined),
            ctrl ? { signal: ctrl.signal } : {}
          ).then(function (r) {
            if (timer) clearTimeout(timer);
            return r.ok ? r.json() : null;
          }).then(function (data) {
            if (!data || !Array.isArray(data[0]) || curLang !== 'en') return;
            var fullEn = data[0].map(function (seg) { return seg && seg[0] ? seg[0] : ''; }).join('');
            var parts = fullEn.split(/\\s*====\\s*/);
            if (parts.length === slice.length) {
              for (var k = 0; k < slice.length; k++) {
                var enTxt = parts[k].trim();
                if (enTxt) {
                  netTransCache[slice[k].zh] = enTxt;
                  if (curLang === 'en') slice[k].node.nodeValue = slice[k].lead + enTxt + slice[k].trail;
                }
              }
            }
            nextBatch();
          }).catch(function () {
            if (timer) clearTimeout(timer);
          });
        }
        nextBatch();
      }

      function applyDetailLang(lang, persist) {
        curLang = (lang === 'en') ? 'en' : 'zh';
        var en = (curLang === 'en');
        document.documentElement.setAttribute('lang', en ? 'en' : 'zh-CN');
        document.documentElement.setAttribute('data-lang', curLang);

        var zhB = document.getElementById('portalDetailZhBtn');
        var enB = document.getElementById('portalDetailEnBtn');
        if (zhB && enB) {
          zhB.classList.toggle('active', !en);
          enB.classList.toggle('active', en);
          zhB.setAttribute('aria-pressed', !en ? 'true' : 'false');
          enB.setAttribute('aria-pressed', en ? 'true' : 'false');
        }
        var thB = document.getElementById('portalDetailThemeBtn');
        if (thB) {
          thB.setAttribute('title', en ? 'Toggle Light / Dark theme' : '切换浅色 / 深色主题');
          thB.setAttribute('aria-label', en ? 'Toggle color theme' : '切换深浅色主题');
        }

        var tagged = document.querySelectorAll('[data-i18n-zh][data-i18n-en]');
        for (var d = 0; d < tagged.length; d++) {
          tagged[d].textContent = en ? tagged[d].getAttribute('data-i18n-en') : tagged[d].getAttribute('data-i18n-zh');
        }

        var cjkRe = /[\\u4e00-\\u9fff]/;
        var pendingOnline = [];
        var walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
        var n;
        while ((n = walker.nextNode())) {
          var p = n.parentElement;
          if (!p) continue;
          var tn = p.tagName;
          if (tn === 'SCRIPT' || tn === 'STYLE' || p.closest('#portalDetailLangSwitch') || p.closest('[data-i18n-zh]')) continue;
          var raw = n.nodeValue || '';
          var trimmed = raw.replace(/\\s+/g, ' ').trim();
          if (!trimmed) continue;

          if (!origTextMap.has(n) && cjkRe.test(trimmed)) {
            origTextMap.set(n, raw);
          }

          var lead = (raw.match(/^\\s*/) || [''])[0];
          var trail = (raw.match(/\\s*$/) || [''])[0];

          if (en) {
            var baseRaw = origTextMap.has(n) ? origTextMap.get(n) : raw;
            var baseTrim = baseRaw.replace(/\\s+/g, ' ').trim();
            var hit = zh2en[baseTrim] || translateDynamicPattern(baseTrim, true) || netTransCache[baseTrim];
            if (hit) {
              n.nodeValue = lead + hit + trail;
            } else if (cjkRe.test(baseTrim) && baseTrim.length >= 2) {
              pendingOnline.push({ node: n, zh: baseTrim, lead: lead, trail: trail });
            }
          } else {
            if (origTextMap.has(n)) {
              n.nodeValue = origTextMap.get(n);
            } else {
              var rev = en2zh[trimmed] || translateDynamicPattern(trimmed, false);
              if (rev) n.nodeValue = lead + rev + trail;
            }
          }
        }

        if (persist) setStoreStr(LANG_KEY, curLang);
        if (en && pendingOnline.length) {
          maybeTranslateRemainingOnline(pendingOnline);
        }
      }

      function initDetailPortalSync() {
        var savedTheme = getStoreStr(THEME_KEY, 'light');
        applyDetailTheme(savedTheme === 'dark' ? 'dark' : 'light', false);

        var savedLang = getStoreStr(LANG_KEY, 'zh');
        applyDetailLang(savedLang === 'en' ? 'en' : 'zh', false);

        var tBtn = document.getElementById('portalDetailThemeBtn');
        if (tBtn) {
          tBtn.addEventListener('click', function () {
            var nextDark = !document.documentElement.classList.contains('dark');
            applyDetailTheme(nextDark ? 'dark' : 'light', true);
          });
        }
        var zBtn = document.getElementById('portalDetailZhBtn');
        var eBtn = document.getElementById('portalDetailEnBtn');
        if (zBtn) zBtn.addEventListener('click', function () { applyDetailLang('zh', true); });
        if (eBtn) eBtn.addEventListener('click', function () { applyDetailLang('en', true); });

        window.addEventListener('storage', function (ev) {
          if (ev.key === THEME_KEY && ev.newValue) {
            var nt = String(ev.newValue).replace(/^"|"$/g, '');
            if (nt === 'light' || nt === 'dark') applyDetailTheme(nt, false);
          } else if (ev.key === LANG_KEY && ev.newValue) {
            var nl = String(ev.newValue).replace(/^"|"$/g, '');
            if (nl === 'zh' || nl === 'en') applyDetailLang(nl, false);
          }
        });
      }

      if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initDetailPortalSync);
      } else {
        initDetailPortalSync();
      }
    })();
  </script>
<!-- PORTAL_DETAIL_SYNC_JS_END -->"""



def build_company_detail_hub_block(comp_folder: str, cur_rel: str, comp_items: list[dict]) -> str:
    """为每个公司的详情页顶部生成该公司子目录分类切换面板（含上一/下一公司、打开独立网页、导出HTML与分段标签栏）。"""
    if not comp_items:
        return ""
    cur_dir = os.path.dirname(cur_rel)
    cat_zh = {
        "Company": "公司深度", "SeniorExecutive": "高管人物", "CorporateCulture": "企业文化",
        "software": "软件生态", "hardware": "硬件产品", "patent": "核心专利",
    }
    counts: dict[str, int] = {}
    for it in comp_items:
        c = it["category"]
        counts[c] = counts.get(c, 0) + 1
    cats = sorted(counts.keys())

    # Resolve Prev / Next / Current hrefs
    cur_idx = 0
    for i, it in enumerate(comp_items):
        if it["relPath"] == cur_rel:
            cur_idx = i
            break
    prev_it = comp_items[(cur_idx - 1) % len(comp_items)]
    next_it = comp_items[(cur_idx + 1) % len(comp_items)]
    cur_it = comp_items[cur_idx]

    def _href(it):
        return os.path.relpath(it["relPath"], cur_dir).replace(os.sep, "/") if cur_dir else it["relPath"]

    prev_href = _href(prev_it)
    next_href = _href(next_it)
    cur_href = _href(cur_it)
    cur_short = (cur_it.get("companyLabel") or comp_folder).split()[0]

    pills_html = [
        f'<button type="button" class="co-hub-pill active" data-co-cat="ALL" data-i18n-zh="全部 ({len(comp_items)})" data-i18n-en="All ({len(comp_items)})">全部 ({len(comp_items)})</button>'
    ]
    for c in cats:
        lbl = cat_zh.get(c, c)
        pills_html.append(
            f'<button type="button" class="co-hub-pill" data-co-cat="{html_mod.escape(c)}" data-i18n-zh="{html_mod.escape(lbl)} ({counts[c]})" data-i18n-en="{html_mod.escape(c)} ({counts[c]})">{html_mod.escape(lbl)} ({counts[c]})</button>'
        )

    cards_html = []
    for it in comp_items:
        rel_href = _href(it)
        is_cur = (it["relPath"] == cur_rel)
        cls = "co-hub-item active" if is_cur else "co-hub-item"
        target_attr = ' target="_blank" rel="noopener noreferrer"' if it.get("isIndex") else ""
        idx_badge = " · index.html ↗" if it.get("isIndex") else ""
        t_zh = it.get("title") or it["subPath"]
        t_en = it.get("titleEn") or t_zh
        cards_html.append(
            f'<a class="{cls}" data-co-item-cat="{html_mod.escape(it["category"])}" href="{html_mod.escape(rel_href)}"{target_attr} title="{html_mod.escape(t_zh)}">'
            f'<span class="co-hub-ticker mono">{html_mod.escape((it.get("ticker") or comp_folder)[:6])}</span>'
            f'<span class="co-hub-info">'
            f'<span class="co-hub-title" data-i18n-zh="{html_mod.escape(t_zh)}" data-i18n-en="{html_mod.escape(t_en)}">{html_mod.escape(t_zh)}</span>'
            f'<span class="co-hub-meta mono">{html_mod.escape(it["subPath"] + idx_badge)}</span>'
            f'</span></a>'
        )

    return (
        "<!-- COMPANY_SUBFOLDER_HUB_START -->\n"
        '      <section class="co-hub-card" id="companySubfolderHub" aria-label="Company Directory">\n'
        '        <div class="co-hub-head">\n'
        f'          <span class="mui-chip mui-chip-primary mono">{html_mod.escape(comp_folder.upper())} · {len(comp_items)} FILES</span>\n'
        '          <strong class="co-hub-heading" data-i18n-zh="本公司目录与研报切换中枢（按首个子目录分类）" data-i18n-en="Company Subfolder &amp; Report Directory (Grouped by First Subfolder)">本公司目录与研报切换中枢（按首个子目录分类）</strong>\n'
        '        </div>\n'
        '        <div class="co-hub-actions">\n'
        f'          <a class="co-hub-btn" id="coHubPrevBtn" href="{html_mod.escape(prev_href)}" data-i18n-zh="← 上一公司" data-i18n-en="← Prev">← 上一公司</a>\n'
        f'          <a class="co-hub-btn" id="coHubNextBtn" href="{html_mod.escape(next_href)}" data-i18n-zh="下一公司 →" data-i18n-en="Next →">下一公司 →</a>\n'
        f'          <a class="co-hub-btn co-hub-btn-primary" id="coHubOpenBtn" href="{html_mod.escape(cur_href)}" target="_blank" rel="noopener noreferrer" data-i18n-zh="打开《{html_mod.escape(cur_short)}》独立网页 ↗" data-i18n-en="Open Standalone Page ↗">打开《{html_mod.escape(cur_short)}》独立网页 ↗</a>\n'
        '          <button type="button" class="co-hub-btn" id="coHubExportBtn" data-i18n-zh="导出 HTML" data-i18n-en="Export HTML">导出 HTML</button>\n'
        '        </div>\n'
        '        <div class="co-hub-pills">\n          ' + "\n          ".join(pills_html) + '\n        </div>\n'
        '        <div class="co-hub-grid">\n          ' + "\n          ".join(cards_html) + '\n        </div>\n'
        '        <script>(function(){var h=document.getElementById("companySubfolderHub");if(!h)return;h.addEventListener("click",function(e){var b=e.target.closest("[data-co-cat]");if(!b)return;var cat=b.getAttribute("data-co-cat");var ps=h.querySelectorAll("[data-co-cat]");for(var i=0;i<ps.length;i++)ps[i].classList.toggle("active",ps[i]===b);var its=h.querySelectorAll("[data-co-item-cat]");for(var j=0;j<its.length;j++){its[j].style.display=(cat==="ALL"||its[j].getAttribute("data-co-item-cat")===cat)?"":"none";}});var ex=document.getElementById("coHubExportBtn");if(ex){ex.addEventListener("click",function(){try{var html="<!DOCTYPE html>"+String.fromCharCode(10)+document.documentElement.outerHTML;var blob=new Blob([html],{type:"text/html;charset=utf-8"});var a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=(document.title||"report").replace(/[^a-zA-Z0-9_\u4e00-\u9fff-]/g,"_")+"_"+new Date().toISOString().slice(0,10)+".html";document.body.appendChild(a);a.click();setTimeout(function(){URL.revokeObjectURL(a.href);a.remove();},400);}catch(err){}});}})();</script>\n'
        '      </section>\n'
        "<!-- COMPANY_SUBFOLDER_HUB_END -->"
    )


def build_mui_related_nav_block(rel_to_index: str, siblings: list[tuple[str, str, str]]) -> str:
    """为 MUI 架构详情页生成侧边栏『站点与关联导航』列表。"""
    home_icon = '<svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M2 5.2L9.5 9.5v8.3L2 13.5V5.2zm20 0L14.5 9.5v8.3L22 13.5V5.2zM12 2L4.6 6.2 12 10.4l7.4-4.2L12 2z"/></svg>'
    doc_icon = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>'

    items_html = [
        "          <li>\n"
        f'            <a href="{html_mod.escape(rel_to_index)}" class="mui-ext-nav-link">\n'
        '              <span class="mui-ext-nav-left">\n'
        f'                <span class="mui-ext-nav-icon">{home_icon}</span>\n'
        '                <span data-i18n-zh="返回数据中心首页" data-i18n-en="Back to Research Hub">返回数据中心首页</span>\n'
        "              </span>\n"
        '              <span class="mui-ext-nav-arrow">↩</span>\n'
        "            </a>\n"
        "          </li>"
    ]
    for href, zh_lbl, en_lbl in siblings:
        items_html.append(
            "          <li>\n"
            f'            <a href="{html_mod.escape(href)}" class="mui-ext-nav-link">\n'
            '              <span class="mui-ext-nav-left">\n'
            f'                <span class="mui-ext-nav-icon">{doc_icon}</span>\n'
            f'                <span data-i18n-zh="{html_mod.escape(zh_lbl)}" data-i18n-en="{html_mod.escape(en_lbl)}">{html_mod.escape(zh_lbl)}</span>\n'
            "              </span>\n"
            '              <span class="mui-ext-nav-arrow">↗</span>\n'
            "            </a>\n"
            "          </li>"
        )
    return (
        "<!-- SITE_RELATED_NAV_START -->\n"
        '        <div class="nav-section-label" data-i18n-zh="站点与关联导航" data-i18n-en="Site &amp; Related Nav">站点与关联导航</div>\n'
        '        <ul class="mui-nav-list">\n'
        + "\n".join(items_html)
        + "\n        </ul>\n"
        "        <!-- SITE_RELATED_NAV_END -->"
    )


def build_classic_related_nav_block(rel_to_index: str, siblings: list[tuple[str, str, str]]) -> str:
    """为经典侧边栏详情页生成『站点与关联导航』模块。"""
    rows = [
        f'        <li><a href="{html_mod.escape(rel_to_index)}" class="PortalSideRelatedLink">'
        '<span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" data-i18n-zh="🏠 返回数据中心首页" data-i18n-en="🏠 Back to Research Hub">🏠 返回数据中心首页</span>'
        '<span style="opacity:0.65;">↩</span></a></li>'
    ]
    for href, zh_lbl, en_lbl in siblings:
        rows.append(
            f'        <li><a href="{html_mod.escape(href)}" class="PortalSideRelatedLink">'
            f'<span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" data-i18n-zh="{html_mod.escape(zh_lbl)}" data-i18n-en="{html_mod.escape(en_lbl)}">{html_mod.escape(zh_lbl)}</span>'
            '<span style="opacity:0.65;">↗</span></a></li>'
        )
    return (
        "<!-- SITE_RELATED_NAV_START -->\n"
        '      <div class="PortalSideRelatedWrap">\n'
        '        <div class="PortalSideRelatedLabel" data-i18n-zh="站点与关联导航" data-i18n-en="Site &amp; Related Nav">站点与关联导航</div>\n'
        '        <ul class="PortalSideRelatedList">\n'
        + "\n".join(rows)
        + "\n        </ul>\n"
        "      </div>\n"
        "      <!-- SITE_RELATED_NAV_END -->"
    )


def build_classic_topbar_block(rel_to_index: str, comp_zh: str, comp_en: str, cat_zh: str, cat_en: str) -> str:
    """为经典详情页顶部注入与首页一致的返回面包屑 + [中文 | EN] + [🌙/☀️] 顶栏。"""
    crumb_zh = f"{comp_zh} · {cat_zh}"
    crumb_en = f"{comp_en} · {cat_en}"
    return (
        "<!-- PORTAL_DETAIL_TOPBAR_START -->\n"
        '    <div class="PortalDetailAppBar">\n'
        '      <div class="PortalDetailAppBar-left">\n'
        f'        <a href="{html_mod.escape(rel_to_index)}" class="PortalBackLink">'
        '<span data-i18n-zh="← 数据中心首页" data-i18n-en="← Research Hub">← 数据中心首页</span></a>\n'
        f'        <span class="PortalCrumbMeta" data-i18n-zh="{html_mod.escape(crumb_zh)}" data-i18n-en="{html_mod.escape(crumb_en)}">{html_mod.escape(crumb_zh)}</span>\n'
        "      </div>\n"
        '      <div class="PortalDetailAppBar-right">\n'
        '        <div id="portalDetailLangSwitch" class="PortalLangSwitch" role="group" aria-label="语言切换">\n'
        '          <button type="button" id="portalDetailZhBtn" class="PortalLangBtn active" aria-pressed="true">中文</button>\n'
        '          <button type="button" id="portalDetailEnBtn" class="PortalLangBtn" aria-pressed="false">EN</button>\n'
        "        </div>\n"
        '        <button type="button" id="portalDetailThemeBtn" class="PortalThemeBtn" aria-label="切换深浅色主题" title="切换浅色 / 深色主题">\n'
        '          <span id="portalDetailThemeIcon" style="display:inline-flex;align-items:center;"></span>\n'
        "        </button>\n"
        "      </div>\n"
        "    </div>\n"
        "    <!-- PORTAL_DETAIL_TOPBAR_END -->"
    )


def sync_all_detail_pages(
    docs_dir: str,
    content_files: list[dict],
    new_cache: dict[str, dict],
    subfolder_by_comp: dict[str, list[dict]] | None = None,
) -> tuple[int, int]:
    """
    遍历全部 HTML 详情页，自动更新：
    1) 站点与关联导航（相对路径返回首页 + 同公司全部关联研报）
    2) 与首页完全一致的浅色/深色主题状态同步 (mui_portal_theme_v4)
    3) 与首页完全一致的右上角 [中文 | EN] 双语切换 (mui_portal_lang_v1)
    仅在内容实际变化时原子写盘，并同步更新增量缓存指纹，确保二次运行零冗余开销。
    """
    index_abs = os.path.join(docs_dir, "index.html")
    by_company: dict[str, list[dict]] = {}
    for rec in content_files:
        by_company.setdefault(str(rec.get("company") or ""), []).append(rec)

    for comp_k in by_company:
        by_company[comp_k].sort(
            key=lambda r: (
                CATEGORY_SORT_ORDER.get(str(r.get("category") or ""), 99),
                str(r.get("relPath") or "").lower(),
            )
        )

    checked_count = 0
    updated_count = 0

    for rec in content_files:
        rel_p = str(rec.get("relPath") or "")
        if not rel_p.lower().endswith((".html", ".htm")):
            continue
        full_p = os.path.join(docs_dir, rel_p.replace("/", os.sep))
        if not os.path.isfile(full_p) or is_portal_html(full_p):
            continue

        orig_html = read_text_with_fallback(full_p)
        if not orig_html or "<body" not in orig_html.lower():
            continue

        checked_count += 1
        page_dir = os.path.dirname(full_p)
        rel_to_index = encode_relative_url(page_dir, index_abs)

        comp_folder = str(rec.get("company") or "根目录")
        comp_label_zh = str(rec.get("companyLabel") or comp_folder)
        cat = str(rec.get("category") or "Company")
        cat_zh = CATEGORY_KNOWN.get(cat, cat)
        cat_en = CATEGORY_KNOWN_EN.get(cat, cat)

        # 构建同公司兄弟研报列表（排除自身与同名重复副本）
        siblings_raw = by_company.get(comp_folder, []) if comp_folder != "根目录" else []
        seen_names = {os.path.basename(full_p).lower()}
        sibling_links: list[tuple[str, str, str]] = []
        for sib in siblings_raw:
            sib_rel = str(sib.get("relPath") or "")
            sib_name_low = os.path.basename(sib_rel).lower()
            if sib_rel == rel_p or sib_name_low in seen_names:
                continue
            seen_names.add(sib_name_low)
            sib_abs = os.path.join(docs_dir, sib_rel.replace("/", os.sep))
            href = encode_relative_url(page_dir, sib_abs)
            zh_lbl, en_lbl = build_concise_nav_labels(sib)
            sibling_links.append((href, zh_lbl, en_lbl))

        work_html = orig_html
        is_mui_detail = 'id="muiSidebar"' in work_html


        # 5. 为外部 CDN 脚本（如 Chart.js）自动添加 defer，避免阻塞详情页首屏渲染
        def _defer_ext_script(m):
            tag = m.group(0)
            if re.search(r'\b(?:defer|async)\b', tag, flags=re.IGNORECASE):
                return tag
            return tag.replace("<script ", "<script defer ", 1)
        work_html = re.sub(r'<script\s+[^>]*src=["\']https?://[^"\']+["\'][^>]*>', _defer_ext_script, work_html, flags=re.IGNORECASE)


        # 6. 为每份详情页注入所属公司的『子目录分类切换面板』
        co_hub_pat = re.compile(r"<!--\s*COMPANY_SUBFOLDER_HUB_START\s*-->[\s\S]*?<!--\s*COMPANY_SUBFOLDER_HUB_END\s*-->")
        has_folder_index = os.path.isfile(os.path.join(os.path.dirname(full_p), "index.html"))
        comp_items = (subfolder_by_comp or {}).get(comp_folder, [])
        co_hub_html = "" if has_folder_index else build_company_detail_hub_block(comp_folder, rel_p, comp_items)
        if co_hub_pat.search(work_html):
            work_html = co_hub_pat.sub(lambda _: co_hub_html, work_html, count=1)
        elif co_hub_html:
            main_m = re.search(r"<main[^>]*>", work_html, flags=re.IGNORECASE)
            if main_m:
                pos = main_m.end()
                work_html = work_html[:pos] + "\n" + co_hub_html + work_html[pos:]

        nav_marker_pat = re.compile(r"<!--\s*SITE_RELATED_NAV_START\s*-->[\s\S]*?<!--\s*SITE_RELATED_NAV_END\s*-->")
        css_marker_pat = re.compile(r"<!--\s*PORTAL_DETAIL_SYNC_CSS_START\s*-->[\s\S]*?<!--\s*PORTAL_DETAIL_SYNC_CSS_END\s*-->")
        top_marker_pat = re.compile(r"<!--\s*PORTAL_DETAIL_TOPBAR_START\s*-->[\s\S]*?<!--\s*PORTAL_DETAIL_TOPBAR_END\s*-->")
        js_marker_pat = re.compile(r"<!--\s*PORTAL_DETAIL_SYNC_JS_START\s*-->[\s\S]*?<!--\s*PORTAL_DETAIL_SYNC_JS_END\s*-->")

        if is_mui_detail:
            # 1. 原地幂等更新 MUI 详情页的『站点与关联导航』块
            mui_nav_block = build_mui_related_nav_block(rel_to_index, sibling_links)
            if nav_marker_pat.search(work_html):
                work_html = nav_marker_pat.sub(lambda _: mui_nav_block, work_html, count=1)
            else:
                mui_nav_pat = re.compile(
                    r'<div\s+class="nav-section-label"[^>]*>(?:站点与关联导航|Site\s*&amp;\s*Related\s*Nav)</div>\s*<ul\s+class="mui-nav-list">[\s\S]*?</ul>',
                    flags=re.IGNORECASE,
                )
                if mui_nav_pat.search(work_html):
                    work_html = mui_nav_pat.sub(lambda _: mui_nav_block, work_html, count=1)

            # 2. 确保主题与语言存储键和首页一致
            work_html = re.sub(
                r"var\s+STORAGE_THEME_KEY\s*=\s*'[^']+';",
                "var STORAGE_THEME_KEY = 'mui_portal_theme_v4';",
                work_html,
            )
            work_html = re.sub(
                r"var\s+STORAGE_LANG_KEY\s*=\s*'[^']+';",
                "var STORAGE_LANG_KEY = 'mui_portal_lang_v1';",
                work_html,
            )

            
            # 4. 将顶栏『数据中心』面包屑链接统一指向本地首页并在当前页打开（不新开标签页）
            def _rewrite_crumb(m):
                tag = m.group(0)
                tag = re.sub(r'href="[^"]*"', f'href="{html_mod.escape(rel_to_index)}"', tag)
                tag = re.sub(r'\s+target="_blank"', '', tag)
                tag = re.sub(r'\s+rel="noopener noreferrer"', '', tag)
                return tag
            work_html = re.sub(r'<a\s+class="crumb-link[^"]*"[^>]*>', _rewrite_crumb, work_html)

            # 3. 原地幂等注入/更新左侧边栏高度与移动端零溢出优化样式
            mui_opt_pat = re.compile(r"<!--\s*PORTAL_MUI_OPT_CSS_START\s*-->[\s\S]*?<!--\s*PORTAL_MUI_OPT_CSS_END\s*-->")
            if mui_opt_pat.search(work_html):
                work_html = mui_opt_pat.sub(lambda _: PORTAL_MUI_OPT_CSS_BLOCK, work_html, count=1)
            elif "</head>" in work_html:
                work_html = work_html.replace("</head>", PORTAL_MUI_OPT_CSS_BLOCK + "\n</head>", 1)
        else:
            # 经典详情页：1. 原地更新或首次注入无闪烁主题与 MUI 浅/深色样式
            if css_marker_pat.search(work_html):
                work_html = css_marker_pat.sub(lambda _: DETAIL_SYNC_CSS_TEMPLATE, work_html, count=1)
            elif "</head>" in work_html:
                work_html = work_html.replace("</head>", DETAIL_SYNC_CSS_TEMPLATE + "\n</head>", 1)

            # 2. 原地更新或首次注入左侧边栏『站点与关联导航』
            classic_nav = build_classic_related_nav_block(rel_to_index, sibling_links)
            if nav_marker_pat.search(work_html):
                work_html = nav_marker_pat.sub(lambda _: classic_nav, work_html, count=1)
            else:
                toc_anchor = re.search(
                    r'(<div\s+class="px-3\s+py-3\s+overflow-y-auto[^"]*">)',
                    work_html,
                )
                if toc_anchor:
                    work_html = (
                        work_html[: toc_anchor.start()]
                        + classic_nav
                        + "\n      "
                        + work_html[toc_anchor.start() :]
                    )
                elif '<aside id="sidebar"' in work_html:
                    m_aside_end = re.search(r"</aside>", work_html, flags=re.IGNORECASE)
                    if m_aside_end:
                        work_html = work_html[: m_aside_end.start()] + classic_nav + "\n  " + work_html[m_aside_end.start() :]

            # 3. 原地更新或首次注入与首页一致的 [中文 | EN] + [🌙/☀️] 顶栏
            topbar_html = build_classic_topbar_block(
                rel_to_index=rel_to_index,
                comp_zh=comp_label_zh,
                comp_en=comp_folder,
                cat_zh=cat_zh,
                cat_en=cat_en,
            )
            if top_marker_pat.search(work_html):
                work_html = top_marker_pat.sub(lambda _: topbar_html, work_html, count=1)
            else:
                body_pos = work_html.lower().find("<body")
                m_main = re.search(r"<main\b[^>]*>", work_html[body_pos:] if body_pos != -1 else work_html, flags=re.IGNORECASE)
                if m_main and body_pos != -1:
                    class _M:
                        def __init__(self, e): self._e = e
                        def end(self): return self._e
                    m_main = _M(body_pos + m_main.end())
                if m_main:
                    ins_pos = m_main.end()
                    work_html = work_html[:ins_pos] + "\n    " + topbar_html + work_html[ins_pos:]
                else:
                    m_body = re.search(r"<body\b[^>]*>", work_html, flags=re.IGNORECASE)
                    if m_body:
                        ins_pos = m_body.end()
                        work_html = work_html[:ins_pos] + "\n  " + topbar_html + work_html[ins_pos:]

            # 4. 原地更新或首次注入双语与深浅色联动脚本
            if js_marker_pat.search(work_html):
                work_html = js_marker_pat.sub(lambda _: DETAIL_SYNC_JS_TEMPLATE, work_html, count=1)
            elif "</body>" in work_html:
                work_html = work_html.replace("</body>", DETAIL_SYNC_JS_TEMPLATE + "\n</body>", 1)

        if work_html != orig_html:
            tmp_detail = full_p + ".sync.tmp"
            with open(tmp_detail, "w", encoding="utf-8") as f:
                f.write(work_html)
            os.replace(tmp_detail, full_p)
            updated_count += 1

            # 更新缓存中的 size 与 mtime_ns，使下一次增量运行直接命中缓存
            try:
                st_after = os.stat(full_p)
                rec["size"] = st_after.st_size
                if rel_p in new_cache:
                    new_cache[rel_p]["size"] = st_after.st_size
                    new_cache[rel_p]["mtime_ns"] = get_mtime_ns(st_after)
            except OSError:
                pass

    return checked_count, updated_count


def inject_catalog_and_nav_into_html(
    html_src: str,
    cat_json: str,
    nav_links: object | None,
    ver_tag: str,
) -> tuple[str, bool]:
    """
    自愈式 HTML 注入引擎：
    1. 自动清除从画布预览版 (Canvas) 下载时残留的 docsInlineFullText 静态快照块
    2. 三级回退匹配注入 window.DOCS_CATALOG 与 window.DOCS_NAV_LINKS，自动恢复标准注释锚点
    3. 自动恢复或更新带防缓存版本号的 ./search_index.js?v=... 引用（同时修复浏览器另存为篡改的 _files 路径）
    """
    html_src = re.sub(
        r"\s*<!-- FULLTEXT_INLINE_START -->[\s\S]*?<!-- FULLTEXT_INLINE_END -->",
        "",
        html_src,
    )
    html_src = re.sub(
        r'\s*<script[^>]*\bid=["\']docsInlineFullText["\'][^>]*>[\s\S]*?</script>',
        "",
        html_src,
        flags=re.IGNORECASE,
    )

    new_cat_block = (
        "<!-- DOCS_CATALOG_START -->\n"
        '  <script id="docsCatalogData">\n'
        f"    window.DOCS_CATALOG = {cat_json};\n"
        "  </script>\n"
        "  <!-- DOCS_CATALOG_END -->"
    )
    catalog_injected = False
    for pattern in (
        r"<!--\s*DOCS_CATALOG_START\s*-->[\s\S]*?<!--\s*DOCS_CATALOG_END\s*-->",
        r'<script[^>]*\bid=["\']docsCatalogData["\'][^>]*>[\s\S]*?</script>',
        r"<script[^>]*>\s*window\.DOCS_CATALOG\s*=[\s\S]*?</script>",
    ):
        if re.search(pattern, html_src, flags=re.IGNORECASE):
            html_src = re.sub(pattern, lambda _: new_cat_block, html_src, count=1, flags=re.IGNORECASE)
            catalog_injected = True
            break

    if nav_links is not None:
        nav_json = json.dumps(nav_links, ensure_ascii=False, indent=2).replace("</script>", "<\\/script>")
        new_nav_block = (
            "<!-- NAV_CONFIG_START -->\n"
            '  <script id="docsNavConfig">\n'
            f"    window.DOCS_NAV_LINKS = {nav_json};\n"
            "  </script>\n"
            "  <!-- NAV_CONFIG_END -->"
        )
        nav_injected = False
        for nav_pat in (
            r"<!--\s*NAV_CONFIG_START\s*-->[\s\S]*?<!--\s*NAV_CONFIG_END\s*-->",
            r'<script[^>]*\bid=["\']docsNavConfig["\'][^>]*>[\s\S]*?</script>',
            r"<script[^>]*>\s*window\.DOCS_NAV_LINKS\s*=[\s\S]*?</script>",
        ):
            if re.search(nav_pat, html_src, flags=re.IGNORECASE):
                html_src = re.sub(nav_pat, lambda _: new_nav_block, html_src, count=1, flags=re.IGNORECASE)
                nav_injected = True
                break
        if not nav_injected and catalog_injected:
            html_src = html_src.replace(
                "<!-- DOCS_CATALOG_START -->",
                new_nav_block + "\n\n  <!-- DOCS_CATALOG_START -->",
                1,
            )

    search_script_tag = f'<script id="docsSearchIndexScript" src="./search_index.js?v={ver_tag}" defer></script>'
    existing_search_pat = (
        r'<script[^>]*(?:\bid=["\']docsSearchIndexScript["\']|\bsrc=["\'][^"\']*search_index\.js[^"\']*["\'])[^>]*>\s*</script>'
    )
    if re.search(existing_search_pat, html_src, flags=re.IGNORECASE):
        html_src = re.sub(existing_search_pat, lambda _: search_script_tag, html_src, count=1, flags=re.IGNORECASE)
    elif "<!-- DOCS_CATALOG_END -->" in html_src:
        html_src = html_src.replace(
            "<!-- DOCS_CATALOG_END -->",
            "<!-- DOCS_CATALOG_END -->\n\n  <!-- Full-Text Search Corpus -->\n  " + search_script_tag,
            1,
        )

    return html_src, catalog_injected


def build_catalog_and_index(
    docs_dir: str,
    rebuild: bool = False,
    max_chars: int = 250000,
    standalone_out: str | None = None,
    sync_details: bool = True,
) -> dict[str, int | str]:
    index_html_path = os.path.join(docs_dir, "index.html")
    if not os.path.isfile(index_html_path):
        raise FileNotFoundError(
            f"未找到数据中心主页文件：{index_html_path}\n"
            f"请确认 index.html 位于 docs/ 文件夹内，或使用 --dir 指定正确的 docs 目录路径。"
        )

    cache_path = os.path.join(docs_dir, ".index_cache.json")
    drive_meta_path = os.path.join(docs_dir, ".drive_metadata.json")
    nav_links_path = os.path.join(docs_dir, "nav_links.json")
    search_js_path = os.path.join(docs_dir, "search_index.js")

    raw_cache = {} if rebuild else (load_json_safe(cache_path, {}) or {})
    old_cache: dict[str, dict] = raw_cache if isinstance(raw_cache, dict) else {}
    has_prior_cache = bool(old_cache)

    raw_meta = load_json_safe(drive_meta_path, {}) or {}
    drive_meta: dict = raw_meta if isinstance(raw_meta, dict) else {}
    nav_links = load_json_safe(nav_links_path, None)

    root_info = drive_meta.get("root") or {
        "id": "root_docs",
        "name": os.path.basename(docs_dir) or "docs",
        "driveUrl": DEFAULT_ROOT_DRIVE_URL,
    }
    known_companies: dict = drive_meta.get("companies") or {}
    known_folders: dict = drive_meta.get("folders") or {}
    known_files: dict = drive_meta.get("files") or {}

    top_folders: list[str] = sorted(
        d for d in os.listdir(docs_dir)
        if os.path.isdir(os.path.join(docs_dir, d))
        and not d.startswith(".")
        and d not in IGNORED_DIRS
        and not d.endswith("_files")
    )

    folder_id_by_rel: dict[str, str] = {"": root_info["id"]}
    companies_list: list[dict] = []
    extra_subfolders: list[dict] = []
    total_subfolders = 0

    for comp_folder in top_folders:
        total_subfolders += 1
        c_meta = known_companies.get(comp_folder, {})
        comp_id = c_meta.get("id") or stable_id("dir", comp_folder)
        comp_drive = c_meta.get("driveUrl") or root_info.get("driveUrl", DEFAULT_ROOT_DRIVE_URL)
        comp_label = c_meta.get("label") or comp_folder
        comp_ticker = (
            c_meta.get("ticker")
            or re.sub(r"[^A-Za-z0-9.&-]", "", comp_folder).upper()[:8]
            or comp_folder[:6].upper()
        )
        comp_sector = c_meta.get("sector") or "新增研究覆盖"
        folder_id_by_rel[comp_folder] = comp_id

        comp_abs = os.path.join(docs_dir, comp_folder)
        subs_dict: dict[str, dict] = {}

        for dp, dns, _ in os.walk(comp_abs):
            dns[:] = sorted(
                d for d in dns
                if not d.startswith(".") and d not in IGNORED_DIRS and not d.endswith("_files")
            )
            for dn in dns:
                total_subfolders += 1
                sub_abs = os.path.join(dp, dn)
                rel_sub = os.path.relpath(sub_abs, docs_dir).replace(os.sep, "/")
                parts = rel_sub.split("/")
                level = len(parts)
                f_meta = known_folders.get(rel_sub, {})
                fid = f_meta.get("id") or stable_id("dir", rel_sub)
                f_drive = f_meta.get("driveUrl") or comp_drive
                folder_id_by_rel[rel_sub] = fid

                if level == 2 and dn in STANDARD_DIMS:
                    subs_dict[dn] = {"id": fid, "driveUrl": f_drive}
                else:
                    parent_rel = "/".join(parts[:-1])
                    parent_id = folder_id_by_rel.get(parent_rel, comp_id)
                    extra_subfolders.append({
                        "name": dn,
                        "id": fid,
                        "parentId": parent_id,
                        "path": rel_sub,
                        "company": comp_folder,
                        "category": parts[1] if len(parts) >= 2 else "Company",
                        "level": level,
                        "driveUrl": f_drive,
                    })

        companies_list.append({
            "folder": comp_folder,
            "id": comp_id,
            "driveUrl": comp_drive,
            "label": comp_label,
            "ticker": comp_ticker,
            "sector": comp_sector,
            "reportCount": 0,
            "subs": subs_dict,
        })

    company_by_folder = {c["folder"]: c for c in companies_list}

    new_cache: dict[str, dict] = {}
    files_list: list[dict] = []
    fulltext_map: dict[str, str] = {}
    parsed_new_count = 0
    reused_cache_count = 0
    now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    portal_meta = known_files.get("index.html", {})
    files_list.append({
        "id": portal_meta.get("id") or "portal_index_html",
        "parentId": root_info["id"],
        "name": "index.html",
        "relPath": "index.html",
        "size": os.path.getsize(index_html_path),
        "modifiedTime": now_iso,
        "title": "数据中心 (index.html)",
        "company": "根目录",
        "companyLabel": "根目录 docs",
        "ticker": "INDEX",
        "sector": "门户与系统文件",
        "category": "Portal",
        "folderCategory": "Root",
        "fileType": "HTML",
        "level": 1,
        "driveUrl": portal_meta.get("driveUrl") or root_info.get("driveUrl", DEFAULT_ROOT_DRIVE_URL),
        "isPortalIndex": True,
        "headings": [],
        "summary": "网站全局全文搜索与自定义导航入口文件",
        "wordCount": 0,
    })

    all_discovered_paths: list[tuple[str, str]] = []
    for dp, dns, fns in os.walk(docs_dir):
        dns[:] = sorted(
            d for d in dns
            if not d.startswith(".") and d not in IGNORED_DIRS and not d.endswith("_files")
        )
        if os.path.abspath(dp) != os.path.abspath(docs_dir) and "index.html" in fns:
            idx_full = os.path.join(dp, "index.html")
            idx_rel = os.path.relpath(idx_full, docs_dir).replace(os.sep, "/")
            all_discovered_paths.append((idx_rel, idx_full))
            dns[:] = []
            continue
        for fn in sorted(fns):
            if fn.startswith(".") or fn.startswith("._") or fn.endswith(".tmp") or fn.lower() in IGNORED_NAMES:
                continue
            full_p = os.path.join(dp, fn)
            rel_p = os.path.relpath(full_p, docs_dir).replace(os.sep, "/")
            if rel_p == "index.html":
                continue
            all_discovered_paths.append((rel_p, full_p))

    all_discovered_paths.sort(key=lambda x: x[0].lower())

    for rel_p, full_p in all_discovered_paths:
        try:
            st = os.stat(full_p)
        except OSError:
            continue

        size = st.st_size
        mtime_ns = get_mtime_ns(st)
        cached = old_cache.get(rel_p)
        is_unchanged = bool(
            cached
            and cached.get("size") == size
            and cached.get("mtime_ns") == mtime_ns
            and isinstance(cached.get("payload"), dict)
        )

        if is_unchanged:
            payload = cached["payload"]
            reused_cache_count += 1
        else:
            payload = extract_file_payload(full_p, os.path.basename(full_p), max_chars=max_chars)
            parsed_new_count += 1

        f_known = known_files.get(rel_p, {})
        file_mtime_iso = datetime.fromtimestamp(st.st_mtime, tz=timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        if is_unchanged and cached.get("modifiedTime"):
            mtime_iso = str(cached["modifiedTime"])
        elif not is_unchanged and has_prior_cache:
            mtime_iso = file_mtime_iso
        else:
            mtime_iso = f_known.get("modifiedTime") or file_mtime_iso

        new_cache[rel_p] = {
            "size": size,
            "mtime_ns": mtime_ns,
            "modifiedTime": mtime_iso,
            "payload": payload,
        }

        parts = rel_p.split("/")
        fn = parts[-1]
        _, ext = os.path.splitext(fn)
        ext_low = ext.lower()
        parent_rel = "/".join(parts[:-1])
        parent_id = folder_id_by_rel.get(parent_rel, root_info["id"])

        if len(parts) >= 2 and parts[0] in company_by_folder:
            comp_obj = company_by_folder[parts[0]]
            comp_name = comp_obj["folder"]
            comp_label = comp_obj["label"]
            comp_ticker = comp_obj["ticker"]
            comp_sector = comp_obj["sector"]
            comp_obj["reportCount"] += 1
        else:
            comp_name = "根目录"
            comp_label = "根目录直传文件"
            comp_ticker = "ROOT"
            comp_sector = "未分组文档"

        category, folder_cat = infer_category(parts, ext_low)
        file_id = f_known.get("id") or stable_id("file", rel_p)
        drive_url = f_known.get("driveUrl") or (
            company_by_folder[comp_name]["driveUrl"]
            if comp_name in company_by_folder
            else root_info.get("driveUrl", DEFAULT_ROOT_DRIVE_URL)
        )

        record = {
            "id": file_id,
            "parentId": parent_id,
            "name": fn,
            "relPath": rel_p,
            "size": size,
            "modifiedTime": mtime_iso,
            "title": payload["title"],
            "company": comp_name,
            "companyLabel": comp_label,
            "ticker": comp_ticker,
            "sector": comp_sector,
            "category": category,
            "folderCategory": folder_cat,
            "fileType": payload.get("fileType", "HTML"),
            "level": len(parts),
            "driveUrl": drive_url,
            "isPortalIndex": False,
            "headings": payload["headings"],
            "headingsEn": payload.get("headingsEn") or [zh_to_en_portal_text(h) for h in payload["headings"]],
            "titleEn": payload.get("titleEn") or zh_to_en_portal_text(title),
            "summary": payload["summary"],
            "summaryEn": payload.get("summaryEn") or zh_to_en_portal_text(payload["summary"]),
            "wordCount": payload.get("wordCount", 0),
        }
        files_list.append(record)
        if payload.get("fullText"):
            fulltext_map[rel_p] = str(payload["fullText"])

    content_files = [f for f in files_list if not f.get("isPortalIndex") and not f.get("isSystem")]

    # 自动同步全站 HTML 详情页（关联导航 + 与首页一致的浅色/深色切换 + 中英文切换）
    detail_checked = 0
    detail_updated = 0
    if sync_details:
        files_by_rel = {f["relPath"]: f for f in content_files if f.get("relPath")}
        subfolder_matrix = build_subfolder_matrix_data(docs_dir, files_by_rel)
        sub_by_comp: dict[str, list[dict]] = {}
        for _it in subfolder_matrix:
            sub_by_comp.setdefault(_it['company'], []).append(_it)
        detail_checked, detail_updated = sync_all_detail_pages(
            subfolder_by_comp=sub_by_comp,
            docs_dir=docs_dir,
            content_files=content_files,
            new_cache=new_cache,
        )

    cat_counts: dict[str, int] = {}
    sector_counts: dict[str, int] = {}
    type_counts: dict[str, int] = {}
    report_bytes = 0
    total_words = 0

    for f in content_files:
        cat_counts[f["category"]] = cat_counts.get(f["category"], 0) + 1
        sector_counts[f["sector"]] = sector_counts.get(f["sector"], 0) + 1
        ft = f.get("fileType", "FILE")
        type_counts[ft] = type_counts.get(ft, 0) + 1
        report_bytes += int(f.get("size") or 0)
        total_words += int(f.get("wordCount") or 0)

    active_companies = sum(1 for c in companies_list if c["reportCount"] > 0)

    catalog = {
        "updatedAt": now_iso,
        "root": root_info,
        "companies": companies_list,
        "extraSubfolders": extra_subfolders,
        "files": files_list,
        "systemFiles": [],
        "stats": {
            "totalTopFolders": len(companies_list),
            "totalSubfolders": total_subfolders,
            "totalReports": len(content_files),
            "totalFiles": len(files_list),
            "activeCompanies": active_companies,
            "totalChars": total_words,
            "categoryCounts": cat_counts,
            "sectorCounts": sector_counts,
            "fileTypeCounts": type_counts,
            "reportBytes": report_bytes,
            "totalBytes": report_bytes,
        },
    }

    write_json_atomic(cache_path, new_cache, indent=None)

    ft_json = json.dumps(fulltext_map, ensure_ascii=False, separators=(",", ":")).replace("</script>", "<\\/script>")
    js_payload = (
        "/* Auto-generated by sync_docs_index.py — Full-Text Search Index */\n"
        f"window.DOCS_FULLTEXT = {ft_json};\n"
        "if (typeof window.onDocsFullTextReady === 'function') { window.onDocsFullTextReady(window.DOCS_FULLTEXT); }\n"
    )
    tmp_js = search_js_path + ".tmp"
    with open(tmp_js, "w", encoding="utf-8") as f:
        f.write(js_payload)
    os.replace(tmp_js, search_js_path)

    html_src = read_text_with_fallback(index_html_path)
    cat_json = json.dumps(catalog, ensure_ascii=False, separators=(",", ":")).replace("</script>", "<\\/script>")
    ver_tag = datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S")

    updated_html, catalog_ok = inject_catalog_and_nav_into_html(
        html_src=html_src,
        cat_json=cat_json,
        nav_links=nav_links,
        ver_tag=ver_tag,
    )
    if not catalog_ok:
        raise RuntimeError(
            f"无法向 {index_html_path} 注入目录数据：该文件不含数据中心门户结构 (DOCS_CATALOG)。\n"
            f"请确认 docs/index.html 未被单篇研报覆盖。"
        )

    files_by_rel = {f["relPath"]: f for f in content_files if f.get("relPath")}
    subfolder_matrix = build_subfolder_matrix_data(docs_dir, files_by_rel)
    matrix_script = '<script id="docsSubfolderMatrixData">window.DOCS_SUBFOLDER_MATRIX = ' + json.dumps(subfolder_matrix, ensure_ascii=False, separators=(",", ":")).replace("</script>", "<\\/script>") + ";</script>"
    updated_html = re.sub(r'\s*<script\s+id="docsSubfolderMatrixData">[\s\S]*?</script>', "", updated_html)
    updated_html = updated_html.replace("<!-- DOCS_CATALOG_END -->", "<!-- DOCS_CATALOG_END -->\n  " + matrix_script, 1)
    updated_html = ensure_index_preview_and_bilingual(updated_html)
    tmp_html = index_html_path + ".tmp"
    with open(tmp_html, "w", encoding="utf-8") as f:
        f.write(updated_html)
    os.replace(tmp_html, index_html_path)

    if standalone_out:
        compact_ft = {k: v[:1100] for k, v in fulltext_map.items()}
        compact_ft_json = json.dumps(compact_ft, ensure_ascii=False, separators=(",", ":")).replace("</script>", "<\\/script>")
        standalone_html = re.sub(
            r'\s*<script[^>]*(?:\bid=["\']docsSearchIndexScript["\']|\bsrc=["\'][^"\']*search_index\.js[^"\']*["\'])[^>]*>\s*</script>',
            "",
            updated_html,
            flags=re.IGNORECASE,
        )
        inline_ft_script = (
            "\n  <!-- FULLTEXT_INLINE_START -->\n"
            '  <script id="docsInlineFullText">\n'
            f"    window.DOCS_FULLTEXT = window.DOCS_FULLTEXT || {compact_ft_json};\n"
            "  </script>\n"
            "  <!-- FULLTEXT_INLINE_END -->"
        )
        standalone_html = standalone_html.replace(
            "<!-- DOCS_CATALOG_END -->",
            "<!-- DOCS_CATALOG_END -->" + inline_ft_script,
            1,
        )
        os.makedirs(os.path.dirname(os.path.abspath(standalone_out)), exist_ok=True)
        with open(standalone_out, "w", encoding="utf-8") as f:
            f.write(standalone_html)

    return {
        "companies": len(companies_list),
        "subfolders": total_subfolders,
        "files": len(content_files),
        "parsed_new": parsed_new_count,
        "reused_cache": reused_cache_count,
        "total_chars": total_words,
        "detail_checked": detail_checked,
        "detail_updated": detail_updated,
        "index_path": index_html_path,
    }


def build_snapshot(docs_dir: str | None = None) -> int:
    """向后兼容旧版 serve.py / build_snapshot.py 调用接口。"""
    target_docs = locate_docs_dir(docs_dir)
    stats = build_catalog_and_index(docs_dir=target_docs, rebuild=False)
    return int(stats["files"])


def infer_rel_path_from_url(url: str) -> str:
    u = url.strip()
    idx = u.find("/docs/")
    if u.startswith("file://") and idx != -1:
        return unquote(u[idx + 6 :])
    if u.startswith("./"):
        return unquote(u[2:])
    return ""


def add_nav_link(docs_dir: str, title: str, url: str, menu: str, desc: str, section: str, same_tab: bool) -> None:
    nav_path = os.path.join(docs_dir, "nav_links.json")
    links = load_json_safe(nav_path, [])
    if not isinstance(links, list):
        links = []
    new_item: dict[str, object] = {
        "id": stable_id("nav", f"{title}_{url}"),
        "menu": menu.strip() or "精选导航",
        "section": section.strip(),
        "title": title.strip(),
        "desc": desc.strip() or url.strip(),
        "url": url.strip(),
        "newTab": not same_tab,
    }
    rel_p = infer_rel_path_from_url(url)
    if rel_p:
        new_item["relPath"] = rel_p
    links.append(new_item)
    write_json_atomic(nav_path, links, indent=2)
    print(f"[导航已添加] [{new_item['menu']}] {title} -> {url}")


def compute_tree_fingerprint(docs_dir: str) -> str:
    h = hashlib.sha1()
    nav_p = os.path.join(docs_dir, "nav_links.json")
    if os.path.isfile(nav_p):
        try:
            st = os.stat(nav_p)
            h.update(f"NAV:{st.st_size}:{get_mtime_ns(st)}\n".encode("utf-8"))
        except OSError:
            pass

    for dp, dns, fns in os.walk(docs_dir):
        dns[:] = sorted(
            d for d in dns
            if not d.startswith(".") and d not in IGNORED_DIRS and not d.endswith("_files")
        )
        rel_d = os.path.relpath(dp, docs_dir).replace(os.sep, "/")
        h.update(f"D:{rel_d}\n".encode("utf-8"))
        for fn in sorted(fns):
            if fn.startswith(".") or fn.startswith("._") or fn.endswith(".tmp") or fn.lower() in IGNORED_NAMES or fn == "index.html":
                continue
            fp = os.path.join(dp, fn)
            try:
                st = os.stat(fp)
                h.update(f"F:{rel_d}/{fn}:{st.st_size}:{get_mtime_ns(st)}\n".encode("utf-8"))
            except OSError:
                continue
    return h.hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser(description="增量同步数据中心全文检索索引、详情页关联导航及主题/双语切换配置")
    parser.add_argument("--dir", help="指定 docs 目录（或其父目录）路径，默认自动识别")
    parser.add_argument("--rebuild", action="store_true", help="忽略增量缓存，强制全量重新解析")
    parser.add_argument("--watch", action="store_true", help="开启实时监听模式，上传/拖入新文件自动更新索引与关联导航")
    parser.add_argument("--zip", nargs="?", const="an_mui_portal.zip", default=None, help="同步完成后将整个门户目录打包为 ZIP 压缩包")
    parser.add_argument("--open", action="store_true", help="索引更新完成后自动在默认浏览器打开 docs/index.html")
    parser.add_argument("--no-sync-details", action="store_true", help="仅更新首页索引，跳过详情页导航与外观同步")
    parser.add_argument("--interval", type=float, default=2.0, help="监听轮询间隔秒数（默认 2.0s）")
    parser.add_argument("--max-chars", type=int, default=250000, help="单文件最大索引字符数（默认 250000）")
    parser.add_argument("--standalone-out", help="额外导出一份内联全文摘要的单文件 HTML 预览包")
    parser.add_argument("--add-link", nargs=2, metavar=("TITLE", "URL"), help="快捷添加一条抽屉导航链接")
    parser.add_argument("--menu", "--group", dest="menu", default="精选导航", help="配合 --add-link 指定导航分组名（默认：精选导航）")
    parser.add_argument("--desc", default="", help="配合 --add-link 指定导航项副标题说明")
    parser.add_argument("--section", default="", help="配合 --add-link 指定分栏小标题")
    parser.add_argument("--same-tab", action="store_true", help="配合 --add-link 设置在当前窗口打开")
    args = parser.parse_args()

    try:
        docs_dir = locate_docs_dir(args.dir)
        if args.add_link:
            add_nav_link(docs_dir, args.add_link[0], args.add_link[1], args.menu, args.desc, args.section, args.same_tab)

        stats = build_catalog_and_index(
            docs_dir=docs_dir,
            rebuild=args.rebuild,
            max_chars=args.max_chars,
            standalone_out=args.standalone_out,
            sync_details=not args.no_sync_details,
        )
        chip_audit = audit_detail_chip_wrapping(docs_dir)
        index_uri = Path(str(stats["index_path"])).resolve().as_uri()
        print(
            f"✅ 索引与详情页同步完成 | 目录: {docs_dir}\n"
            f"   顶层分类: {stats['companies']} 个 · 子目录: {stats['subfolders']} 个 · 可检索文件: {stats['files']} 份\n"
            f"   本次新解析: {stats['parsed_new']} 份 · 命中增量缓存: {stats['reused_cache']} 份 · 全文总字数: {stats['total_chars']:,} 字\n"
            f"   详情页联动: 已扫描 {stats['detail_checked']} 份详情页 (本次更新 {stats['detail_updated']} 份 · 站点关联导航与浅深色/中英文已与首页对齐)\n"
            f"   标签换行检查: 已扫描 {chip_audit['scanned']} 份详情页 · 修复内联不换行 {chip_audit['fixed_inline']} 处 · 缺失样式 {chip_audit['missing_opt']} 份\n"
            f"   主页地址: {index_uri}"
        )

        if args.zip:
            zip_target = Path(args.zip)
            if not zip_target.is_absolute():
                zip_target = Path.cwd() / zip_target
            out_zip = package_portal_zip(Path(docs_dir).parent, zip_target)
            print(f"📦 已打包完整门户归档: {out_zip}")

        if args.open:
            try:
                webbrowser.open(index_uri)
            except Exception:
                pass

        if args.watch:
            print(f"👀 正在实时监听文件夹变动（每 {args.interval:.1f}s 检测一次，按 Ctrl+C 停止）...")
            last_fp = compute_tree_fingerprint(docs_dir)
            try:
                while True:
                    time.sleep(max(0.5, args.interval))
                    cur_fp = compute_tree_fingerprint(docs_dir)
                    if cur_fp != last_fp:
                        st_now = build_catalog_and_index(
                            docs_dir=docs_dir,
                            rebuild=False,
                            max_chars=args.max_chars,
                            sync_details=not args.no_sync_details,
                        )
                        last_fp = compute_tree_fingerprint(docs_dir)
                        now_str = datetime.now().strftime("%H:%M:%S")
                        print(
                            f"[{now_str}] 🔄 检测到文件或导航变更 -> 已自动更新！当前文件总数: {st_now['files']} 份 "
                            f"(新增/变动解析 {st_now['parsed_new']} 份 · 更新详情页 {st_now['detail_updated']} 份)"
                        )
            except KeyboardInterrupt:
                print("\n已停止实时监听。")
    except Exception as exc:
        print(f"\n❌ 同步失败：{exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()