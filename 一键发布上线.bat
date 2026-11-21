@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 🚀 王安安数据中心 — 正在重建全站索引并发布到线上...
python3 sync_docs_index.py --dir ./docs 2>nul || python sync_docs_index.py --dir ./docs
type nul >> docs\.nojekyll
if exist .git (
  git add -A
  git commit -m "deploy: sync 王安安数据中心 portal & MUI reports"
  git push -u origin main
  echo ✅ 已推送至 GitHub Pages: https://antuotuo2000.github.io/ReactDocs/
) else (
  echo ⚠️ 当前目录未检测到 .git，请在 ReactDocs 仓库根目录运行。
)
pause
