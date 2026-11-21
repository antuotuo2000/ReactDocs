#!/bin/bash
export PATH="/opt/homebrew/bin:/usr/local/bin:/Library/Frameworks/Python.framework/Versions/Current/bin:$PATH"
cd "$(dirname "$0")" || exit 1
xattr -dr com.apple.quarantine sync_docs_index.py docs 2>/dev/null || true

echo "=================================================="
echo "🚀 王安安数据中心 — 正在重建全站索引并发布到线上..."
echo "=================================================="

if command -v python3 >/dev/null 2>&1; then
  python3 sync_docs_index.py --dir ./docs
elif command -v python >/dev/null 2>&1; then
  python sync_docs_index.py --dir ./docs
else
  echo "❌ 未检测到 Python 3，请先安装 Python 3。"
  read -n 1 -s -r -p "按任意键关闭窗口..."
  exit 1
fi

touch docs/.nojekyll

if [ -d ".git" ]; then
  git add -A
  if git diff --cached --quiet; then
    echo "ℹ️ 当前没有检测到文件变更，正在检查远程推送状态..."
  else
    git commit -m "deploy: 更新王安安数据中心专著与全站索引 ($(date '+%Y-%m-%d %H:%M'))"
  fi
  BRANCH=$(git rev-parse --abbrev-ref HEAD 2>/dev/null || echo "main")
  git push -u origin "$BRANCH"
  echo ""
  echo "✅ 已成功推送到 GitHub！GitHub Pages 将在 30~60 秒内自动完成线上更新："
  echo "🌐 线上地址: https://antuotuo2000.github.io/ReactDocs/"
else
  echo "⚠️ 当前目录尚未关联 .git 仓库，请将文件复制到你的 ReactDocs 本地仓库目录后再次双击运行。"
fi

echo ""
read -n 1 -s -r -p "按任意键关闭窗口..."
