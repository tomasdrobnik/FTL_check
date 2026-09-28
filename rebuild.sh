#!/bin/bash
# One-command rebuild: merge latest xlsx exports, regenerate index.html, commit and push.
set -e
cd "$(dirname "$0")"

if [ ! -d venv ]; then
  python3 -m venv venv
  ./venv/bin/pip install --quiet openpyxl
fi

./venv/bin/python build.py

git add -A
if git diff --cached --quiet; then
  echo "Nothing changed."
else
  git commit -m "Rebuild $(date +%Y-%m-%d)"
  git push
fi
