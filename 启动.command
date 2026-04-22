#!/bin/bash
# 航线载量分析 — 一键启动
cd "$(dirname "$0")"
PY=python3
if ! $PY -c "import streamlit, docx, pandas" 2>/dev/null; then
  echo "首次运行，安装依赖…"
  $PY -m pip install --user -r requirements.txt
fi
exec $PY -m streamlit run app.py
