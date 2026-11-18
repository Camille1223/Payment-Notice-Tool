#!/bin/bash
set -e
cd "$(dirname "$0")"
if [ ! -x .venv/bin/python ]; then
  python3 -m venv .venv
  .venv/bin/python -m pip install -r requirements.txt
fi
export PORT="${PORT:-5050}"
echo "浏览器地址：http://127.0.0.1:${PORT}"
echo "生成 PDF 时如提示控制 Microsoft Word，请允许。关闭本窗口或按 Ctrl+C 可停止。"
if curl --silent --fail --max-time 2 "http://127.0.0.1:${PORT}/health" >/dev/null; then
  open "http://127.0.0.1:${PORT}"
  echo "工具已在运行。"
  exit 0
fi
export OPEN_BROWSER=1
exec .venv/bin/python app.py
