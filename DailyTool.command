#!/bin/bash
cd "$(dirname "$0")"
if [ ! -d ".venv" ]; then
  python3 -m venv .venv
  .venv/bin/pip install -r requirements.txt
fi
SITE=$(.venv/bin/python -c "import site; print(site.getsitepackages()[0])")
if /usr/bin/python3 -c "import tkinter" >/dev/null 2>&1; then
  export PYTHONPATH="$PWD:$SITE${PYTHONPATH:+:$PYTHONPATH}"
  exec /usr/bin/python3 -m dailytool
fi
echo "Không mở được cửa sổ DailyTool trên máy này. Hãy dùng dist/DailyTool.exe trên Windows."
exit 1
