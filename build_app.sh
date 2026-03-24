#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

if [[ ! -d ".venv" ]]; then
  echo "创建虚拟环境 .venv ..."
  python3 -m venv .venv
fi

source .venv/bin/activate

python -m pip install --upgrade pip
python -m pip install --upgrade pyinstaller

# Optional drag-and-drop support package (if install fails, GUI still works)
python -m pip install --upgrade tkinterdnd2 || true

pyinstaller \
  --noconfirm \
  --windowed \
  --name "报关资料生成器" \
  --add-data "templates:templates" \
  --add-data "data:data" \
  gui_app.py

echo
echo "打包完成: dist/报关资料生成器.app"
echo "使用环境: $(pwd)/.venv"
