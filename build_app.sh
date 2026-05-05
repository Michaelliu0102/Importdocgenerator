#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

if [[ ! -d ".venv" ]]; then
  echo "创建虚拟环境 .venv ..."
  python3 -m venv .venv
fi

source .venv/bin/activate

python -m pip uninstall -y tkinterdnd2 || true
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install --upgrade pyinstaller

python - <<'PY'
from pathlib import Path
import shutil
import site

vendor = Path("vendor/tkdnd/osx-arm64/libtcl9tkdnd2.9.5.dylib").resolve()
if not vendor.exists():
    raise SystemExit(0)

site_packages = Path(site.getsitepackages()[0])
target_dir = site_packages / "tkinterdnd2" / "tkdnd" / "osx-arm64"
if not target_dir.exists():
    raise SystemExit(0)

target = target_dir / vendor.name
shutil.copy2(vendor, target)

legacy = target_dir / "libtkdnd2.9.3.dylib"
if legacy.exists():
    legacy.rename(target_dir / "libtkdnd2.9.3.dylib.bak")
PY

pyinstaller \
  --noconfirm \
  --windowed \
  --name "ClearanceOS" \
  --add-data "templates:templates" \
  --add-data "export_templates:export_templates" \
  --add-data "data:data" \
  gui_app.py

echo
echo "打包完成: dist/ClearanceOS.app"
echo "使用环境: $(pwd)/.venv"
