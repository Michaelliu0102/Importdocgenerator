#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

pick_python_with_tk() {
  local py
  for py in \
    /usr/local/bin/python3.11 \
    /opt/homebrew/bin/python3.11 \
    /usr/local/bin/python3.12 \
    /opt/homebrew/bin/python3.12 \
    /usr/local/bin/python3.13 \
    /opt/homebrew/bin/python3.13 \
    /usr/local/bin/python3 \
    /opt/homebrew/bin/python3 \
    python3
  do
    if [[ -x "$py" ]]; then
      if "$py" - <<'PYEOF' >/dev/null 2>&1
import tkinter
assert tkinter.TkVersion >= 8.6, "Tk too old"
PYEOF
      then
        echo "$py"
        return 0
      fi
    fi
  done
  return 1
}

HOST_ARCH="$(uname -m)"
TARGET_ARCH="${TARGET_ARCH:-$HOST_ARCH}"

if [[ "$TARGET_ARCH" == "x86_64" && "$HOST_ARCH" != "x86_64" ]]; then
  echo "错误: 当前 shell 运行在 $HOST_ARCH，不能直接产出 Intel 版 app。"
  echo "请在 Intel Mac 上执行，或先进入 Rosetta x86_64 shell 后再运行本脚本。"
  exit 1
fi

PY_BIN="$(pick_python_with_tk || true)"
if [[ -z "${PY_BIN:-}" ]]; then
  echo "错误: 未找到带 Tk 8.6+ 的 Python 3。"
  echo "Intel Mac 常见路径是 /usr/local/bin/python3.11。"
  exit 1
fi

echo "宿主架构: $HOST_ARCH"
echo "目标架构: $TARGET_ARCH"
echo "使用 Python: $PY_BIN"
echo "Tk 版本: $("$PY_BIN" -c 'import tkinter; print(tkinter.TkVersion)')"
if [[ "$TARGET_ARCH" == "arm64" ]]; then
  echo "提示: 当前生成的是 Apple Silicon 版 app，不能发给 Intel Mac。"
fi

if [[ ! -d ".venv" ]]; then
  echo "创建虚拟环境 .venv ..."
  "$PY_BIN" -m venv .venv
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
  --paths "$(pwd)" \
  --hidden-import main \
  --hidden-import eur_invoice_standalone \
  --add-data "templates:templates" \
  --add-data "export_templates:export_templates" \
  --add-data "data:data" \
  gui_app.py

TOC_FILE="build/ClearanceOS/PYZ-00.toc"
if [[ ! -f "$TOC_FILE" ]] || ! grep -Fq "('main'," "$TOC_FILE"; then
  echo "错误: PyInstaller 产物中未找到 main 模块，构建已中止。"
  exit 1
fi
if ! grep -Fq "('eur_invoice_standalone'," "$TOC_FILE"; then
  echo "错误: PyInstaller 产物中未找到 eur_invoice_standalone 模块，构建已中止。"
  exit 1
fi

echo
echo "打包完成: dist/ClearanceOS.app"
echo "使用环境: $(pwd)/.venv"
