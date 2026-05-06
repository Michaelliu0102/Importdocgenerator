#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

# Prefer Homebrew Python, then fall back to system python if needed.
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
    /usr/bin/python3 \
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

PY_BIN="$(pick_python_with_tk || true)"
if [[ -z "${PY_BIN:-}" ]]; then
  echo "错误: 未找到带 Tk 8.6+ 的 Python。"
  echo "Apple Silicon 可安装 /opt/homebrew/bin/python3.13，Intel Mac 可安装 /usr/local/bin/python3.11。"
  exit 1
fi

echo "使用 Python: $PY_BIN"
TK_VERSION="$("$PY_BIN" -c "import tkinter; print(tkinter.TkVersion)")"
echo "Tk 版本: $TK_VERSION"

# If .venv exists but was built with a different Python, recreate it
NEED_VENV=0
if [[ ! -d ".venv" ]]; then
  NEED_VENV=1
elif [[ ! -x ".venv/bin/python" ]]; then
  NEED_VENV=1
else
  VENV_TK=$(.venv/bin/python -c "import tkinter; print(tkinter.TkVersion)" 2>/dev/null || echo "0")
  if python3 -c "exit(0 if float('$VENV_TK') >= 8.6 else 1)" 2>/dev/null; then
    NEED_VENV=0
  else
    echo "当前 .venv 的 Tk 版本过旧 ($VENV_TK)，正在重建..."
    NEED_VENV=1
  fi
fi

if [[ "$NEED_VENV" == "1" ]]; then
  rm -rf .venv
  echo "正在用 $PY_BIN 创建虚拟环境..."
  "$PY_BIN" -m venv .venv
fi

source .venv/bin/activate

python -m pip install --upgrade pip -q
python -m pip uninstall -y tkinterdnd2 -q 2>/dev/null || true
python -m pip install -r requirements.txt -q

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

echo "启动 GUI..."
python gui_app.py
