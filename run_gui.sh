#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

# Prefer Homebrew Python 3.14 (Tk 9.0), then 3.13, then system python
pick_python_with_tk() {
  for py in /opt/homebrew/bin/python3.14 /opt/homebrew/bin/python3.13 /opt/homebrew/bin/python3 /usr/bin/python3 python3; do
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
  echo "请运行: brew install python-tk@3.14"
  exit 1
fi

echo "使用 Python: $PY_BIN"
"$PY_BIN" -c "import tkinter; print(f'Tk 版本: {tkinter.TkVersion}')"

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
python -m pip install -r requirements.txt -q
python -m pip install -U tkinterdnd2 -q 2>/dev/null || echo "提示: tkinterdnd2 安装失败，拖拽功能不可用。"

echo "启动 GUI..."
python gui_app.py
