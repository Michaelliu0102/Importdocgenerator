#!/usr/bin/env bash
set -euo pipefail

if [[ $# -ne 1 ]]; then
  echo "用法: $0 /path/to/AppName.app"
  exit 1
fi

APP_PATH="$1"
BIN_NAME="$(basename "$APP_PATH" .app)"
BIN_PATH="$APP_PATH/Contents/MacOS/$BIN_NAME"

if [[ ! -f "$BIN_PATH" ]]; then
  echo "错误: 未找到可执行文件: $BIN_PATH"
  exit 1
fi

ARCHES="$(lipo -archs "$BIN_PATH" 2>/dev/null || true)"
if [[ -z "$ARCHES" ]]; then
  ARCHES="$(file "$BIN_PATH")"
fi

MINOS="$(
  otool -l "$BIN_PATH" 2>/dev/null |
    awk '
      $1 == "cmd" && $2 == "LC_BUILD_VERSION" { in_build = 1; next }
      in_build && $1 == "minos" { print $2; exit }
    '
)"

echo "App: $APP_PATH"
echo "Binary: $BIN_PATH"
echo "Architectures: $ARCHES"
if [[ -n "$MINOS" ]]; then
  echo "Minimum macOS: $MINOS"
fi

if [[ "$ARCHES" == *"x86_64"* ]]; then
  echo "Intel Mac: supported"
else
  echo "Intel Mac: not supported"
fi

if [[ "$ARCHES" == *"arm64"* ]]; then
  echo "Apple Silicon Mac: supported"
fi
