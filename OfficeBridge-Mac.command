#!/bin/sh
# OfficeBridge uses only the built-in macOS automation runtime.
set -eu
OFFICEBRIDGE_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
printf '%s\n' 'OfficeBridge · 本机 Office 核对' '请选择文件；本机 Office 可能请求访问本次结果文件夹。' '完成后会自动打开结果页面。这个窗口可以保留。'
/usr/bin/osascript -l JavaScript "$OFFICEBRIDGE_DIR/scripts/office-check-mac.js" --bundle "$OFFICEBRIDGE_DIR" "$@"
