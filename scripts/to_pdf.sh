#!/usr/bin/env bash
# docx -> pdf via WPS wpscli (macOS).
#   bash to_pdf.sh <input.docx> [output.pdf]
# WPS 沙盒只能读写用户可见目录(桌面/文档),不能碰 /tmp 与任意路径,
# 故把输入复制到 ~/Documents 中转,产物搬回目标位置后清理。
set -euo pipefail

IN="${1:?用法: to_pdf.sh <input.docx> [output.pdf]}"
OUT="${2:-${IN%.docx}.pdf}"
WPS="/Applications/wpsoffice.app/Contents/MacOS/wpscli"

if [ ! -f "$IN" ]; then echo "输入不存在: $IN" >&2; exit 1; fi
if [ ! -x "$WPS" ]; then
  echo "未找到 wpscli: $WPS (假定 macOS + WPS;有 LibreOffice 可改用 soffice --headless --convert-to pdf)" >&2
  exit 1
fi

IN_ABS="$(cd "$(dirname "$IN")" && pwd)/$(basename "$IN")"
OUT_ABS="$(cd "$(dirname "$OUT")" && pwd)/$(basename "$OUT")"
STAGE="$HOME/Documents/.gongwen-wps-tmp-$$"
mkdir -p "$STAGE"
trap 'rm -rf "$STAGE"' EXIT

cp "$IN_ABS" "$STAGE/in.docx"
NO_PROXY="*" "$WPS" word2pdf "$STAGE/in.docx" -o "$STAGE/out.pdf" >/dev/null
if [ ! -f "$STAGE/out.pdf" ]; then echo "wpscli 未产出 PDF,请检查 WPS 是否可用" >&2; exit 1; fi
mv "$STAGE/out.pdf" "$OUT_ABS"
echo "PDF: $OUT_ABS"
