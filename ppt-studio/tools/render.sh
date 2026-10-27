#!/usr/bin/env bash
# .pptx → PDF → 슬라이드별 PNG (육안 QA용)
# 사용법: bash ppt-studio/tools/render.sh <deck.pptx> [출력폴더]
set -euo pipefail
deck="$1"
out="${2:-$(dirname "$deck")/render}"
mkdir -p "$out"
soffice --headless --convert-to pdf --outdir "$out" "$deck" >/dev/null 2>&1
pdf="$out/$(basename "${deck%.*}").pdf"
pdftoppm -png -r 60 "$pdf" "$out/slide"
echo "PDF: $pdf"
ls "$out"/slide-*.png
