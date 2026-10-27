#!/usr/bin/env bash
# convert_pdf.sh — конвертация itog/*.docx в PDF через LibreOffice (soffice).
#
# Использование:
#   tools/convert_pdf.sh                 # все docx из itog/
#   tools/convert_pdf.sh file.docx       # один файл из itog/
#
# Требуется soffice на хосте. Fallback (см. README):
#   docker run --rm -v "$PWD/itog:/work" ghcr.io/linuxserver/libreoffice \
#       soffice --headless --convert-to pdf --outdir /work /work/Файл.docx
set -euo pipefail

DIR="$(cd "$(dirname "$0")/.." && pwd)/itog"

if [ "$#" -ge 1 ]; then
  FILES=()
  for f in "$@"; do
    FILES+=("$DIR/$f")
  done
else
  FILES=("$DIR"/*.docx)
fi

for f in "${FILES[@]}"; do
  [ -f "$f" ] || { echo "Нет файла: $f" >&2; exit 1; }
  echo "PDF: $(basename "$f")"
  soffice --headless --convert-to pdf --outdir "$DIR" "$f"
done

echo "Готово: $DIR"
