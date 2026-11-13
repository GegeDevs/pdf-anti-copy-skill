#!/usr/bin/env bash
# DOCX -> PDF (LibreOffice) -> PDF gambar (teks tidak bisa diblok/copy).
#
#   ./docx-to-image-pdf.sh file.docx [file2.docx ...]
#   DPI=150 ./docx-to-image-pdf.sh *.docx
#
# Hasil: ./PDF/<nama>.pdf (teks) dan ./PDF-GAMBAR/<nama>.pdf (gambar).
# Butuh: libreoffice (soffice), poppler-utils, python3.
set -euo pipefail

dpi="${DPI:-200}"
here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
pdf_dir="$PWD/PDF"
img_dir="$PWD/PDF-GAMBAR"
mkdir -p "$pdf_dir" "$img_dir"

[ "$#" -ge 1 ] || { echo "pemakaian: $0 file.docx [...]"; exit 2; }

gagal=0
for docx in "$@"; do
  base="$(basename "${docx%.*}")"
  tmp="$(mktemp -d)"
  if ! soffice --headless --convert-to pdf --outdir "$tmp" "$docx" >/dev/null 2>&1; then
    echo "GAGAL $base -- LibreOffice tidak bisa mengonversi"; gagal=$((gagal + 1)); rm -rf "$tmp"; continue
  fi
  mv "$tmp/$base.pdf" "$pdf_dir/$base.pdf"
  rm -rf "$tmp"
  if python3 "$here/pdf-to-image-pdf.py" "$pdf_dir/$base.pdf" -d "$dpi" -o "$img_dir/$base.pdf"; then :; else
    gagal=$((gagal + 1))
  fi
done

echo
echo "PDF teks   : $pdf_dir"
echo "PDF gambar : $img_dir"
[ "$gagal" -eq 0 ] || exit 1