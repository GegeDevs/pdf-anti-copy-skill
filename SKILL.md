---
name: pdf-anti-copy
description: Use when a PDF must not be selectable, copyable, or searchable. Converts DOCX/PDF into image-only PDF (DOCX→PDF→JPG→PDF) with reusable Windows and Linux scripts.
version: 1.0.0
author: Gege Desembri (GegeDevs), Hermes Agent
license: MIT
platforms: [windows, linux, macos]
metadata:
  hermes:
    tags: [pdf, docx, anti-copy, raster, image-pdf, word, libreoffice]
    related_skills: [pdf, docx, universal-pdf-reader]
---

# PDF Anti-Copy (Image-Only PDF)

Turn documents into PDFs whose text cannot be blocked, copied, searched, or extracted — every page becomes one JPEG image and the text layer disappears. Use when the user asks for "PDF yang tidak bisa dicopy", "anti-copy", "jadikan gambar", "kunci teks", or needs a form/attachment that must not be lifted as text.

Do not use it when the recipient must select text, search, or run accessibility tools. Rasterising destroys all of that; say so once before proceeding.

## What it produces

```
PDF/<nama>.pdf          PDF teks biasa (dari Word/LibreOffice) — tetap disimpan
PDF-GAMBAR/<nama>.pdf   PDF gambar, nama file sama, teks tidak bisa diblok
```

Both files land next to the source documents (`PDF/` and `PDF-GAMBAR/` relative to the script). Re-running overwrites, so it does not pile up duplicates.

## Windows (Word fidelity — use when layout must match MS Word)

Copy `scripts/docx-to-image-pdf.vbs`, `scripts/image-pdf.ps1`, `scripts/run-windows.bat` into the folder that holds the `.docx` files, then double-click `run-windows.bat`.

```
run-windows.bat                 # semua .docx di folder itu
run-windows.bat /batas:1        # uji 1 file dulu
run-windows.bat /dpi:150 /q:80  # lebih cepat & kecil
```

Word is driven from VBS (`cscript`), rasterisation + PDF assembly from `image-pdf.ps1` (WinRT `Windows.Data.Pdf`, Ghostscript fallback). Never call Word from PowerShell — see `references/pitfalls.md`.

## Linux / macOS (no Word needed)

```bash
chmod +x scripts/*.sh scripts/*.py
scripts/docx-to-image-pdf.sh *.docx        # LibreOffice -> PDF -> PDF gambar
DPI=150 scripts/docx-to-image-pdf.sh a.docx
scripts/pdf-to-image-pdf.py in.pdf -o out.pdf -d 200   # PDF saja
```

One-off `docx2pdf`-free path: `soffice --headless --convert-to pdf --outdir . file.docx` then the Python script. Layout follows LibreOffice, so check one page against Word before delivering a print-critical document.

## Verify the result really is image-only

```bash
pdfinfo out.pdf | grep -E 'Pages|Page size'   # halaman & ukuran harus sama dengan aslinya
pdfimages -list out.pdf                        # harus 1 gambar JPEG per halaman
pdftotext out.pdf - | wc -c                    # harus 0-3 byte (hanya newline)
pdftoppm -r 80 -png -f 1 -l 1 out.pdf cek      # cek visual tidak gepeng
```

Size rule of thumb at 200 dpi: ~0.2 MB per A4 page. Text-heavy pages ~0.35 MB, mostly-empty pages ~0.15 MB.

## Rules of thumb

- 150 dpi is enough for screen reading; 200 dpi for print; 300 dpi only when scanned detail matters.
- Keep JPEG quality 80-85. Above 90 the file grows fast with no visible gain on text.
- Do not delete the text PDF unless the user asks: it is the only searchable copy left.
- Batch of 50+ files: run on the host (LibreOffice) when layout is not critical, on Windows when it is. Word export alone runs ~1.7 s/file.

## Pitfalls

Read `references/pitfalls.md` before changing the scripts: Word COM hanging under PowerShell, PDF objects/xref written by hand, `DCTDecode` requirements, and the VBS↔PowerShell exit-code contract.