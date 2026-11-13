---
name: pdf-anti-copy
description: Use when a PDF must not be selectable, copyable, or searchable. Converts DOCX to image-only PDF via MS Word (DOCX→PDF→JPG→PDF) with reusable Windows scripts.
version: 1.1.0
author: Gege Desembri (GegeDevs), Hermes Agent
license: MIT
platforms: [windows]
metadata:
  hermes:
    tags: [pdf, docx, anti-copy, raster, image-pdf, word]
    related_skills: [pdf, docx, universal-pdf-reader]
---

# PDF Anti-Copy (Image-Only PDF, MS Word)

Turn documents into PDFs whose text cannot be blocked, copied, searched, or extracted — every page becomes one JPEG image and the text layer disappears. Use when the user asks for "PDF yang tidak bisa dicopy", "anti-copy", "jadikan gambar", "kunci teks", or needs a form/attachment that must not be lifted as text.

Windows only, MS Word only. Word must do the DOCX→PDF step: LibreOffice re-flows complex `.docx` (tables, headers, merged cells) and the structure falls apart — that is why the LibreOffice path was removed. Do not reintroduce it for layout-sensitive documents.

Warn once before starting: rasterising kills text selection, search, and screen-reader access by design.

## What it produces

```
PDF/<nama>.pdf          PDF teks dari Word — disimpan, satu-satunya versi yang bisa dicari
PDF-GAMBAR/<nama>.pdf   PDF gambar, nama file sama, teks tidak bisa diblok
```

Both folders sit next to the script, i.e. in the folder that holds the `.docx` files. Re-running overwrites, so nothing piles up. Keep the output name identical to the source (no `(gambar)` suffix) — the folder already separates the two versions.

## Usage

Copy `scripts/docx-to-image-pdf.vbs`, `scripts/image-pdf.ps1`, `scripts/run-windows.bat` into the folder with the `.docx` files, then double-click `run-windows.bat`.

```
run-windows.bat                 # semua .docx di folder itu
run-windows.bat /batas:1        # uji 1 file dulu (wajib untuk batch besar)
run-windows.bat /dpi:150 /q:80  # lebih cepat & kecil
```

Output per file: `[12/57] OK  FORM ....docx -- 9 hal, 1.72 MB, 3.4 s`. Failures print `GAGAL` plus a reason line and the batch keeps going.

Requirements: MS Word installed, Windows PowerShell 5.1 (default `powershell.exe`), optional Ghostscript (`gswin64c`) as rasteriser fallback. No Python, no PDF library, no extra install.

## Progress and diagnostics (the user watches the console)

Long batches must show where they are at every moment. A run that hangs while printing nothing cannot be diagnosed by the user, only guessed at.

- Print the stage BEFORE the blocking call, not after: `Menyiapkan Word...`, `Word siap (versi 16.0)`, `[Word] buka <file>`, `[Word] ekspor PDF...`, `render halaman n/N`, then the finished line with pages, MB, and seconds. A message printed only after a 5-minute hang leaves the user staring at a header.
- Run one file first (`/batas:1`), confirm the `OK` line, then the batch. Never start 50+ files on an untested path.
- When a batch stalls with no output, bisect with a throwaway probe that prints timestamps around each Word call (`start` → create object → open one document → export → quit). Do not add retries or timeouts first: a blocked COM call never yields, so retries only multiply the wait. Task Manager showing CPU use says a process is working, not which step it is in.
- `run-windows.bat` must call `cscript //nologo "%~dp0docx-to-image-pdf.vbs" %*` and end with `pause`. `%*` is what lets `/dpi:150` and `/batas:1` pass through without editing the script; the file needs CRLF line endings.

## Verify the result really is image-only

1. Buka `PDF-GAMBAR/<nama>.pdf` di Adobe/Edge, coba blok satu paragraf — tidak ada yang tersorot.
2. Ukuran wajar: ± 0.2 MB per halaman A4 @200 dpi (teks padat ± 0.35 MB, halaman kosong ± 0.15 MB).
3. Kalau poppler tersedia (`pdfinfo`, `pdfimages`, `pdftotext`), ini bukti paling kuat:

```bash
pdfinfo out.pdf | grep -E 'Pages|Page size'   # jumlah halaman & ukuran sama dengan aslinya
pdfimages -list out.pdf                        # 1 gambar JPEG per halaman
pdftotext out.pdf - | wc -c                    # 0-3 byte = lapisan teks hilang
```

## Rules of thumb

- 150 dpi cukup untuk baca layar; 200 dpi untuk cetak; 300 dpi hanya kalau detail scan penting.
- JPEG quality 80-85. Di atas 90 file membengkak tanpa tambahan yang terlihat.
- Jangan hapus `PDF/` kecuali diminta: itu satu-satunya salinan yang bisa dicari.
- Timing: Word export ± 1.7 s/file, raster ± 0.2-0.4 s/halaman → 57 dokumen/169 halaman ± 3-4 menit.

## Pitfalls

Baca `references/pitfalls.md` sebelum mengubah skrip: Word COM menggantung di PowerShell (harus lewat VBS), aturan objek/xref PDF yang ditulis tangan, `DCTDecode`, dan kontrak exit code VBS↔PowerShell.