# pdf-anti-copy

Hermes skill + standalone scripts: change DOCX/PDF into **image-only PDF** so the
text cannot be selected, copied, searched, or extracted.

```
DOCX ──Word/LibreOffice──> PDF ──WinRT/pdftoppm──> JPG ──own PDF writer──> image-only PDF
```

No paid tools, no Office add-ins, no PDF library: the final PDF is assembled by
hand (JPEG inside `DCTDecode` XObjects), so a Windows machine with Word — or a
Linux box with LibreOffice — is enough.

## Output

```
PDF/<name>.pdf          plain text PDF, kept for searching/archiving
PDF-GAMBAR/<name>.pdf   image-only PDF, same file name, different folder
```

## Windows (keeps MS Word layout)

1. Put `docx-to-image-pdf.vbs`, `image-pdf.ps1`, `run-windows.bat` in the folder
   with your `.docx` files.
2. Double-click `run-windows.bat`.

```
run-windows.bat                 # every .docx in that folder
run-windows.bat /batas:1        # smoke-test one file
run-windows.bat /dpi:150 /q:80  # smaller and faster
```

Requirements: MS Word installed, Windows PowerShell 5.1 (default), optional
Ghostscript as a fallback rasteriser.

## Linux / macOS

```bash
scripts/docx-to-image-pdf.sh *.docx          # DOCX -> PDF -> image PDF
DPI=150 scripts/docx-to-image-pdf.sh a.docx
scripts/pdf-to-image-pdf.py in.pdf -o out.pdf -d 200   # PDF only
```

Requirements: `libreoffice` (only for the DOCX step), `poppler-utils`
(`pdftoppm`, `pdfinfo`), `python3` standard library.

## Verify that the text is really gone

```bash
pdfinfo out.pdf | grep -E 'Pages|Page size'   # same page count/size as source
pdfimages -list out.pdf                        # one JPEG image per page
pdftotext out.pdf - | wc -c                    # 0-3 bytes = no text layer
pdftoppm -r 80 -png -f 1 -l 1 out.pdf cek      # visual check
```

Typical size at 200 dpi: ~0.2 MB per A4 page.

## Files

| File | Purpose |
| --- | --- |
| `scripts/docx-to-image-pdf.vbs` | Windows orchestrator: Word export + hand-off to PowerShell |
| `scripts/image-pdf.ps1` | PDF → JPG (WinRT, Ghostscript fallback) → image PDF |
| `scripts/run-windows.bat` | Double-click launcher, forwards `/dpi: /q: /batas:` |
| `scripts/docx-to-image-pdf.sh` | Linux/macOS DOCX → PDF → image PDF |
| `scripts/pdf-to-image-pdf.py` | Dependency-free image PDF writer (janitor for any PDF) |
| `references/pitfalls.md` | Word COM hang under PowerShell, PDF object/xref rules, VBS↔PS contract |

## Known limits

- Rasterising removes text selection, search, and screen-reader access by design.
- Scanned/form-heavy documents can look slightly softer at 150 dpi.
- LibreOffice output can differ from Word by a few points; check one page before
  delivering a print-critical document.

MIT licensed.