# pdf-anti-copy

Hermes skill + standalone Windows scripts: change `.docx` into **image-only PDF**
so the text cannot be selected, copied, searched, or extracted.

```
DOCX ──MS Word──> PDF ──WinRT/Ghostscript──> JPG ──own PDF writer──> image-only PDF
```

No paid tools, no Office add-ins, no PDF library: the final PDF is assembled by
hand (JPEG inside `DCTDecode` XObjects). A Windows machine with MS Word is all
you need.

Word is not optional: LibreOffice re-flows complex documents (tables, headers,
merged cells) and the layout breaks, so no LibreOffice/Linux path is shipped.

## Output

```
PDF/<name>.pdf          plain text PDF from Word, kept for searching/archiving
PDF-GAMBAR/<name>.pdf   image-only PDF, same file name, different folder
```

## Usage

1. Put `docx-to-image-pdf.vbs`, `image-pdf.ps1`, `run-windows.bat` in the folder
   with your `.docx` files.
2. Double-click `run-windows.bat`.

```
run-windows.bat                 # every .docx in that folder
run-windows.bat /batas:1        # smoke-test one file
run-windows.bat /dpi:150 /q:80  # smaller and faster
```

Requirements: MS Word, Windows PowerShell 5.1 (default), optional Ghostscript
(`gswin64c`) as a fallback rasteriser.

## Verify that the text is really gone

1. Open `PDF-GAMBAR/<name>.pdf`, try to select a paragraph — nothing highlights.
2. Size sanity: ~0.2 MB per A4 page at 200 dpi.
3. With poppler available, this is the hard proof:

```bash
pdfinfo out.pdf | grep -E 'Pages|Page size'   # same page count/size as source
pdfimages -list out.pdf                        # one JPEG image per page
pdftotext out.pdf - | wc -c                    # 0-3 bytes = no text layer
```

## Files

| File | Purpose |
| --- | --- |
| `scripts/docx-to-image-pdf.vbs` | Orchestrator: Word export per file, hand-off to PowerShell |
| `scripts/image-pdf.ps1` | PDF → JPG (WinRT, Ghostscript fallback) → image PDF |
| `scripts/run-windows.bat` | Double-click launcher, forwards `/dpi: /q: /batas:` |
| `references/pitfalls.md` | Word COM hang under PowerShell, PDF object/xref rules, VBS↔PS contract |

## Known limits

- Rasterising removes text selection, search, and screen-reader access by design.
- Scanned/form-heavy documents look slightly softer at 150 dpi.
- Word COM can hang when driven from PowerShell; the scripts drive Word from VBS
  for that reason. Do not "fix" it with retries.

## Verified

Full batch (57 documents, 0 failures) on Windows 10 Pro 22H2 build 19045.2311 with
Microsoft Office LTSC Professional Plus 2021, driven by `run-windows.bat`.
That is the known-good baseline; other Office builds may rasterise differently.

MIT licensed.