# Pitfalls — DOCX/PDF → image-only PDF

## 1. Word COM hangs under PowerShell, works under VBS

Observed on a VirtualBox Windows VM (Office 16):

- PowerShell `New-Object -ComObject Word.Application` succeeds (`$word.Version` prints).
- The next call (`Documents.Add` / `Documents.Open`) never returns: no output, CPU busy, minutes to hours.
- The same machine runs the older VBS (`cscript`, same `Documents.Open` + `ExportAsFixedFormat` arguments) at ~1.7 s per document, 57 documents in 97 s.

Conclusion: drive Word from VBS and hand off only the raster/assembly work to PowerShell. Do not "fix" this by adding timeouts inside PowerShell — the COM call never yields.

Diagnostics that separate the two halves:

```powershell
# tulis di .bat supaya kelihatan sampai mana jalan
powershell -NoProfile -Command "$w=New-Object -ComObject Word.Application; Write-Host 'Word siap'"
```

Silence after `Word siap` = hang on document open/export, not on COM activation.

## 2. Rasterisation on Windows without extra installs

Primary: WinRT `Windows.Data.Pdf.PdfDocument` + `PdfPage.RenderToStreamAsync`. Requires `Add-Type -AssemblyName System.Runtime.WindowsRuntime`, WinRT type loading via `[Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]`, and an `AsTask` bridge:

```powershell
$asTaskOp = @([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
```

Notes:

- Create WinRT objects with `::new()`, not `New-Object` (constructor not discoverable).
- `RenderToStreamAsync` emits PNG; re-encode to JPEG with `System.Drawing` (`EncoderParameters` quality) to keep files small.
- WinRT is available in Windows PowerShell 5.1, NOT in PowerShell 7 — launch with `powershell.exe`.
- Set `DestinationWidth` from `page.Size.Width / 72 * dpi`; the height follows the aspect ratio.
- Fallback: `gswin64c -sDEVICE=jpeg -r 200 -dJPEGQ=85`. If neither works, fail loudly with one line; never emit a text PDF silently.

## 3. Hand-written PDF (no dependencies)

Layout used by both scripts: catalog → pages → per page `{page, image XObject, content stream}`.

```
k   = 3 + i*3   image XObject: /Subtype /Image /Filter /DCTDecode /Width /Height /ColorSpace /DeviceRGB
k+1 = 4 + i*3   page: /MediaBox [0 0 Wpt Hpt] /Resources << /XObject << /Im0 k 0 R >> >> /Contents k+2 0 R
k+2 = 5 + i*3   content: "q Wpt 0 0 Hpt 0 0 cm /Im0 Do Q"
```

Rules that break viewers when violated:

- `MediaBox` is in points = pixels × 72 / dpi. A4 at 200 dpi is 1654×2339 px → 595.44×842.04 pt. Using pixel counts as points gives an enormous page.
- `/Length` must be the exact byte count of the stream, and the stream must be separated from the dictionary by `stream\n` … `\nendstream`.
- xref offsets are byte offsets of each `N 0 obj`. Write them as `%010d 00000 n` with a trailing space; the free entry is `0000000000 65535 f`.
- Keep JPEGs as-is with `DCTDecode`; do not re-deflate them.
- In PowerShell read/write JPEG bytes with `[System.IO.File]::ReadAllBytes` / `MemoryStream`, not `Set-Content` (text encoders corrupt binary).

## 4. VBS ↔ PowerShell contract

`docx-to-image-pdf.vbs` runs `sh.Exec(cmd)` per file and reads child stdout:

- Child prints `PAGES=<n>` on success, `ERROR: ...` on failure, and exits non-zero on failure.
- Read with `StdOut.ReadLine` until `AtEndOfStream`. `ReadAll` deadlocks on some hosts.
- Do not rely on `\r` progress lines from the child: `ReadLine` returns only at newline, so a self-overwriting counter looks frozen. Print per-page progress in the child only when the child owns the console.
- After the streams end, poll `ex.Status` until it stops being `0` before reading `ex.ExitCode`.
- Exit code 1 from `exit 1` inside `catch` propagates correctly (verified).

## 5. LibreOffice path (host)

`soffice --headless --convert-to pdf --outdir <dir> file.docx` then `pdf-to-image-pdf.py`. Fast and headless, but fonts/spacing can differ from Word; check one page if the document is print-critical. `pdf-to-image-pdf.py` reads page size from `pdfinfo` so the output keeps the original page dimensions instead of assuming A4.

Verify with the commands in SKILL.md; `pdftotext` returning 0-3 bytes is the actual proof that the text layer is gone.