# Ubah SATU file PDF menjadi PDF gambar (setiap halaman jadi JPEG).
# Dipanggil oleh EKSPOR-PDF-GAMBAR.vbs; boleh juga dijalankan manual.
#   powershell -File GAMBAR-PDF.ps1 -Pdf "in.pdf" -Out "out.pdf" -Dpi 200 -Quality 85
param(
    [Parameter(Mandatory = $true)][string]$Pdf,
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$Dpi = 200,
    [int]$Quality = 85
)

$ErrorActionPreference = 'Stop'
$tmp = Join-Path $env:TEMP ('gbr-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $tmp | Out-Null

function Save-Jpeg {
    param([System.IO.Stream]$Source, [string]$OutFile, [int]$Quality)
    Add-Type -AssemblyName System.Drawing
    $img = [System.Drawing.Image]::FromStream($Source)
    $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
    $ep = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, [long]$Quality)
    $img.Save($OutFile, $codec, $ep)
    $img.Dispose(); $ep.Dispose()
}

function Convert-PdfToJpg {
    param([string]$PdfPath, [string]$OutPrefix, [int]$Dpi = 200)
    $list = @()

    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime
        $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
        $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
        $null = [Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType = WindowsRuntime]
        $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]

        $asTaskOp = @([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
                $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
                $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
        $asTaskAction = @([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
                $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
                $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]

        $sf = $null; $doc = $null
        $t = $asTaskOp.MakeGenericMethod([Windows.Storage.StorageFile]).Invoke($null, @([Windows.Storage.StorageFile]::GetFileFromPathAsync($PdfPath)))
        $t.Wait(-1) | Out-Null; $sf = $t.Result
        $t = $asTaskOp.MakeGenericMethod([Windows.Data.Pdf.PdfDocument]).Invoke($null, @([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($sf)))
        $t.Wait(-1) | Out-Null; $doc = $t.Result

        for ($i = 0; $i -lt $doc.PageCount; $i++) {
            $page = $doc.GetPage($i)
            $rs = [Windows.Storage.Streams.InMemoryRandomAccessStream]::new()
            $ro = [Windows.Data.Pdf.PdfPageRenderOptions]::new()
            $ro.DestinationWidth = [uint32][math]::Round($page.Size.Width / 72.0 * $Dpi)
            $t = $asTaskAction.Invoke($null, @($page.RenderToStreamAsync($rs, $ro)))
            $t.Wait(-1) | Out-Null
            $rs.Seek(0)
            $net = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($rs)
            $buf = New-Object System.IO.MemoryStream
            $net.CopyTo($buf); $buf.Position = 0
            $out = '{0}_{1:D4}.jpg' -f $OutPrefix, ($i + 1)
            Save-Jpeg -Source $buf -OutFile $out -Quality $Quality
            $buf.Dispose(); $net.Dispose(); $rs.Dispose()
            $list += $out
        }
    }
    catch {
        Write-Host ('catatan: render WinRT gagal -- ' + $_.Exception.Message)
        $list = @()
    }

    if ($list.Count -eq 0) {
        $gs = Get-Command gswin64c.exe, gswin32c.exe, gs.exe -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($gs) {
            & $gs.Source -dNOPAUSE -dBATCH -dSAFER -sDEVICE=jpeg -r $Dpi -dJPEGQ=$Quality `
                -sOutputFile="$($OutPrefix)_%04d.jpg" $PdfPath | Out-Null
            $list = @(Get-ChildItem -Path "$($OutPrefix)_*.jpg" | Sort-Object Name | ForEach-Object { $_.FullName })
        }
    }
    if ($list.Count -eq 0) { throw 'PDF tidak bisa dirender (WinRT dan Ghostscript gagal).' }
    return $list
}

function New-ImagePdf {
    param([string[]]$JpgPaths, [string]$OutFile, [int]$Dpi = 200)

    Add-Type -AssemblyName System.Drawing
    $ms = New-Object System.IO.MemoryStream
    $bytes = [byte[]](0x25, 0x50, 0x44, 0x46, 0x2D, 0x31, 0x2E, 0x34, 0x0A, 0x25, 0xE2, 0xE3, 0xCF, 0xD3, 0x0A)
    $ms.Write($bytes, 0, $bytes.Length)

    $pages = @()
    foreach ($jpg in $JpgPaths) {
        $img = [System.Drawing.Image]::FromFile($jpg)
        $pages += [pscustomobject]@{
            Path = $jpg; W = $img.Width; H = $img.Height
            Wpt  = [math]::Round($img.Width * 72.0 / $Dpi, 2)
            Hpt  = [math]::Round($img.Height * 72.0 / $Dpi, 2)
        }
        $img.Dispose()
    }

    $kids = @()
    for ($i = 0; $i -lt $pages.Count; $i++) { $kids += ('{0} 0 R' -f (4 + $i * 3)) }

    $offsets = @{}
    $offsets[1] = $ms.Position
    $b = [System.Text.Encoding]::ASCII.GetBytes("1 0 obj`n<< /Type /Catalog /Pages 2 0 R >>`nendobj`n")
    $ms.Write($b, 0, $b.Length)

    $offsets[2] = $ms.Position
    $b = [System.Text.Encoding]::ASCII.GetBytes(("2 0 obj`n<< /Type /Pages /Kids [{0}] /Count {1} >>`nendobj`n" -f ($kids -join ' '), $pages.Count))
    $ms.Write($b, 0, $b.Length)

    for ($i = 0; $i -lt $pages.Count; $i++) {
        $p = $pages[$i]
        $imgObj = 3 + $i * 3; $pageObj = 4 + $i * 3; $contObj = 5 + $i * 3

        $offsets[$pageObj] = $ms.Position
        $b = [System.Text.Encoding]::ASCII.GetBytes(("{0} 0 obj`n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {1} {2}] /Resources << /XObject << /Im0 {3} 0 R >> >> /Contents {4} 0 R >>`nendobj`n" -f $pageObj, $p.Wpt, $p.Hpt, $imgObj, $contObj))
        $ms.Write($b, 0, $b.Length)

        $jb = [System.IO.File]::ReadAllBytes($p.Path)
        $offsets[$imgObj] = $ms.Position
        $b = [System.Text.Encoding]::ASCII.GetBytes(("{0} 0 obj`n<< /Type /XObject /Subtype /Image /Width {1} /Height {2} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length {3} >>`nstream`n" -f $imgObj, $p.W, $p.H, $jb.Length))
        $ms.Write($b, 0, $b.Length)
        $ms.Write($jb, 0, $jb.Length)
        $b = [System.Text.Encoding]::ASCII.GetBytes("`nendstream`nendobj`n")
        $ms.Write($b, 0, $b.Length)

        $content = "q {0} 0 0 {1} 0 0 cm /Im0 Do Q" -f $p.Wpt, $p.Hpt
        $offsets[$contObj] = $ms.Position
        $b = [System.Text.Encoding]::ASCII.GetBytes(("{0} 0 obj`n<< /Length {1} >>`nstream`n{2}`nendstream`nendobj`n" -f $contObj, $content.Length, $content))
        $ms.Write($b, 0, $b.Length)
    }

    $xrefPos = $ms.Position
    $size = $pages.Count * 3 + 3
    $xref = "xref`n0 $size`n0000000000 65535 f `n"
    for ($n = 1; $n -lt $size; $n++) { $xref += (('{0:D10} 00000 n ' -f $offsets[$n]) + "`n") }
    $b = [System.Text.Encoding]::ASCII.GetBytes($xref)
    $ms.Write($b, 0, $b.Length)
    $b = [System.Text.Encoding]::ASCII.GetBytes(("trailer`n<< /Size {0} /Root 1 0 R >>`nstartxref`n{1}`n%%EOF`n" -f $size, $xrefPos))
    $ms.Write($b, 0, $b.Length)

    [System.IO.File]::WriteAllBytes($OutFile, $ms.ToArray())
    $ms.Dispose()
    return $pages.Count
}

try {
    $jpgs = @(Convert-PdfToJpg -PdfPath $Pdf -OutPrefix (Join-Path $tmp 'p') -Dpi $Dpi)
    $n = New-ImagePdf -JpgPaths $jpgs -OutFile $Out -Dpi $Dpi
    Write-Host ("PAGES={0}" -f $n)
}
catch {
    Write-Host ('ERROR: ' + $_.Exception.Message)
    exit 1
}
finally {
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}