' Ekspor semua .docx -> PDF (Word) -> JPG -> PDF gambar (teks tidak bisa diblok/copy).
' Jalankan lewat run-windows.bat. Word diekspor dari VBS (cara ini terbukti jalan),
' raster + perakitan PDF dikerjakan image-pdf.ps1.
'
' Opsi:  run-windows.bat /dpi:150 /q:80 /batas:3
Option Explicit

If InStr(LCase(WScript.FullName), "cscript") = 0 Then
  CreateObject("WScript.Shell").Run "cscript //nologo """ & WScript.ScriptFullName & """", 1, False
  WScript.Quit
End If

Dim fso, sh, folder, imgDir, pdfDir, ps1, word, doc, f, a, kv
Dim total, idx, berhasil, gagal, log, t0, t1, dpi, q, batas
Dim tmpPdf, outPdf, cmd, ex, line, pages, rc, mb, base

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
imgDir = folder & "\PDF-GAMBAR"
pdfDir = folder & "\PDF"
ps1 = folder & "\image-pdf.ps1"

dpi = 200 : q = 85 : batas = 0
For Each a In WScript.Arguments
  kv = LCase(a)
  If Left(kv, 5) = "/dpi:" Then dpi = CLng(Mid(kv, 6))
  If Left(kv, 3) = "/q:" Then q = CLng(Mid(kv, 4))
  If Left(kv, 7) = "/batas:" Then batas = CLng(Mid(kv, 8))
Next

Sub Say(t)
  On Error Resume Next
  WScript.StdOut.WriteLine t
  If Err.Number <> 0 Then WScript.Echo t
  Err.Clear
End Sub

If Not fso.FileExists(ps1) Then
  Say "image-pdf.ps1 tidak ditemukan di " & folder
  WScript.Quit 1
End If
If Not fso.FolderExists(imgDir) Then fso.CreateFolder imgDir
If Not fso.FolderExists(pdfDir) Then fso.CreateFolder pdfDir

total = 0
For Each f In fso.GetFolder(folder).Files
  If LCase(fso.GetExtensionName(f.Name)) = "docx" Then total = total + 1
Next
If total = 0 Then
  Say "Tidak ada file .docx di folder ini."
  WScript.Quit 1
End If
If batas > 0 And batas < total Then total = batas

Say "Ekspor " & total & " file: DOCX -> PDF -> JPG -> PDF gambar (" & dpi & " dpi, JPEG q" & q & ")"
Say ""
Say "Menyiapkan Word..."

Set word = CreateObject("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
word.AutomationSecurity = 3
Say "Word siap (versi " & word.Version & ")."
Say ""

idx = 0 : berhasil = 0 : gagal = 0 : log = "" : t0 = Timer

For Each f In fso.GetFolder(folder).Files
  If LCase(fso.GetExtensionName(f.Name)) = "docx" Then
    If batas > 0 And idx >= batas Then Exit For
    idx = idx + 1
    t1 = Timer
    base = fso.GetBaseName(f.Name)
    tmpPdf = pdfDir & "\" & base & ".pdf"
    outPdf = imgDir & "\" & base & ".pdf"

    On Error Resume Next
    Set doc = word.Documents.Open(f.Path, False, True, False)
    If Err.Number <> 0 Then
      gagal = gagal + 1
      log = log & "  GAGAL BUKA   : " & f.Name & " -- " & Err.Description & vbCrLf
      Say "[" & idx & "/" & total & "] GAGAL  " & f.Name & " (tidak bisa dibuka)"
      Err.Clear
    Else
      doc.ExportAsFixedFormat tmpPdf, 17
      If Err.Number <> 0 Then
        gagal = gagal + 1
        log = log & "  GAGAL EKSPOR : " & f.Name & " -- " & Err.Description & vbCrLf
        Say "[" & idx & "/" & total & "] GAGAL  " & f.Name & " (ekspor PDF gagal)"
        Err.Clear
        doc.Close False
      Else
        doc.Close False
        pages = 0
        cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & ps1 & """ -Pdf """ & tmpPdf & """ -Out """ & outPdf & """ -Dpi " & dpi & " -Quality " & q
        Set ex = sh.Exec(cmd)
        Do While Not ex.StdOut.AtEndOfStream
          line = ex.StdOut.ReadLine
          If Left(line, 6) = "PAGES=" Then
            pages = CLng(Mid(line, 7))
          ElseIf Trim(line) <> "" Then
            Say "    " & line
          End If
        Loop
        Do While Not ex.StdOut.AtEndOfStream
          ex.StdOut.ReadLine
        Loop
        Do While Not ex.StdErr.AtEndOfStream
          line = ex.StdErr.ReadLine
          If Trim(line) <> "" Then Say "    " & line
        Loop
        Do While ex.Status = 0
          WScript.Sleep 100
        Loop
        rc = ex.ExitCode
        If Err.Number <> 0 Then rc = 1 : Err.Clear

        If rc = 0 And pages > 0 And fso.FileExists(outPdf) Then
          berhasil = berhasil + 1
          mb = FormatNumber(fso.GetFile(outPdf).Size / 1048576, 2)
          Say "[" & idx & "/" & total & "] OK     " & f.Name & " -- " & pages & " hal, " & mb & " MB, " & FormatNumber(Timer - t1, 1) & " s"
        Else
          gagal = gagal + 1
          log = log & "  GAGAL GAMBAR : " & f.Name & " -- keluar kode " & rc & vbCrLf
          Say "[" & idx & "/" & total & "] GAGAL  " & f.Name & " (gagal jadi gambar)"
        End If
      End If
    End If
    Err.Clear
  End If
Next

word.Quit
Say ""
Say "Selesai dalam " & FormatNumber(Timer - t0, 1) & " detik."
Say "Berhasil: " & berhasil & " file   Gagal: " & gagal & " file"
If log <> "" Then
  Say ""
  Say "Rincian gagal:"
  Say log
End If
Say ""
Say "PDF teks   ada di: " & pdfDir
Say "PDF gambar ada di: " & imgDir
