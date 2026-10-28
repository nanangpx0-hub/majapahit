<#
.SYNOPSIS
    Konversi dokumen DOCX menjadi PDF menggunakan Word COM Automation.

.DESCRIPTION
    Digunakan oleh create_report_susenas.py dan juga dapat dijalankan mandiri:
      powershell -NoProfile -ExecutionPolicy Bypass -File convert_to_pdf.ps1 `
          -DocxPath "C:\path\laporan.docx" -PdfPath "C:\path\laporan.pdf"

.PARAMETER DocxPath
    Path lengkap dokumen Word (.docx) yang akan dikonversi.
.PARAMETER PdfPath
    Path lengkap hasil keluaran PDF.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,

    [Parameter(Mandatory = $true)]
    [string]$PdfPath
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# wdFormatPDF = 17
$wdFormatPDF = 17

if (-not (Test-Path -LiteralPath $DocxPath)) {
    Write-Error "Berkas DOCX tidak ditemukan: $DocxPath"
    exit 1
}

$docxFull = (Resolve-Path -LiteralPath $DocxPath).Path
$pdfFull = [System.IO.Path]::GetFullPath($PdfPath)

$pdfDir = Split-Path -Parent $pdfFull
if ($pdfDir -and -not (Test-Path -LiteralPath $pdfDir)) {
    New-Item -ItemType Directory -Force -Path $pdfDir | Out-Null
}
if (Test-Path -LiteralPath $pdfFull) {
    Remove-Item -LiteralPath $pdfFull -Force
}

$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0          # wdAlertsNone
    $word.Options.SaveNormalPrompt = $false

    # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
    $doc = $word.Documents.Open($docxFull, $false, $true, $false)

    $doc.SaveAs2($pdfFull, $wdFormatPDF)
    $doc.Close($false)
    $doc = $null
}
catch {
    Write-Error ("Konversi gagal: " + $_.Exception.Message)
    exit 2
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close($false) } catch { }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
    }
    if ($word -ne $null) {
        try { $word.Quit() } catch { }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if (-not (Test-Path -LiteralPath $pdfFull)) {
    Write-Error "PDF tidak terbentuk: $pdfFull"
    exit 3
}

$size = (Get-Item -LiteralPath $pdfFull).Length
Write-Output ("PDF dibuat: {0} ({1:N0} bytes)" -f $pdfFull, $size)
exit 0
