# make_pdf.ps1 - convert the CV .docx to PDF with LibreOffice (uniform spacing + full font embedding).
# Word's COM export (docx2pdf / ExportAsFixedFormat) splits Cambria into several subsets and leaves
# some of them (plus ArialMT/SymbolMT) NOT embedded -> viewers substitute fonts -> uneven letter
# spacing. LibreOffice embeds every subset and keeps metric spacing (verified 2026-06-06, 2026-09-15).
# Usage:  powershell -File make_pdf.ps1                 (newest Jang_Gukhwa_CV_v*.docx, PDF next to it)
#         powershell -File make_pdf.ps1 -Publish        (... and copy to ../public/cv/Jang_Gukhwa_CV.pdf)
#         powershell -File make_pdf.ps1 Jang_Gukhwa_CV_v1.docx
param([string]$Docx = "", [switch]$Publish)

# clear conda PYTHON* vars so LibreOffice's bundled python doesn't warn
$env:PYTHONHOME = $null; $env:PYTHONPATH = $null

Set-Location $PSScriptRoot
if ($Docx -eq "") {
    $Docx = (Get-ChildItem "Jang_Gukhwa_CV_v*.docx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1).Name
    if ($null -eq $Docx) { throw "No Jang_Gukhwa_CV_v*.docx in $PSScriptRoot" }
}
$soffice = "C:\Program Files\LibreOffice\program\soffice.exe"
$full = (Resolve-Path $Docx).Path
$dir  = Split-Path -Parent $full
$pdf  = $full -replace '\.docx$', '.pdf'
& $soffice --headless --convert-to pdf --outdir $dir $full | Out-Null
if (-not (Test-Path $pdf)) { throw "LibreOffice produced no PDF for $Docx" }
Write-Output ("PDF -> " + $pdf + "  (" + [math]::Round((Get-Item $pdf).Length / 1KB) + " KB)")
if ($Publish) {
    $pub = Join-Path $PSScriptRoot "..\public\cv\Jang_Gukhwa_CV.pdf"
    Copy-Item $pdf $pub -Force
    Write-Output ("published -> " + (Resolve-Path $pub).Path)
}
