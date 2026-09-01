# Coverage Debt at Scale - repaginate and export the built .docx
# ==============================================================
# python-docx writes the XML but never lays the document out, so page count and
# float placement are unknown until Word paginates it. This opens the built
# file, repaginates, saves it back (so the .docx carries Word's layout cache)
# and exports a PDF with fonts embedded.
#
# Run after build_paper.py, before verify_format.py.

param(
    [string]$Docx = "$PSScriptRoot\Coverage_Debt_at_Scale_ieeeconf.docx",
    [string]$Pdf  = "$PSScriptRoot\Coverage_Debt_at_Scale_ieeeconf.pdf"
)

if (-not (Test-Path $Docx)) { throw "build it first: $Docx" }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$doc = $word.Documents.Open($Docx, $false, $false, $false)
$doc.Repaginate()
$pages = $doc.ComputeStatistics(2)     # wdStatisticPages
Write-Output "pages: $pages"

$doc.SaveAs2($Docx, 12)
# ExportAsFixedFormat(path, format=wdExportFormatPDF, openAfter, optimize,
#                     range, from, to, item, docProps, docStructureTags,
#                     bitmapMissingFonts, useISO19005_1 ...)
$doc.ExportAsFixedFormat($Pdf, 17, $false, 0, 0, 1, $pages, 0, $true, $true, 0, $true, $true, $false)
$doc.Close($false)
$word.Quit()

Write-Output "wrote $Pdf"
Write-Output "now run: python scrub_metadata.py"
if ($pages -gt 6) { Write-Warning "over 6 pages - RA-L charges for 7 and 8, and hard-stops at 8" }
