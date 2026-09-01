# Coverage Debt at Scale - step 0 of the rebuild
# ==============================================
# ieeeconf_letter.dot is a binary Word 97-2003 template (OLE2), so python-docx
# cannot open it. Word can. Documents.Add(<template>) gives a new document that
# already inherits the template's styles.xml, numbering.xml, settings.xml,
# fontTable.xml, header1.xml and sectPr -- which is the whole point: the paper
# is built FROM the template, not styled to look like it.
#
# Run once from this folder, then:
#   python build_paper.py
#   python verify_format.py
#
# Needs Word installed. template_probe.dotx is only used by verify_format.py,
# to read the reference values back out of the template instead of trusting
# hard-coded numbers.

param([string]$Template = "$PSScriptRoot\ieeeconf_letter.dot")

if (-not (Test-Path $Template)) { throw "template not found: $Template" }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$doc = $word.Documents.Add($Template, $false, 0, $false)
$doc.SaveAs2("$PSScriptRoot\base.docx", 12)              # wdFormatXMLDocument
$doc.Close($false)

$probe = $word.Documents.Open($Template, $false, $true, $false)
$probe.SaveAs2("$PSScriptRoot\template_probe.dotx", 14)  # wdFormatXMLTemplate
$probe.Close($false)

$word.Quit()
Write-Output "wrote base.docx and template_probe.dotx"
