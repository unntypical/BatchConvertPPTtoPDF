
$scriptpath = $MyInvocation.MyCommand.Path
$curr_path = Split-Path $scriptpath

[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Powerpoint") > $null
[Reflection.Assembly]::LoadWithPartialname("Office") > $null
$ppt_app = New-Object "Microsoft.Office.Interop.Powerpoint.ApplicationClass" 

Get-ChildItem -Path $curr_path -Recurse -Filter *.ppt? | ForEach-Object {
    Write-Host "Processing" $_.FullName "..."
    $document = $ppt_app.Presentations.Open($_.FullName)
    $pdf_filename = "$($curr_path)\$($_.BaseName)_cvt.pdf"
    
    $exportPath           = $pdf_filename
    $fixedFormatType      = [Microsoft.Office.Interop.PowerPoint.PpFixedFormatType]::ppFixedFormatTypePDF
    $intent               = [Microsoft.Office.Interop.PowerPoint.PpFixedFormatIntent]::ppFixedFormatIntentScreen
    $handoutOrder         = [Microsoft.Office.Interop.PowerPoint.PpPrintHandoutOrder]::ppPrintHandoutVerticalFirst
    $outputType           = [Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputNotesPages
    $rangeType            = [Microsoft.Office.Interop.PowerPoint.PpPrintRangeType]::ppPrintAll
	  $frameSlides          = [Microsoft.Office.Core.MsoTriState]::msoFalse   
    $printHiddenSlides    = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $printRange           = $document.PrintOptions.Ranges.Add(1, $document.Slides.Count)
    $slideShowName        = "Slideshow Name"
    $includeDocProperties = $false
    $keepIRMSettings      = $true
    $docStructureTags     = $true
    $bitmapMissingFonts   = $true
    $useISO19005_1        = $false
    $includeMarkup        = $true
    $externalExporter     = $null

    $document.ExportAsFixedFormat2($exportPath, $fixedFormatType, $intent, $frameSlides, $handoutOrder, $outputType, $printHiddenSlides, $printRange, $rangeType, $slideShowName, $includeDocProperties, $keepIRMSettings, $docStructureTags, $bitmapMissingFonts, $useISO19005_1, $includeMarkup)
    $document.Close()
}

$ppt_app.Quit()

[System.GC]::Collect();
[System.GC]::WaitForPendingFinalizers();
[System.GC]::Collect();
[System.GC]::WaitForPendingFinalizers();
