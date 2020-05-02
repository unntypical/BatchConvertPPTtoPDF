# If PowerShell exits with an error, check if unsigned scripts are allowed in your system.
# You can allow them by calling PowerShell as an Administrator and typing
# Set-ExecutionPolicy Unrestricted

$curr_path = Split-Path -parent $MyInvocation.MyCommand.Path
$ppt_app   = New-Object -ComObject PowerPoint.Application
Get-ChildItem -Path $curr_path -Recurse -Filter *.ppt? | ForEach-Object {
    Write-Host "Proscessing" $_.FullName "..."
    $document     = $ppt_app.Presentations.Open($_.FullName)
    $pdf_filename = "$($curr_path)\$($_.BaseName)_cvt.pdf"
    $opt          = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
    $document.SaveAs($pdf_filename, $opt)
    $document.Close()
}
$ppt_app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt_app)
