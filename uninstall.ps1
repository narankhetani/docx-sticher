# Removes DOCX Stitcher and its shortcuts from Windows (uv itself is left installed).
#
#   powershell -ExecutionPolicy ByPass -c "irm https://raw.githubusercontent.com/narankhetani/docx-sticher/main/uninstall.ps1 | iex"

$AppName = "DOCX Stitcher"
foreach ($Folder in "Programs", "Desktop", "SendTo") {
    Remove-Item (Join-Path ([Environment]::GetFolderPath($Folder)) "$AppName.lnk") -ErrorAction SilentlyContinue
}
if (Get-Command uv -ErrorAction SilentlyContinue) {
    uv tool uninstall docx-stitcher
}
Write-Host "$AppName has been removed." -ForegroundColor Green
