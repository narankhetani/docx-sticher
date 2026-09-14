# Installs (or updates) DOCX Stitcher on Windows.
#
#   powershell -ExecutionPolicy ByPass -c "irm https://raw.githubusercontent.com/narankhetani/docx-sticher/main/install.ps1 | iex"
#
# What it does: installs uv if needed, installs the app, and adds Start Menu, Desktop
# and "Send to" shortcuts. Run it again at any time to update.

$ErrorActionPreference = "Stop"
$Source = "docx-stitcher @ https://github.com/narankhetani/docx-sticher/archive/refs/heads/main.zip"
$AppName = "DOCX Stitcher"

function Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }

# 1. uv
if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Step "Installing uv (Python package manager)"
    powershell -NoProfile -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
    $env:Path = "$env:USERPROFILE\.local\bin;$env:Path"
    if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
        throw "uv was installed but can't be found. Close this window, open a new PowerShell and run the installer again."
    }
} else {
    uv self update *> $null  # older uv versions ship a Python whose Tk is broken; ignore if uv came from winget etc.
}

# 2. The app (downloads a private copy of Python if needed; no git required)
Step "Installing $AppName"
uv python install 3.13
if ($LASTEXITCODE -ne 0) { throw "Couldn't install Python." }
uv python upgrade 3.13 *> $null
uv tool install --python 3.13 --force --reinstall $Source
if ($LASTEXITCODE -ne 0) { throw "Couldn't install $AppName." }
uv tool update-shell *> $null  # puts docx-stitcher on PATH for new terminals

$BinDir = (uv tool dir --bin).Trim()
$ToolDir = (uv tool dir).Trim()
$Exe = Join-Path $BinDir "docx-stitcher-app.exe"
$Icon = Get-ChildItem -Path (Join-Path $ToolDir "docx-stitcher") -Recurse -Filter "icon.ico" -ErrorAction SilentlyContinue |
    Select-Object -First 1 -ExpandProperty FullName
if (-not (Test-Path $Exe)) { throw "Installed, but $Exe is missing." }

# 3. Shortcuts
Step "Adding shortcuts"
$Shell = New-Object -ComObject WScript.Shell
$Places = @(
    [Environment]::GetFolderPath("Programs"),  # Start Menu
    [Environment]::GetFolderPath("Desktop"),
    [Environment]::GetFolderPath("SendTo")     # right-click a folder > Send to
)
foreach ($Place in $Places) {
    $Link = $Shell.CreateShortcut((Join-Path $Place "$AppName.lnk"))
    $Link.TargetPath = $Exe
    $Link.WorkingDirectory = [Environment]::GetFolderPath("MyDocuments")
    $Link.Description = "Merge Word documents into one"
    if ($Icon) { $Link.IconLocation = $Icon }
    $Link.Save()
}

Write-Host ""
Write-Host "$AppName is installed." -ForegroundColor Green
Write-Host "  - Open it from the Start Menu or the Desktop shortcut."
Write-Host "  - Or right-click a folder of .docx files > Send to > $AppName."
Write-Host "  - To update later, run this installer again."
Start-Process $Exe
