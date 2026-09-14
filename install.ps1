# Installs (or updates) DOCX Stitcher on Windows.
#
#   powershell -ExecutionPolicy ByPass -c "irm https://raw.githubusercontent.com/narankhetani/docx-sticher/main/install.ps1 | iex"
#
# What it does: installs uv if needed, installs the app, and adds Start Menu, Desktop
# and "Send to" shortcuts. Run it again at any time to update.

# Not "Stop": Windows PowerShell 5.1 turns anything a program writes to stderr (uv's normal
# progress messages) into a fatal error. Failures are detected with $LASTEXITCODE instead.
$ErrorActionPreference = "Continue"
$ProgressPreference = "SilentlyContinue"
$Source = "docx-stitcher @ https://github.com/narankhetani/docx-sticher/archive/refs/heads/main.zip"
$AppName = "DOCX Stitcher"

function Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }

function Fail($text) { throw "Installation failed: $text" }  # throw, not exit, so the window stays open

# Runs uv, hiding its output unless it fails.
function Invoke-Uv {
    $output = & uv @args 2>&1 | ForEach-Object { "$_" }
    if ($LASTEXITCODE -ne 0) {
        $output | Write-Host
        return $false
    }
    return $true
}

# 1. uv
if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Step "Installing uv (Python package manager)"
    powershell -NoProfile -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
    $env:Path = "$env:USERPROFILE\.local\bin;$env:Path"
    if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
        Fail "uv was installed but can't be found. Close this window, open a new PowerShell and run the installer again."
    }
} else {
    # Older uv versions ship a Python whose Tk is broken. This fails harmlessly if uv came from winget/pip.
    Step "Updating uv"
    Invoke-Uv self update | Out-Null
}

# 2. Python and the app (no git or admin rights needed)
Step "Installing Python"
if (-not (Invoke-Uv python install 3.13)) { Fail "couldn't install Python." }
Invoke-Uv python upgrade 3.13 | Out-Null

Step "Installing $AppName"
if (-not (Invoke-Uv tool install --python 3.13 --force --reinstall $Source)) { Fail "couldn't install $AppName." }
Invoke-Uv tool update-shell | Out-Null  # puts docx-stitcher on PATH for new terminals

$BinDir = (& uv tool dir --bin 2>$null | Out-String).Trim()
$ToolDir = (& uv tool dir 2>$null | Out-String).Trim()
$Exe = Join-Path $BinDir "docx-stitcher-app.exe"
if (-not (Test-Path $Exe)) { Fail "the app was installed, but $Exe is missing." }
$Icon = Get-ChildItem -Path (Join-Path $ToolDir "docx-stitcher") -Recurse -Filter "icon.ico" -ErrorAction SilentlyContinue |
    Select-Object -First 1 -ExpandProperty FullName

# 3. Shortcuts
Step "Adding shortcuts"
$Shell = New-Object -ComObject WScript.Shell
foreach ($Folder in "Programs", "Desktop", "SendTo") {  # Start Menu, Desktop, right-click > Send to
    try {
        $Link = $Shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath($Folder)) "$AppName.lnk"))
        $Link.TargetPath = $Exe
        $Link.WorkingDirectory = [Environment]::GetFolderPath("MyDocuments")
        $Link.Description = "Merge Word documents into one"
        if ($Icon) { $Link.IconLocation = $Icon }
        $Link.Save()
    } catch {
        Write-Host "  Couldn't add the $Folder shortcut: $_" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "$AppName is installed." -ForegroundColor Green
Write-Host "  - Open it from the Start Menu or the Desktop shortcut."
Write-Host "  - Or right-click a folder of .docx files > Send to > $AppName."
Write-Host "  - To update later, run this installer again."
Start-Process $Exe
