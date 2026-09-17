$ErrorActionPreference = "Stop"

function Find-Python {
    $Py = Get-Command py -ErrorAction SilentlyContinue
    if ($Py) {
        $Resolved = & $Py.Source -3 -c "import sys; print(sys.executable)"
        if ($LASTEXITCODE -eq 0 -and $Resolved) { return $Resolved.Trim() }
    }
    $Python = Get-Command python -ErrorAction SilentlyContinue
    if ($Python) { return $Python.Source }
    throw "Python 3.11 or newer was not found. Install Python for Windows, then run install_xExcel.bat again."
}

$Python = Find-Python
& $Python -c "import sys; raise SystemExit(0 if sys.version_info >= (3,11) else 1)"
if ($LASTEXITCODE -ne 0) { throw "xExcel Viewer requires Python 3.11 or newer." }

$InstallRoot = Join-Path $env:LOCALAPPDATA "xExcel Viewer"
$VenvRoot = Join-Path $InstallRoot "venv"
$VenvPython = Join-Path $VenvRoot "Scripts\python.exe"
$VenvPythonw = Join-Path $VenvRoot "Scripts\pythonw.exe"
$IconSource = Join-Path $PSScriptRoot "mdir\assets\xexcel.ico"
$InstalledIcon = Join-Path $InstallRoot "xexcel.ico"
$BinRoot = Join-Path $InstallRoot "bin"

New-Item -ItemType Directory -Path $InstallRoot -Force | Out-Null
New-Item -ItemType Directory -Path $BinRoot -Force | Out-Null

$NeedVenv = $true
if (Test-Path -LiteralPath $VenvPython) {
    try {
        & $VenvPython -c "import sys; raise SystemExit(0 if sys.version_info >= (3,11) else 1)"
        if ($LASTEXITCODE -eq 0) { $NeedVenv = $false }
    } catch {
        $NeedVenv = $true
    }
}

if ($NeedVenv) {
    if (Test-Path -LiteralPath $VenvRoot) {
        Write-Host "[1/4] Recreating incompatible private Python environment..." -ForegroundColor Yellow
        Remove-Item -LiteralPath $VenvRoot -Recurse -Force
    } else {
        Write-Host "[1/4] Creating private Python environment..." -ForegroundColor Green
    }
    & $Python -m venv $VenvRoot
    if ($LASTEXITCODE -ne 0) { throw "Could not create the private Python environment." }
} else {
    Write-Host "[1/4] Existing private Python environment is compatible." -ForegroundColor Green
}

Write-Host "[2/4] Installing xExcel Viewer..." -ForegroundColor Green
& $VenvPython -m pip install --disable-pip-version-check --prefer-binary --upgrade --force-reinstall $PSScriptRoot
if ($LASTEXITCODE -ne 0) { throw "Could not install xExcel Viewer." }

if (Test-Path -LiteralPath $IconSource) { Copy-Item $IconSource $InstalledIcon -Force }

$Launcher = '@echo off' + "`r`n" + '"' + $VenvPythonw + '" -P -m mdir %*' + "`r`n"
Set-Content -LiteralPath (Join-Path $BinRoot "xexcel.cmd") -Value $Launcher -Encoding Ascii

Write-Host "[3/4] Creating Desktop and Start Menu shortcuts..." -ForegroundColor Green
$Shell = New-Object -ComObject WScript.Shell
$Desktop = [Environment]::GetFolderPath("Desktop")
$StartMenu = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
foreach ($ShortcutPath in @((Join-Path $Desktop "xExcel Viewer.lnk"),(Join-Path $StartMenu "xExcel Viewer.lnk"))) {
    $Shortcut = $Shell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $VenvPythonw
    $Shortcut.Arguments = "-P -m mdir"
    $Shortcut.WorkingDirectory = [Environment]::GetFolderPath("UserProfile")
    if (Test-Path $InstalledIcon) { $Shortcut.IconLocation = "$InstalledIcon,0" }
    $Shortcut.Description = "xExcel Viewer - Excel files on the left, full editable workbook on the right"
    $Shortcut.Save()
}

Write-Host "[4/4] Registering Open with xExcel Viewer..." -ForegroundColor Green
$ProgId = "xExcelViewer.Workbook"
$ProgRoot = "HKCU:\Software\Classes\$ProgId"
New-Item -Path $ProgRoot -Force | Out-Null
Set-Item -Path $ProgRoot -Value "Excel Workbook - xExcel Viewer"
New-Item -Path "$ProgRoot\DefaultIcon" -Force | Out-Null
Set-Item -Path "$ProgRoot\DefaultIcon" -Value "`"$InstalledIcon`",0"
New-Item -Path "$ProgRoot\shell\open\command" -Force | Out-Null
Set-Item -Path "$ProgRoot\shell\open\command" -Value "`"$VenvPythonw`" -P -m mdir `"%1`""
foreach ($Ext in @(".xlsx", ".xlsm", ".xltx", ".xltm", ".xls")) {
    $OpenWith = "HKCU:\Software\Classes\$Ext\OpenWithProgids"
    New-Item -Path $OpenWith -Force | Out-Null
    New-ItemProperty -Path $OpenWith -Name $ProgId -Value "" -PropertyType String -Force | Out-Null
}

$Check = & $VenvPython -P -m mdir --check
if ($LASTEXITCODE -ne 0) { throw "Installation validation failed." }
Write-Host ""
Write-Host $Check -ForegroundColor Green
Write-Host "Installed successfully: $InstallRoot" -ForegroundColor Green
