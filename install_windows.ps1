param(
    [switch]$ForcePythonInstall,
    [switch]$SkipPipUpgrade
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Get-PythonCommand {
    if (-not $ForcePythonInstall) {
        $python = Get-Command python -ErrorAction SilentlyContinue
        if ($python -and (Test-PythonCommand -PythonCommand @($python.Source))) {
            return @($python.Source)
        }

        $py = Get-Command py -ErrorAction SilentlyContinue
        if ($py -and (Test-PythonCommand -PythonCommand @($py.Source, "-3"))) {
            return @($py.Source, "-3")
        }
    }

    return $null
}

function Test-PythonCommand {
    param([string[]]$PythonCommand)

    try {
        Invoke-Python -PythonCommand $PythonCommand -Arguments @("-c", "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)")
        return $LASTEXITCODE -eq 0
    }
    catch {
        return $false
    }
}

function Update-ProcessPath {
    $machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
    [Environment]::SetEnvironmentVariable("Path", "$machinePath;$userPath", "Process")
}

function Install-Python {
    Write-Step "Python was not found. Installing Python with winget"

    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        throw "Python is not installed and winget is not available. Install Python from https://www.python.org/downloads/windows/ and rerun this script."
    }

    & $winget.Source install `
        --exact `
        --id Python.Python.3.12 `
        --scope user `
        --accept-package-agreements `
        --accept-source-agreements

    Update-ProcessPath

    $pythonCommand = Get-PythonCommand
    if (-not $pythonCommand) {
        throw "Python installation completed, but Python is not available in this PowerShell session. Open a new PowerShell window and rerun this script."
    }

    return $pythonCommand
}

function Invoke-Python {
    param(
        [string[]]$PythonCommand,
        [string[]]$Arguments
    )

    $executable = $PythonCommand[0]
    $baseArgs = @()
    if ($PythonCommand.Count -gt 1) {
        $baseArgs = $PythonCommand[1..($PythonCommand.Count - 1)]
    }

    & $executable @baseArgs @Arguments
}

function Get-IsVirtualEnvironment {
    param([string[]]$PythonCommand)

    $output = Invoke-Python -PythonCommand $PythonCommand -Arguments @("-c", "import sys; print(int(sys.prefix != sys.base_prefix))")
    $value = $output | Select-Object -First 1
    return ($LASTEXITCODE -eq 0 -and $value -and $value.Trim() -eq "1")
}

function Add-UserScriptsToPath {
    param([string[]]$PythonCommand)

    $userBase = Invoke-Python -PythonCommand $PythonCommand -Arguments @("-m", "site", "--user-base")
    if ($LASTEXITCODE -eq 0 -and $userBase) {
        $scriptsPath = Join-Path ($userBase | Select-Object -First 1) "Scripts"
        if (Test-Path $scriptsPath) {
            $currentPath = [Environment]::GetEnvironmentVariable("Path", "Process")
            if ($currentPath -notlike "*$scriptsPath*") {
                [Environment]::SetEnvironmentVariable("Path", "$currentPath;$scriptsPath", "Process")
            }
        }
    }
}

function Get-UserScriptsPath {
    param([string[]]$PythonCommand)

    $userBase = Invoke-Python -PythonCommand $PythonCommand -Arguments @("-m", "site", "--user-base")
    if ($LASTEXITCODE -eq 0 -and $userBase) {
        return Join-Path ($userBase | Select-Object -First 1) "Scripts"
    }

    return $null
}

function Get-InstalledAppCommandPath {
    param([string[]]$PythonCommand)

    $command = Get-Command gdck-admission -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $scriptsPath = Get-UserScriptsPath -PythonCommand $PythonCommand
    if ($scriptsPath) {
        $candidate = Join-Path $scriptsPath "gdck-admission.exe"
        if (Test-Path $candidate) {
            return $candidate
        }
        $candidate = Join-Path $scriptsPath "gdck-admission-script.py"
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    return $null
}

function Get-PackageIconPath {
    param([string[]]$PythonCommand)

    $output = Invoke-Python -PythonCommand $PythonCommand -Arguments @(
        "-c",
        "from gdck_admission.app import APP_ICON_FILE; print(APP_ICON_FILE)"
    )
    if ($LASTEXITCODE -eq 0 -and $output) {
        $iconPath = ($output | Select-Object -First 1).Trim()
        if (Test-Path $iconPath) {
            return $iconPath
        }
    }

    return Join-Path $projectRoot "gdck_admission\assets\app.ico"
}

function New-DesktopShortcut {
    param(
        [string]$TargetPath,
        [string]$IconPath
    )

    $desktop = [Environment]::GetFolderPath("Desktop")
    $shortcutPath = Join-Path $desktop "GDCK Admission.lnk"
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.WorkingDirectory = [Environment]::GetFolderPath("UserProfile")
    if ($IconPath -and (Test-Path $IconPath)) {
        $shortcut.IconLocation = $IconPath
    }
    $shortcut.Description = "Open GDCK Admission"
    $shortcut.Save()
    return $shortcutPath
}

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

if (-not (Test-Path (Join-Path $projectRoot "pyproject.toml"))) {
    throw "pyproject.toml was not found. Run this script from the GDCK_ADMISSION project folder."
}

$pythonCommand = Get-PythonCommand
if (-not $pythonCommand) {
    $pythonCommand = Install-Python
}

Write-Step "Using Python"
Invoke-Python -PythonCommand $pythonCommand -Arguments @("--version")

Write-Step "Ensuring pip is available"
Invoke-Python -PythonCommand $pythonCommand -Arguments @("-m", "ensurepip", "--upgrade")

if (-not $SkipPipUpgrade) {
    Write-Step "Updating packaging tools"
    Invoke-Python -PythonCommand $pythonCommand -Arguments @("-m", "pip", "install", "--upgrade", "pip", "setuptools", "wheel")
}

Write-Step "Installing GDCK Admission and dependencies"
$installArgs = @("-m", "pip", "install", ".")
if (-not (Get-IsVirtualEnvironment -PythonCommand $pythonCommand)) {
    $installArgs = @("-m", "pip", "install", "--user", ".")
}
Invoke-Python -PythonCommand $pythonCommand -Arguments $installArgs
Add-UserScriptsToPath -PythonCommand $pythonCommand

Write-Step "Checking installed command"
$commandPath = Get-InstalledAppCommandPath -PythonCommand $pythonCommand
if ($commandPath) {
    Write-Host "Installed command: $commandPath" -ForegroundColor Green

    Write-Step "Creating desktop shortcut"
    $iconPath = Get-PackageIconPath -PythonCommand $pythonCommand
    $shortcutPath = New-DesktopShortcut -TargetPath $commandPath -IconPath $iconPath
    Write-Host "Desktop shortcut: $shortcutPath" -ForegroundColor Green
}
else {
    Write-Host "The app was installed, but gdck-admission is not on this PowerShell PATH yet." -ForegroundColor Yellow
    Write-Host "Open a new PowerShell window, then run: gdck-admission"
}

Write-Host ""
Write-Host "Installation complete." -ForegroundColor Green
Write-Host "Run the app with: gdck-admission"
