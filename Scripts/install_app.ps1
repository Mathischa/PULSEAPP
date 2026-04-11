param(
    [string]$AppName = "PulseApp",
    [string]$SourcePy = ".\0.000(test).py",
    [string]$InstallRoot = "C:\Tools\PulseApp",
    [string]$PythonVersionShort = "3.11",
    [string]$PythonVersionFull = "3.11.9"
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "=== Installation de $AppName ===" -ForegroundColor Cyan
Write-Host ""

# ---------- 1) Vérifie le fichier source ----------
if (!(Test-Path $SourcePy)) {
    throw "Le fichier source '$SourcePy' est introuvable. Mets 0.000(test).py à côté de install_app.ps1."
}

# ---------- 2) Prépare les chemins ----------
$InstallRoot   = $InstallRoot.TrimEnd('\')
$VenvDir       = Join-Path $InstallRoot "venv"
$MainPy        = Join-Path $InstallRoot "app.py"
$Requirements  = Join-Path $InstallRoot "requirements.txt"
$LogsDir       = Join-Path $InstallRoot "logs"
$LauncherBat   = Join-Path $InstallRoot "Lancer_$AppName.bat"
$DesktopPath   = [Environment]::GetFolderPath("Desktop")
$ShortcutPath  = Join-Path $DesktopPath "$AppName.lnk"
$TempDir       = Join-Path $env:TEMP "$AppName-setup"
$PyInstaller   = Join-Path $TempDir "python-$PythonVersionFull-amd64.exe"
$PyDownloadUrl = "https://www.python.org/ftp/python/$PythonVersionFull/python-$PythonVersionFull-amd64.exe"

if (!(Test-Path $InstallRoot)) { New-Item -ItemType Directory -Path $InstallRoot -Force | Out-Null }
if (!(Test-Path $LogsDir)) { New-Item -ItemType Directory -Path $LogsDir -Force | Out-Null }
if (!(Test-Path $TempDir)) { New-Item -ItemType Directory -Path $TempDir -Force | Out-Null }

# ---------- 3) Détecte Python 3.11 ----------
$PythonCmd = $null

try {
    & py -$PythonVersionShort --version *> $null
    $PythonCmd = "py -$PythonVersionShort"
    Write-Host "Python $PythonVersionShort déjà détecté via py launcher." -ForegroundColor Green
} catch {
    Write-Host "Python $PythonVersionShort non détecté. Téléchargement de Python $PythonVersionFull..." -ForegroundColor Yellow

    Invoke-WebRequest -Uri $PyDownloadUrl -OutFile $PyInstaller

    if (!(Test-Path $PyInstaller)) {
        throw "Téléchargement de Python impossible."
    }

    Write-Host "Installation silencieuse de Python..." -ForegroundColor Yellow

    $installArgs = @(
        "/quiet"
        "InstallAllUsers=0"
        "PrependPath=1"
        "Include_launcher=1"
        "Include_pip=1"
        "Include_tcltk=1"
        "Include_test=0"
        "Shortcuts=0"
    )

    $proc = Start-Process -FilePath $PyInstaller -ArgumentList $installArgs -Wait -PassThru
    if ($proc.ExitCode -ne 0) {
        throw "L'installation silencieuse de Python a échoué. Code retour : $($proc.ExitCode)"
    }

    Start-Sleep -Seconds 3

    try {
        & py -$PythonVersionShort --version *> $null
        $PythonCmd = "py -$PythonVersionShort"
        Write-Host "Python $PythonVersionShort installé avec succès." -ForegroundColor Green
    } catch {
        # secours : chemin utilisateur classique
        $CandidatePython = Join-Path $env:LocalAppData "Programs\Python\Python311\python.exe"
        if (Test-Path $CandidatePython) {
            $PythonCmd = "`"$CandidatePython`""
            Write-Host "Python trouvé via chemin local utilisateur." -ForegroundColor Green
        } else {
            throw "Python semble installé, mais introuvable ensuite."
        }
    }
}

# ---------- 4) Copie de l'application ----------
Copy-Item $SourcePy $MainPy -Force
Write-Host "Application copiée dans $InstallRoot" -ForegroundColor Green

# ---------- 5) requirements.txt ----------
@"
pandas
numpy
matplotlib
seaborn
openpyxl
pillow
customtkinter
scikit-learn
catboost
xgboost
lightgbm
"@ | Set-Content -Path $Requirements -Encoding UTF8

Write-Host "requirements.txt créé." -ForegroundColor Green

# ---------- 6) Crée le venv ----------
Write-Host "Création / mise à jour du venv..." -ForegroundColor Yellow

if (Test-Path $VenvDir) {
    Remove-Item $VenvDir -Recurse -Force
}

if ($PythonCmd -like "py*") {
    & py -$PythonVersionShort -m venv $VenvDir
} else {
    Invoke-Expression "$PythonCmd -m venv `"$VenvDir`""
}

$VenvPython  = Join-Path $VenvDir "Scripts\python.exe"
$VenvPip     = Join-Path $VenvDir "Scripts\pip.exe"

if (!(Test-Path $VenvPython)) {
    throw "Le venv n'a pas pu être créé."
}

# ---------- 7) Upgrade pip ----------
Write-Host "Mise à jour de pip/setuptools/wheel..." -ForegroundColor Yellow
& $VenvPython -m pip install --upgrade pip setuptools wheel

# ---------- 8) Installe les dépendances ----------
Write-Host "Installation des dépendances..." -ForegroundColor Yellow
& $VenvPip install -r $Requirements

# ---------- 9) Lanceur BAT ----------
$BatContent = @"
@echo off
title $AppName
cd /d "$InstallRoot"
echo Lancement de $AppName...
echo.
"$VenvPython" "$MainPy" 1>> "$LogsDir\stdout.log" 2>> "$LogsDir\stderr.log"
echo.
echo Code retour : %errorlevel%
echo.
pause
"@
Set-Content -Path $LauncherBat -Value $BatContent -Encoding ASCII
Write-Host "Lanceur créé : $LauncherBat" -ForegroundColor Green

# ---------- 10) Raccourci Bureau ----------
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $LauncherBat
$Shortcut.WorkingDirectory = $InstallRoot
$Shortcut.Description = "Lancer $AppName"
$Shortcut.IconLocation = "$env:SystemRoot\System32\shell32.dll,220"
$Shortcut.WindowStyle = 1
$Shortcut.Save()

Write-Host "Raccourci Bureau créé : $ShortcutPath" -ForegroundColor Green

# ---------- 11) Nettoyage léger ----------
if (Test-Path $PyInstaller) {
    Remove-Item $PyInstaller -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "=== Installation terminée ===" -ForegroundColor Cyan
Write-Host "Application : $MainPy"
Write-Host "Venv        : $VenvDir"
Write-Host "Logs        : $LogsDir"
Write-Host "Raccourci   : $ShortcutPath"
Write-Host ""
Write-Host "Le lancement se fera maintenant AVEC console." -ForegroundColor Cyan