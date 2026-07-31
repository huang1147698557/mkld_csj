[CmdletBinding()]
param(
    [ValidateSet("exe", "msi", "app-image")]
    [string]$Type = "exe",

    [string]$Version = "1.1.0",

    [string]$JavafxJmods = $env:JAVAFX_JMODS,

    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$AppName = "Procalc5"
$FatJarName = "procalc5-app.jar"
$MainClass = "com.sd.discovery.single.ProCalc5App"
$TargetJar = Join-Path $ProjectDir "target\\$FatJarName"
$InputDir = Join-Path $ProjectDir "target\\jpackage-input"
$OutputDir = Join-Path $ProjectDir "dist"

function Find-Jpackage {
    $candidates = @()
    if ($env:JAVA_HOME) {
        $candidates += Join-Path $env:JAVA_HOME "bin\\jpackage.exe"
    }

    $bundledJpackages = Get-ChildItem -Path (Join-Path $ProjectDir ".tools") -Recurse -Filter "jpackage.exe" -ErrorAction SilentlyContinue
    $candidates += $bundledJpackages.FullName

    $command = Get-Command jpackage.exe -ErrorAction SilentlyContinue
    if ($command) {
        $candidates += $command.Source
    }

    foreach ($candidate in $candidates | Select-Object -Unique) {
        if (Test-Path $candidate) {
            return (Resolve-Path $candidate).Path
        }
    }

    throw "jpackage.exe was not found. Install a JDK 17 or newer and set JAVA_HOME."
}

function Find-Maven {
    $candidates = @()
    $command = Get-Command mvn.cmd, mvn -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command) {
        $candidates += $command.Source
    }

    $bundledMaven = Get-ChildItem -Path (Join-Path $ProjectDir ".tools") -Recurse -Filter "mvn.cmd" -ErrorAction SilentlyContinue
    $candidates += $bundledMaven.FullName

    foreach ($candidate in $candidates | Where-Object { $_ } | Select-Object -Unique) {
        if (Test-Path $candidate) {
            return (Resolve-Path $candidate).Path
        }
    }

    throw "Maven was not found. Install Maven, or first build target\\$FatJarName and use -SkipBuild."
}

function Find-JavafxJmods([string]$configuredPath) {
    $candidates = @($configuredPath)
    $candidates += Join-Path $ProjectDir "javafx-sdk\\javafx-jmods-17.0.2"
    $candidates += Join-Path $ProjectDir "javafx-sdk\\javafx-jmods"

    foreach ($candidate in $candidates | Where-Object { $_ } | Select-Object -Unique) {
        if ((Test-Path $candidate) -and (Test-Path (Join-Path $candidate "javafx.controls.jmod"))) {
            return (Resolve-Path $candidate).Path
        }
    }

    throw "JavaFX jmods were not found. Download the Windows JavaFX jmods and pass -JavafxJmods <path>, or set JAVAFX_JMODS."
}

Set-Location $ProjectDir
$Jpackage = Find-Jpackage
$JavafxJmodsPath = Find-JavafxJmods $JavafxJmods
$JdkHome = Split-Path (Split-Path $Jpackage -Parent) -Parent
$env:JAVA_HOME = $JdkHome
$env:PATH = "$JdkHome\\bin;$env:PATH"
$bundledWix = Get-ChildItem -Path (Join-Path $ProjectDir ".tools") -Recurse -Filter "candle.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $bundledWix) {
    $bundledWix = Get-ChildItem -Path (Join-Path $ProjectDir ".tools") -Recurse -Filter "wix.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
}
if ($bundledWix) {
    $env:PATH = "$(Split-Path $bundledWix.FullName -Parent);$env:PATH"
}

if (-not $SkipBuild) {
    $maven = Find-Maven

    Write-Host "===== Step 1: Build JavaFX application JAR ====="
    & $maven clean package -Pjavafx-app -DskipTests
    if ($LASTEXITCODE -ne 0) {
        throw "Maven build failed with exit code $LASTEXITCODE."
    }
}

if (-not (Test-Path $TargetJar)) {
    throw "Application JAR was not found: $TargetJar"
}

Write-Host "===== Step 2: Prepare jpackage input ====="
if (Test-Path $InputDir) {
    Remove-Item -LiteralPath $InputDir -Recurse -Force
}
New-Item -ItemType Directory -Path $InputDir -Force | Out-Null
Copy-Item -LiteralPath $TargetJar -Destination (Join-Path $InputDir $FatJarName)
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

if ($Type -eq "exe") {
    $packagePath = Join-Path $OutputDir "$AppName-$Version.exe"
} elseif ($Type -eq "msi") {
    $packagePath = Join-Path $OutputDir "$AppName-$Version.msi"
} else {
    $packagePath = Join-Path $OutputDir $AppName
}

if (Test-Path $packagePath) {
    Remove-Item -LiteralPath $packagePath -Recurse -Force
}

Write-Host "===== Step 3: Create Windows $Type package ====="
$modules = "javafx.controls,javafx.fxml,java.base,java.logging,java.sql,java.xml,java.naming,java.management,java.desktop,jdk.charsets,jdk.crypto.ec,jdk.localedata,jdk.zipfs,java.instrument,jdk.unsupported,jdk.httpserver,jdk.crypto.cryptoki,jdk.net,jdk.sctp,java.net.http"
$arguments = @(
    "--name", $AppName,
    "--input", $InputDir,
    "--main-jar", $FatJarName,
    "--main-class", $MainClass,
    "--type", $Type,
    "--app-version", $Version,
    "--dest", $OutputDir,
    "--vendor", "Procalc5",
    "--description", "Procalc5 rotor calculation automation",
    "--module-path", $JavafxJmodsPath,
    "--add-modules", $modules,
    "--java-options", "-Dfile.encoding=UTF-8",
    "--java-options", "--add-opens=javafx.graphics/com.sun.javafx.tk=ALL-UNNAMED",
    "--java-options", "--add-opens=javafx.controls/javafx.scene.control=ALL-UNNAMED"
)

if ($Type -eq "exe") {
    # Create a regular user installer with the same desktop UI as the DMG app.
    $arguments += @(
        "--win-dir-chooser",
        "--win-menu",
        "--win-menu-group", $AppName,
        "--win-shortcut",
        "--win-shortcut-prompt",
        "--win-per-user-install"
    )
}

& $Jpackage @arguments
if ($LASTEXITCODE -ne 0) {
    throw "jpackage failed with exit code $LASTEXITCODE. Install WiX Toolset 3.11+ when building an EXE or MSI installer."
}

Write-Host ""
Write-Host "Package created: $packagePath"
Write-Host "Input JAR: $TargetJar"
