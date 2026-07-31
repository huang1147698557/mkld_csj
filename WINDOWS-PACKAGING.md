# Windows installer build

The Windows installer uses the same JavaFX entry point as the macOS DMG:
`com.sd.discovery.single.ProCalc5App`. It installs the visual desktop application rather than the legacy command-line entry point.

## Prerequisites

- Windows 10 or Windows 11
- JDK 17 or newer, with `JAVA_HOME` pointing to that JDK
- Maven 3.8 or newer
- JavaFX 17.0.2 Windows jmods
- WiX Toolset 3.11 or newer in `PATH` when building an `exe` or `msi`

Download the JavaFX **Windows jmods** archive, extract it, and either set `JAVAFX_JMODS` to the directory that contains `javafx.controls.jmod` or place that directory at `javafx-sdk/javafx-jmods-17.0.2` in the project root.

## Build

Run from PowerShell:

```powershell
.\build-windows.ps1
```

This creates `dist\Procalc5-1.1.0.exe`. The installer provides an installation-directory selector, Start Menu entry, and desktop shortcut. It is a per-user install and does not require administrator privileges.

Other output types:

```powershell
.\build-windows.ps1 -Type msi
.\build-windows.ps1 -Type app-image
```

`app-image` produces a portable application directory and does not require WiX. To package an existing `target\procalc5-app.jar` without running Maven again, pass `-SkipBuild`.

The `build-windows.bat` file is a double-clickable wrapper for the same command.
