# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

AESCConstruct2026 is a SpaceClaim AddIn for ANSYS Discovery (v251). It provides structural modeling tools: frame/profile generation, fasteners, plates, connectors, rib cutouts, engraving, and custom component properties. The UI is WPF (XAML + code-behind) hosted inside SpaceClaim's dockable/floating panels.

## Build

- **Solution:** `AESCConstruct2026.sln` (Visual Studio 2017+)
- **Framework:** .NET Framework 4.8, C# 12.0
- **Output:** Class Library DLL
- **Debug output:** `C:\Program Files\ANSYS Inc\v251\scdm\Addins\AESCConstruct2026\`
- **Release output:** `C:\Program Files\ANSYS Inc\v251\scdm\Addins\AESCConstruct2026\`
- Build via Visual Studio or `msbuild AESCConstruct2026.sln`
- No test framework is configured; testing is done manually in SpaceClaim

**AddIn registration:** `AESCConstruct2026.xml` must exist in the output directory alongside the DLL. If missing after build, copy it manually (see README.md).

## Architecture

### Entry Point
`Construct2026.cs` implements `IExtensibility` and `IRibbonExtensibility`. The `Connect()` method initializes the SpaceClaim API, validates the license, and registers all ribbon commands with `Command.Create()`. Commands use `KeepAlive(true)` to prevent GC collection.

### SpaceClaim API
All SpaceClaim types come from `SpaceClaim.Api.V242`. Key aliases used throughout:
- `Application` = `SpaceClaim.Api.V242.Application`
- `Window` = `SpaceClaim.Api.V242.Window`

For detailed API reference (types, methods, properties), see `SpaceClaim_API_Reference.md`.

### Module Structure
Each feature module follows the pattern: `Module/` (business logic), `UI/` (XAML controls), optionally `Commands/` and `Utilities/`.

| Module | Purpose |
|--------|---------|
| **FrameGenerator** | Profile extrusion along curves, joint creation. Contains `ProfileBase`/`JointBase` hierarchies with specialized subclasses (Rectangular, H, L, T, U, Circular, DXF, CSV profiles; Miter, Straight, T, Trim joints) |
| **Fastener** | Bolt/Nut/Washer placement from CSV data files |
| **Plates** | Standard plate geometry with holes/slots/fillets |
| **Connector** | TubeLocker-style connectors with cylindrical geometry patterns |
| **RibCutout** | Rib cutout operations on structural members |
| **Engraving** | Text/pattern engraving on faces |
| **CustomComponent** | Custom property assignment to components |

### UI System
- **UIManager** (`UIMain/UIManager.cs`): Registers docked and floating panel modes for each module's control. Uses `ElementHost` (WinForms-WPF bridge) for SpaceClaim panel integration.
- **Ribbon** (`UIMain/Ribbon.xml`): Defines the "AESC Construct" tab with button groups. Labels resolved via `GetRibbonLabel` callback using the localization system.
- **Localization** (`UIMain/Language.cs`): CSV-based translation loaded from `%ProgramData%\AESCConstruct\Language\languageConstruct.csv`. Access via `Language.Translate("key")`.

### Licensing
`Licensing/LicenseSpot.cs` wraps LicenseSpot.Framework. Supports node-locked and network licenses. Commands are enabled/disabled based on `ConstructLicenseSpot.IsValid`. License file at `%ProgramData%\IPManager\AESC_License_Construct2026.lic`.

### Data Files
All runtime data (profile CSVs, fastener CSVs, language files, connector properties) stored under `%ProgramData%\AESCConstruct\`. Paths configured in `Properties/Settings.settings` and `app.config`.

### Logging
`Utilities/Logger.cs` (namespace `AESCConstruct2026.FrameGenerator.Utilities`) writes to `%ProgramData%\AESCConstruct\AESCConstruct2026_Log.txt`. Status messages shown in SpaceClaim via `Application.ReportStatus()`.

## Command Naming Convention

SpaceClaim commands use dotted names: `AESCConstruct2026.ExportExcel`, `AESCConstruct2026.Plate`, `AESC.Construct.SetMode3D`, etc. These are registered in `Construct2026.cs` and referenced in `Ribbon.xml`.

## Installer

Inno Setup script at `Installer/Construct2026Installer.iss`. Outputs to `Installer/Output/`.

## Key Namespaces

- `AESCConstruct2026` — root, entry point, resources
- `AESCConstruct2026.FrameGenerator` — profiles, joints, commands, utilities
- `AESCConstruct2026.FrameGenerator.Utilities` — Logger, DXFImportHelper, CompNameHelper (files in `Utilities/`)
- `AESCConstruct2026.Fastener` — fastener module and UI
- `AESCConstruct2026.Connector` — TubeLocker connector module
- `AESCConstruct2026.UIMain` — ribbon, panel management, settings. Note: `Engraving/EngravingService.cs` also uses this namespace despite residing in the Engraving folder.
- `AESCConstruct2026.Localization` — language/translation support (`UIMain/Language.cs`)
- `AESCConstruct2026.Licensing` — license validation
