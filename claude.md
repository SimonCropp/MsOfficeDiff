# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

MsWordDiff is a .NET global tool that compares two Word documents using Microsoft Word's COM automation. It opens Word with a comparison view showing tracked changes between documents.

## Solution

`src/MsWordDiff.slnx`

## Build Commands

```bash
# Build the solution
dotnet build src/MsWordDiff.slnx

# Build release
dotnet build src --configuration Release

# Install as global tool locally
dotnet pack src/MsWordDiff/MsWordDiff.csproj
dotnet tool install -g MsWordDiff --add-source ./nugets
```

## Architecture

The codebase is minimal with the following components:

- **Word.cs** - Core functionality using COM interop to automate Microsoft Word:
  - Opens documents read-only via `Word.Application` COM object
  - Uses `CompareDocuments` API to generate comparison view
  - Creates Windows Job Object to terminate Word when parent process exits
  - Configures Word window (minimizes ribbon, optionally shows source documents and revision pane)
  - Accepts `quiet` parameter to control source documents visibility

- **CompareCommand.cs** - Main CLI command using CliFx:
  - Validates document paths
  - Supports `--quiet` option to hide source documents
  - Reads settings asynchronously and merges with command-line options (CLI overrides settings)
  - Invokes `Word.Launch()` with merged configuration

- **SettingsCommand.cs** - View settings command:
  - `settings` - Display settings file path and current settings

- **SetQuietCommand.cs** - Configure Quiet mode:
  - `set-quiet <true|false>` - Set default Quiet mode value

- **Settings.cs** - Settings model with `Quiet` property

- **SettingsManager.cs** - Async settings persistence:
  - Reads/writes JSON settings from `%USERPROFILE%\.config\MsWordDiff\settings.json`
  - `SetQuiet()` - Update Quiet setting
  - Handles missing/corrupted settings gracefully
  - Configurable settings path for testing

- **Program.cs** - CLI entry point using CliFx

- **Logging.cs** - Serilog configuration writing to console and rolling file logs

## MsExcelDiff Project

MsExcelDiff is a separate .NET global tool (`diffexcel`) that compares two Excel workbooks using Microsoft's Spreadsheet Compare (`SPREADSHEETCOMPARE.EXE`), bundled with Office Professional Plus / Microsoft 365 Apps for Enterprise.

### Architecture

- **SpreadsheetCompare.cs** - Core functionality:
  - Auto-detects `SPREADSHEETCOMPARE.EXE` in common Office install paths (Office16/Office15, x86/x64)
  - Auto-detects `AppVLP.exe` (Office Click-to-Run virtualization launcher) and uses it when present
  - Writes both file paths to a temp file (one per line) and passes it as argument
  - `AppVLP.exe` is a launcher that exits immediately; the code finds the real `SPREADSHEETCOMPARE` process by name
  - Creates Windows Job Object to terminate Spreadsheet Compare when parent exits
  - Temp file is deleted by the exe on success; cleaned up on failure

- **CompareCommand.cs** - Default CliFx command:
  - Two positional `FileInfo` parameters for workbook paths
  - Reads settings for optional custom exe path
  - Invokes `SpreadsheetCompare.Launch()`

- **Settings.cs** - Settings model with `SpreadsheetComparePath` property

- **SettingsManager.cs** - Async settings persistence at `%USERPROFILE%\.config\MsExcelDiff\settings.json`

- **SetSpreadsheetComparePathCommand.cs** - `set-path` command to configure custom exe location

- **SettingsCommand.cs** - `settings` command to display current settings

### Key Technical Detail: AppVLP.exe

Click-to-Run Office installs require launching `SPREADSHEETCOMPARE.EXE` via `AppVLP.exe` (the App-V virtualization layer). The exe crashes (`0xC0000409`) if launched directly because it depends on the virtual filesystem/registry that `AppVLP.exe` provides. The start menu shortcut uses the same pattern:

```
AppVLP.exe "...\SPREADSHEETCOMPARE.EXE"
```

### Build & Install

```bash
dotnet pack src/MsExcelDiff/MsExcelDiff.csproj
dotnet tool install -g MsExcelDiff --add-source ./nugets
```

## Testing

Uses TUnit framework.

```bash
# Run MsWordDiff tests
dotnet run --project src/Tests

# Run MsExcelDiff tests
dotnet run --project src/ExcelTests
```

MsWordDiff `[Explicit]` tests require Word to be installed. MsExcelDiff `[Explicit]` tests require Spreadsheet Compare (Office Professional Plus / Microsoft 365 Apps for Enterprise).

## Key Technical Details

- Uses dynamic COM interop (no Word interop assemblies required)
- Windows-only due to COM automation and Windows Forms dependency
- Targets .NET 10.0 with `RollForward: LatestMajor` for compatibility
- Warnings CA1416 (platform compatibility) suppressed intentionally
