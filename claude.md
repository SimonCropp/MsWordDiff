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

## Testing

Uses TUnit framework. Tests are marked `[Explicit]` as they require Word to be installed.

```bash
dotnet run --project src/Tests
```

## Key Technical Details

- Uses dynamic COM interop (no Word interop assemblies required)
- Windows-only due to COM automation and Windows Forms dependency
- Targets .NET 10.0 with `RollForward: LatestMajor` for compatibility
- Warnings CA1416 (platform compatibility) suppressed intentionally
