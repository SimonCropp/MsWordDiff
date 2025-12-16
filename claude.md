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

The codebase is minimal with a single main component:

- **Word.cs** - Core functionality using COM interop to automate Microsoft Word:
  - Opens documents read-only via `Word.Application` COM object
  - Uses `CompareDocuments` API to generate comparison view
  - Creates Windows Job Object to terminate Word when parent process exits
  - Configures Word window (minimizes ribbon, shows source documents and revision pane)

- **Program.cs** - CLI entry point, validates paths and invokes `Word.Launch()`

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
