# <img src="/src/icon.png" height="30px"> MsWordDiff

[![Build status](https://ci.appveyor.com/api/projects/status/sw0q9icdhhtau66m/branch/main?svg=true)](https://ci.appveyor.com/project/SimonCropp/MsWordDiff)
[![NuGet Status](https://img.shields.io/nuget/v/MsWordDiff.svg?label=MsWordDiff)](https://www.nuget.org/packages/MsWordDiff/)

A .NET tool that compares two Word documents side by side using Microsoft Word's built-in document comparison feature.<!-- singleLineInclude: intro. path: /src/intro.include.md -->

**See [Milestones](../../milestones?state=closed) for release notes.**


## Requirements

 * Windows
 * Microsoft Word installed
 * .NET 10.0 or later


## Installation

Install as a global tool:

```
dotnet tool install -g MsWordDiff
```

https://nuget.org/packages/MsWordDiff/



## Usage

```
diffword <path1> <path2> [--quiet] [--watch]
```

Where `<path1>` and `<path2>` are paths to the Word documents to compare.

Example:

```
diffword original.docx modified.docx
```

This will open Microsoft Word with a comparison view showing the differences between the two documents. The tool will wait until Word is closed before exiting.

<img src="/src/diff.png">


### Options

#### --quiet

Hide source documents in the comparison view, showing only the comparison document.

Example:

```
diffword original.docx modified.docx --quiet
```

#### --watch

Automatically refresh the comparison when source files change.

Example:

```
diffword original.docx modified.docx --watch
```

When enabled:
- Monitors both source files for changes
- Refreshes the comparison 500ms after the last detected change
- Preserves scroll position and zoom level
- Any edits to the comparison document will be discarded on refresh

Note: The comparison window must remain open for file watching to work.

Options can be combined:

```
diffword original.docx modified.docx --watch --quiet
```


### Configuration

The default behavior of options can be configured using settings commands.

#### View settings

```
diffword settings
```

This displays the settings file path and current settings. By default, settings are stored in:
```
%USERPROFILE%\.config\MsWordDiff\settings.json
```

#### Configure default Quiet mode

```
diffword set-quiet <true|false>
```

Set the default value for the Quiet option. When set to `true`, the source documents will be hidden by default.

Examples:

```
diffword set-quiet true
diffword set-quiet false
```

Note: Command-line options always override settings file values.


## How It Works

The tool uses COM automation to:

 1. Open both documents in Microsoft Word (read-only)
 2. Generate a comparison document using Word's `CompareDocuments` feature
 3. Display the comparison with tracked changes highlighting differences
 4. Automatically close Word when the parent process exits
