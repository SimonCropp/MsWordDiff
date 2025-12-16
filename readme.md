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


## Usage

```
diffword <path1> <path2>
```

Where `<path1>` and `<path2>` are paths to the Word documents to compare.

Example:

```
diffword original.docx modified.docx
```

This will open Microsoft Word with a comparison view showing the differences between the two documents. The tool will wait until Word is closed before exiting.


## NuGet

 * https://nuget.org/packages/MsWordDiff/

<img src="/src/diff.png">


## How It Works

The tool uses COM automation to:

 1. Open both documents in Microsoft Word (read-only)
 2. Generate a comparison document using Word's `CompareDocuments` feature
 3. Display the comparison with tracked changes highlighting differences
 4. Automatically close Word when the parent process exits
