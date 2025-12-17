public static partial class Word
{
    static volatile bool refreshRequested;

    public static void Launch(string path1, string path2, bool quiet = false, bool watch = false)
    {
        var wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            throw new("Microsoft Word is not installed");
        }

        // Create a job object that kills child processes when this process exits
        var job = CreateJobObject(IntPtr.Zero, null);
        var info = new JOBOBJECT_EXTENDED_LIMIT_INFORMATION
        {
            BasicLimitInformation = new()
            {
                LimitFlags = jobObjectLimitKillOnJobClose
            }
        };
        SetInformationJobObject(job, jobObjectExtendedLimitInformation, ref info, (uint)Marshal.SizeOf(info));

        dynamic word = Activator.CreateInstance(wordType)!;

        // WdAlertLevel.wdAlertsNone = 0
        word.DisplayAlerts = 0;

        var compare = CreateComparison(word, path1, path2);

        word.Visible = true;

        if (!quiet)
        {
            // WdShowSourceDocuments.wdShowSourceDocumentsBoth = 3
            // Shows the original and revised documents alongside the comparison
            word.ActiveWindow.ShowSourceDocuments = 3;
        }

        MinimizeRibbon(word);

        // Get process from Word's window handle and assign to job
        var hwnd = (IntPtr)word.ActiveWindow.Hwnd;
        GetWindowThreadProcessId(hwnd, out var processId);
        using var process = Process.GetProcessById(processId);
        AssignProcessToJobObject(job, process.Handle);

        // Bring Word to the foreground
        SetForegroundWindow(hwnd);

        if (watch)
        {
            compare = RunWithFileWatching(word, compare, path1, path2, quiet, process);
        }
        else
        {
            process.WaitForExit();
        }

        Marshal.ReleaseComObject(compare);
        Marshal.ReleaseComObject(word);
        CloseHandle(job);
    }

    static dynamic CreateComparison(dynamic word, string path1, string path2)
    {
        // Create temporary copies to avoid interfering with files open in other Word instances
        var temp1 = Path.Combine(Path.GetTempPath(), $"diffword-{Guid.NewGuid()}{Path.GetExtension(path1)}");
        var temp2 = Path.Combine(Path.GetTempPath(), $"diffword-{Guid.NewGuid()}{Path.GetExtension(path2)}");

        try
        {
            File.Copy(path1, temp1, overwrite: true);
            File.Copy(path2, temp2, overwrite: true);

            var doc1 = Open(word, temp1);
            var doc2 = Open(word, temp2);

            // WdCompareDestination.wdCompareDestinationNew = 2
            // WdGranularity.wdGranularityWordLevel = 1
            var compare = word.CompareDocuments(
                doc1,
                doc2,
                Destination: 2,
                Granularity: 1,
                CompareFormatting: true,
                CompareCaseChanges: true,
                CompareWhitespace: true,
                CompareTables: true,
                CompareHeaders: true,
                CompareFootnotes: true,
                CompareTextboxes: true,
                CompareFields: true,
                CompareComments: true,
                CompareMoves: true,
                RevisedAuthor: "",
                IgnoreAllComparisonWarnings: true);

            doc1.Close(SaveChanges: false);
            doc2.Close(SaveChanges: false);

            // Mark as saved so Word won't prompt to save on close
            compare.Saved = true;

            compare.AutoSaveOn = false;
            compare.ShowSpellingErrors = false;
            compare.ShowGrammaticalErrors = false;

            return compare;
        }
        finally
        {
            // Clean up temporary files
            if (File.Exists(temp1))
            {
                try { File.Delete(temp1); } catch { /* Ignore cleanup errors */ }
            }
            if (File.Exists(temp2))
            {
                try { File.Delete(temp2); } catch { /* Ignore cleanup errors */ }
            }
        }
    }

    static dynamic Open(dynamic word, string path) =>
        word.Documents.Open(path, ReadOnly: true, AddToRecentFiles: false);

    struct ViewState
    {
        public int ScrollTop;
        public int ZoomPercentage;
    }

    static ViewState SaveViewState(dynamic window) =>
        new()
        {
            ScrollTop = window.VerticalPercentScrolled,
            ZoomPercentage = window.View.Zoom.Percentage
        };

    static void RestoreViewState(dynamic window, ViewState state)
    {
        window.View.Zoom.Percentage = state.ZoomPercentage;
        window.VerticalPercentScrolled = state.ScrollTop;
    }

    static dynamic RunWithFileWatching(dynamic word, dynamic currentCompare, string path1, string path2, bool quiet, Process process)
    {
        using var fileWatcher = new FileWatcherManager(path1, path2, () =>
        {
            refreshRequested = true;
        });

        while (!process.HasExited)
        {
            if (refreshRequested)
            {
                refreshRequested = false;
                currentCompare = RefreshComparison(word, currentCompare, path1, path2, quiet);
            }
            Thread.Sleep(100);
        }

        return currentCompare;
    }

    static dynamic RefreshComparison(dynamic word, dynamic oldCompare, string path1, string path2, bool quiet)
    {
        dynamic? newCompare = null;
        try
        {
            var viewState = SaveViewState(word.ActiveWindow);

            // Create new comparison first before releasing old one
            newCompare = CreateComparison(word, path1, path2);

            // Only close and release old compare after new one is successfully created
            word.ActiveDocument.Close(SaveChanges: false);

            if (!quiet)
            {
                word.ActiveWindow.ShowSourceDocuments = 3;
            }

            RestoreViewState(word.ActiveWindow, viewState);

            Log.Information("Comparison refreshed");

            Marshal.ReleaseComObject(oldCompare);
            return newCompare;
        }
        catch (Exception ex)
        {
            if (newCompare != null)
            {
                Marshal.ReleaseComObject(newCompare);
            }

            Log.Warning(ex, "Failed to refresh comparison");
            // Return old compare if refresh fails (safe since we haven't released it yet)
            return oldCompare;
        }
    }

    static void MinimizeRibbon(dynamic word)
    {
        if (!word.CommandBars.GetPressedMso("MinimizeRibbon"))
        {
            word.CommandBars.ExecuteMso("MinimizeRibbon");
        }
    }

    const uint jobObjectLimitKillOnJobClose = 0x2000;
    const int jobObjectExtendedLimitInformation = 9;

    [StructLayout(LayoutKind.Sequential)]
    struct JOBOBJECT_BASIC_LIMIT_INFORMATION
    {
        public long PerProcessUserTimeLimit;
        public long PerJobUserTimeLimit;
        public uint LimitFlags;
        public nuint MinimumWorkingSetSize;
        public nuint MaximumWorkingSetSize;
        public uint ActiveProcessLimit;
        public nuint Affinity;
        public uint PriorityClass;
        public uint SchedulingClass;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct IO_COUNTERS
    {
        public ulong ReadOperationCount;
        public ulong WriteOperationCount;
        public ulong OtherOperationCount;
        public ulong ReadTransferCount;
        public ulong WriteTransferCount;
        public ulong OtherTransferCount;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct JOBOBJECT_EXTENDED_LIMIT_INFORMATION
    {
        public JOBOBJECT_BASIC_LIMIT_INFORMATION BasicLimitInformation;
        public IO_COUNTERS IoInfo;
        public nuint ProcessMemoryLimit;
        public nuint JobMemoryLimit;
        public nuint PeakProcessMemoryUsed;
        public nuint PeakJobMemoryUsed;
    }

    [LibraryImport("user32.dll")]
    internal static partial uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool SetForegroundWindow(IntPtr hWnd);

    [LibraryImport("kernel32.dll", EntryPoint = "CreateJobObjectW", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
    private static partial IntPtr CreateJobObject(IntPtr lpJobAttributes, string? lpName);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool SetInformationJobObject(IntPtr hJob, int jobObjectInfoClass, ref JOBOBJECT_EXTENDED_LIMIT_INFORMATION lpJobObjectInfo, uint cbJobObjectInfoLength);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool AssignProcessToJobObject(IntPtr hJob, IntPtr hProcess);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool CloseHandle(IntPtr hObject);
}