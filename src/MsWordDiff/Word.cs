public static partial class Word
{
    public static void Launch(string path1, string path2, bool quiet = false)
    {
        var wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            throw new("Microsoft Word is not installed");
        }

        var job = CreateJobToKillChildProcessesWhenThisProcessExits();
        dynamic? word = null;
        dynamic? compare = null;
        Process? wordProcess = null;

        try
        {
            // Capture existing Word processes before launching new instance
            var existingWordPids = Process.GetProcessesByName("WINWORD")
                .Select(p => p.Id)
                .ToHashSet();

            word = Activator.CreateInstance(wordType)!;

            // WdAlertLevel.wdAlertsNone = 0
            word.DisplayAlerts = 0;

            // Find and assign new Word process to job object immediately to prevent zombie processes
            // if this process is killed before normal cleanup
            wordProcess = FindNewWordProcess(existingWordPids);
            if (wordProcess == null)
            {
                Log.Warning("Could not find Word process to assign to job object early");
            }
            else
            {
                Log.Information("Found Word process {ProcessId}, assigning to job object", wordProcess.Id);
                var assigned = AssignProcessToJobObject(job, wordProcess.Handle);
                Log.Information("Word process assigned to job: {Success}", assigned);
            }

            var doc1 = Open(word, path1);
            var doc2 = Open(word, path2);

            // WdCompareDestination.wdCompareDestinationNew = 2
            // WdGranularity.wdGranularityWordLevel = 1
            compare = word.CompareDocuments(
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

            word.Visible = true;

            ApplyQuiet(quiet, word);

            HideNavigationPane(word);

            MinimizeRibbon(word);

            // Get process from Word's window handle as fallback if not found earlier
            var hwnd = (IntPtr)word.ActiveWindow.Hwnd;
            if (wordProcess == null)
            {
                GetWindowThreadProcessId(hwnd, out var processId);
                wordProcess = Process.GetProcessById(processId);
                AssignProcessToJobObject(job, wordProcess.Handle);
            }

            // Bring Word to the foreground
            SetForegroundWindow(hwnd);

            wordProcess.WaitForExit();
        }
        finally
        {
            CleanupWord(word, compare, wordProcess, job);
            RestoreRibbon(wordType);
        }
    }

    static Process? FindNewWordProcess(HashSet<int> existingPids)
    {
        // Wait for Word process to start (up to 2 seconds)
        for (var i = 0; i < 20; i++)
        {
            try
            {
                var allProcesses = Process.GetProcessesByName("WINWORD");
                var newProcesses = allProcesses
                    .Where(p => !existingPids.Contains(p.Id))
                    .ToList();

                if (newProcesses.Count > 0)
                {
                    // Return the most recently started process
                    var selected = newProcesses.OrderByDescending(p =>
                    {
                        try
                        {
                            return p.StartTime;
                        }
                        catch
                        {
                            return DateTime.MinValue;
                        }
                    }).First();

                    // Dispose all processes except the one we're returning
                    foreach (var process in allProcesses)
                    {
                        if (process.Id != selected.Id)
                        {
                            process.Dispose();
                        }
                    }

                    return selected;
                }

                // Dispose all if none matched
                foreach (var process in allProcesses)
                {
                    process.Dispose();
                }
            }
            catch
            {
                // Process may have exited or access denied
            }

            Thread.Sleep(100);
        }

        return null;
    }

    static void CleanupWord(dynamic? word, dynamic? compare, Process? wordProcess, IntPtr job)
    {
        // Release COM objects
        try
        {
            if (compare != null)
            {
                Marshal.ReleaseComObject(compare);
            }
        }
        catch
        {
            // Ignore errors during cleanup
        }

        try
        {
            if (word != null)
            {
                try
                {
                    // Try to quit Word gracefully
                    word.Quit(SaveChanges: false);
                }
                catch
                {
                    // Ignore if Word already closed
                }

                Marshal.ReleaseComObject(word);
            }
        }
        catch
        {
            // Ignore errors during cleanup
        }

        // Dispose process handle
        try
        {
            wordProcess?.Dispose();
        }
        catch
        {
            // Ignore errors during cleanup
        }

        // Close job object
        try
        {
            if (job != IntPtr.Zero)
            {
                CloseHandle(job);
            }
        }
        catch
        {
            // Ignore errors during cleanup
        }
    }

    static IntPtr CreateJobToKillChildProcessesWhenThisProcessExits()
    {
        var job = CreateJobObject(IntPtr.Zero, null);
        var info = new JOBOBJECT_EXTENDED_LIMIT_INFORMATION
        {
            BasicLimitInformation = new()
            {
                LimitFlags = jobObjectLimitKillOnJobClose
            }
        };
        SetInformationJobObject(job, jobObjectExtendedLimitInformation, ref info, (uint)Marshal.SizeOf(info));
        return job;
    }

    static void ApplyQuiet(bool quiet, dynamic word)
    {
        if (quiet)
        {
            // WdShowSourceDocuments.wdShowSourceDocumentsNone = 0
            // Hides the source documents, showing only the comparison
            word.ActiveWindow.ShowSourceDocuments = 0;
        }
        else
        {
            // WdShowSourceDocuments.wdShowSourceDocumentsBoth = 3
            // Shows the original and revised documents alongside the comparison
            word.ActiveWindow.ShowSourceDocuments = 3;
        }
    }

    static dynamic Open(dynamic word, string path)
    {
        var doc = word.Documents.Open(path, ReadOnly: true, AddToRecentFiles: false);
        // Hide document window to prevent flickering while preparing comparison
        doc.ActiveWindow.Visible = false;
        return doc;
    }

    static void HideNavigationPane(dynamic word) =>
        word.ActiveWindow.DocumentMap = false;

    static void MinimizeRibbon(dynamic word)
    {
        if (!word.CommandBars.GetPressedMso("MinimizeRibbon"))
        {
            word.CommandBars.ExecuteMso("MinimizeRibbon");
        }
    }

    static void RestoreRibbon(Type wordType)
    {
        dynamic word = Activator.CreateInstance(wordType)!;
        try
        {
            word.DisplayAlerts = 0;

            // Must be visible for settings to persist, but minimize to reduce flash
            // WdWindowState.wdWindowStateMinimize = 2
            word.WindowState = 2;
            word.Visible = true;

            if (word.CommandBars.GetPressedMso("MinimizeRibbon"))
            {
                word.CommandBars.ExecuteMso("MinimizeRibbon");
            }

            word.Quit();
        }
        finally
        {
            Marshal.ReleaseComObject(word);
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
