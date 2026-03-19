public static partial class SpreadsheetCompare
{
    static readonly string[] programFolders =
    [
        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
    ];

    static readonly string[] searchRelativePaths =
    [
        @"Microsoft Office\root\Office16\DCF\SPREADSHEETCOMPARE.EXE",
        @"Microsoft Office\root\Office15\DCF\SPREADSHEETCOMPARE.EXE",
        // Click-to-Run installs place the exe inside a virtual filesystem (vfs) directory
        // rather than the standard Office16/Office15 location
        @"Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE",
        @"Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office15\DCF\SPREADSHEETCOMPARE.EXE",
        @"Microsoft Office\root\vfs\ProgramFilesX64\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE",
        @"Microsoft Office\root\vfs\ProgramFilesX64\Microsoft Office\Office15\DCF\SPREADSHEETCOMPARE.EXE"
    ];

    public static string? FindExecutable(string? settingsPath = null)
    {
        if (settingsPath != null && File.Exists(settingsPath))
        {
            return settingsPath;
        }

        foreach (var folder in programFolders)
        {
            foreach (var relative in searchRelativePaths)
            {
                var path = Path.Combine(folder, relative);
                if (File.Exists(path))
                {
                    return path;
                }
            }
        }

        return null;
    }

    static string? FindAppVlp()
    {
        foreach (var folder in programFolders)
        {
            var path = Path.Combine(folder, @"Microsoft Office\root\Client\AppVLP.exe");
            if (File.Exists(path))
            {
                return path;
            }
        }

        return null;
    }

    public static void Launch(string path1, string path2, string? exePath = null)
    {
        var exe = FindExecutable(exePath);
        if (exe == null)
        {
            throw new("Spreadsheet Compare (SPREADSHEETCOMPARE.EXE) was not found. " +
                       "It is included with Office Professional Plus / Microsoft 365 Apps for Enterprise. " +
                       "If installed in a custom location, use the 'set-path' command to configure the path.");
        }

        var tempFile = Path.GetTempFileName();
        File.WriteAllText(tempFile, $"{path1}{Environment.NewLine}{path2}");

        var job = CreateJobToKillChildProcessesWhenThisProcessExits();

        try
        {
            // Click-to-Run Office installs require launching via AppVLP.exe (the App-V
            // virtualization layer). SPREADSHEETCOMPARE.EXE crashes if launched directly.
            var appVlp = FindAppVlp();
            ProcessStartInfo startInfo;

            if (appVlp != null)
            {
                startInfo = new()
                {
                    FileName = appVlp,
                    Arguments = $"\"{exe}\" {tempFile}",
                    UseShellExecute = false
                };
            }
            else
            {
                // Non-Click-to-Run install: launch directly
                startInfo = new()
                {
                    FileName = exe,
                    Arguments = tempFile,
                    UseShellExecute = true
                };
            }

            // Serialize the snapshot-launch-identify sequence across concurrent
            // diffexcel instances. Without this, concurrent instances snapshot the
            // same PID set, race to claim the same SPREADSHEETCOMPARE process, and
            // leave others orphaned (not in any job object, so they survive when
            // diffexcel is killed).
            Process? uiProcess;
            using (var mutex = new Mutex(false, @"Global\MsExcelDiff_Launch"))
            {
                mutex.WaitOne();
                try
                {
                    var existingPids = GetSpreadsheetComparePids();

                    using var launcher = Process.Start(startInfo)
                        ?? throw new("Failed to start Spreadsheet Compare process");

                    // AppVLP.exe is a launcher that exits after starting the real process.
                    // Find the actual SPREADSHEETCOMPARE process and wait on it.
                    launcher.WaitForExit();

                    uiProcess = WaitForProcess(existingPids);

                    if (uiProcess != null)
                    {
                        AssignProcessToJobObject(job, uiProcess.Handle);
                    }
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
            }

            if (uiProcess == null)
            {
                throw new("Spreadsheet Compare did not start. Ensure the application is installed correctly.");
            }

            try
            {
                uiProcess.WaitForExit();
            }
            finally
            {
                uiProcess.Dispose();
            }
        }
        catch when (TryDeleteTempFile(tempFile))
        {
            // unreachable: TryDeleteTempFile always returns false
            throw;
        }
        finally
        {
            CloseHandle(job);
        }
    }

    static bool TryDeleteTempFile(string tempFile)
    {
        if (File.Exists(tempFile))
        {
            try
            {
                File.Delete(tempFile);
            }
            catch
            {
                // Best effort cleanup
            }
        }

        return false;
    }

    static HashSet<int> GetSpreadsheetComparePids() =>
        GetProcessPids("SPREADSHEETCOMPARE");

    internal static HashSet<int> GetProcessPids(string processName)
    {
        var processes = Process.GetProcessesByName(processName);
        var pids = processes.Select(p => p.Id).ToHashSet();
        foreach (var p in processes)
        {
            p.Dispose();
        }

        return pids;
    }

    static Process? WaitForProcess(HashSet<int> existingPids) =>
        WaitForProcess("SPREADSHEETCOMPARE", existingPids);

    internal static Process? WaitForProcess(string processName, HashSet<int> existingPids, int maxAttempts = 100)
    {
        for (var i = 0; i < maxAttempts; i++)
        {
            var processes = Process.GetProcessesByName(processName);
            Process? result = null;
            foreach (var p in processes)
            {
                if (result == null && !existingPids.Contains(p.Id))
                {
                    result = p;
                }
                else
                {
                    p.Dispose();
                }
            }

            if (result != null)
            {
                return result;
            }

            Thread.Sleep(100);
        }

        return null;
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
