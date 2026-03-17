public static partial class SpreadsheetCompare
{
    static readonly string[] searchPaths =
    [
        @"C:\Program Files\Microsoft Office\root\Office16\DCF\SPREADSHEETCOMPARE.EXE",
        @"C:\Program Files (x86)\Microsoft Office\root\Office16\DCF\SPREADSHEETCOMPARE.EXE",
        @"C:\Program Files\Microsoft Office\root\Office15\DCF\SPREADSHEETCOMPARE.EXE",
        @"C:\Program Files (x86)\Microsoft Office\root\Office15\DCF\SPREADSHEETCOMPARE.EXE"
    ];

    public static string? FindExecutable(string? settingsPath = null)
    {
        if (settingsPath != null && File.Exists(settingsPath))
        {
            return settingsPath;
        }

        foreach (var path in searchPaths)
        {
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

        var job = CreateJobToKillChildProcessesWhenThisProcessExits();
        string? tempFile = null;

        try
        {
            tempFile = Path.GetTempFileName();
            File.WriteAllText(tempFile, $"{path1}{Environment.NewLine}{path2}");

            var startInfo = new ProcessStartInfo
            {
                FileName = exe,
                Arguments = $"\"{tempFile}\"",
                UseShellExecute = false
            };

            using var process = Process.Start(startInfo)
                ?? throw new("Failed to start Spreadsheet Compare process");

            AssignProcessToJobObject(job, process.Handle);

            process.WaitForExit();

            // The launcher process exits quickly after spawning the real UI process.
            // Wait for all processes in the job to exit (i.e., the user closes the UI).
            WaitForJobToEmpty(job);
        }
        catch
        {
            // Clean up temp file on failure (the exe deletes it on success)
            if (tempFile != null && File.Exists(tempFile))
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

            throw;
        }
        finally
        {
            CloseHandle(job);
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

    static void WaitForJobToEmpty(IntPtr job)
    {
        while (true)
        {
            var info = new JOBOBJECT_BASIC_ACCOUNTING_INFORMATION();
            if (!QueryInformationJobObject(
                    job,
                    jobObjectBasicAccountingInformation,
                    ref info,
                    (uint)Marshal.SizeOf(info),
                    out _))
            {
                break;
            }

            if (info.ActiveProcesses == 0)
            {
                break;
            }

            Thread.Sleep(500);
        }
    }

    const uint jobObjectLimitKillOnJobClose = 0x2000;
    const int jobObjectExtendedLimitInformation = 9;
    const int jobObjectBasicAccountingInformation = 1;

    [StructLayout(LayoutKind.Sequential)]
    struct JOBOBJECT_BASIC_ACCOUNTING_INFORMATION
    {
        public long TotalUserTime;
        public long TotalKernelTime;
        public long ThisPeriodTotalUserTime;
        public long ThisPeriodTotalKernelTime;
        public uint TotalPageFaultCount;
        public uint TotalProcesses;
        public uint ActiveProcesses;
        public uint TotalTerminatedProcesses;
    }

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
    private static partial bool QueryInformationJobObject(IntPtr hJob, int jobObjectInfoClass, ref JOBOBJECT_BASIC_ACCOUNTING_INFORMATION lpJobObjectInfo, uint cbJobObjectInfoLength, out uint lpReturnLength);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool CloseHandle(IntPtr hObject);
}
