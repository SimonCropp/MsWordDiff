using System.Collections.Concurrent;
using System.Diagnostics;

public class ConcurrentLaunchTests
{
    string _processName = null!;
    string _exePath = null!;
    readonly List<Process> _processes = [];

    [Before(Test)]
    public void Setup()
    {
        _processName = $"test_{Guid.NewGuid():N}";
        _exePath = Path.Combine(Path.GetTempPath(), $"{_processName}.exe");
        File.Copy(@"C:\Windows\System32\cmd.exe", _exePath);
    }

    [After(Test)]
    public void Cleanup()
    {
        foreach (var p in _processes)
        {
            try
            {
                p.Kill();
            }
            catch
            {
            }

            p.Dispose();
        }

        _processes.Clear();

        try
        {
            File.Delete(_exePath);
        }
        catch
        {
        }
    }

    Process StartTestProcess()
    {
        var p = Process.Start(new ProcessStartInfo(_exePath, "/c ping -n 300 127.0.0.1 > nul")
        {
            CreateNoWindow = true,
            UseShellExecute = false
        })!;
        _processes.Add(p);
        return p;
    }

    [Test]
    public async Task GetProcessPids_ReturnsRunningProcessPids()
    {
        var p1 = StartTestProcess();
        var p2 = StartTestProcess();

        var pids = SpreadsheetCompare.GetProcessPids(_processName);

        await Assert.That(pids).Contains(p1.Id);
        await Assert.That(pids).Contains(p2.Id);
    }

    [Test]
    public async Task WaitForProcess_FindsNewProcess()
    {
        var p = StartTestProcess();

        using var found = SpreadsheetCompare.WaitForProcess(_processName, []);

        await Assert.That(found).IsNotNull();
        await Assert.That(found!.Id).IsEqualTo(p.Id);
    }

    [Test]
    public async Task WaitForProcess_SkipsExistingPids()
    {
        var existing = StartTestProcess();
        var newProcess = StartTestProcess();
        var existingPids = new HashSet<int>
        {
            existing.Id
        };

        using var found = SpreadsheetCompare.WaitForProcess(_processName, existingPids);

        await Assert.That(found).IsNotNull();
        await Assert.That(found!.Id).IsNotEqualTo(existing.Id);
    }

    [Test]
    public async Task WaitForProcess_ReturnsNull_WhenAllPidsExcluded()
    {
        var process = StartTestProcess();
        var existingPids = new HashSet<int>
        {
            process.Id
        };

        // Use maxAttempts=1 to avoid 10s timeout
        using var found = SpreadsheetCompare.WaitForProcess(_processName, existingPids, maxAttempts: 1);

        await Assert.That(found).IsNull();
    }

    [Test]
    public async Task SerializedIdentification_YieldsUniqueProcesses()
    {
        const int count = 5;
        var identifiedPids = new ConcurrentBag<int>();
        var mutexName = $@"Global\Test_{Guid.NewGuid():N}";

        // Simulate N concurrent diffexcel instances, each doing the
        // mutex-protected snapshot-launch-identify sequence.
        // The mutex ensures each snapshot sees previously identified processes.
        var tasks = Enumerable
            .Range(0, count)
            .Select(_ => Task.Run(() =>
            {
                using var mutex = new Mutex(false, mutexName);
                mutex.WaitOne();
                try
                {
                    var existing = SpreadsheetCompare.GetProcessPids(_processName);
                    StartTestProcess();
                    using var found = SpreadsheetCompare.WaitForProcess(_processName, existing);
                    if (found != null)
                    {
                        identifiedPids.Add(found.Id);
                    }
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
            }))
            .ToArray();

        await Task.WhenAll(tasks);

        await Assert.That(identifiedPids.Count).IsEqualTo(count);
        await Assert.That(identifiedPids.Distinct().Count()).IsEqualTo(count);
    }

    [Test]
    public async Task UnsynchronizedIdentification_CanYieldDuplicates()
    {
        // Demonstrates the bug: when all callers use the same PID snapshot
        // (as happens without mutex serialization), they all identify the
        // same process, leaving others orphaned.
        var snapshot = SpreadsheetCompare.GetProcessPids(_processName);

        const int count = 3;
        for (var i = 0; i < count; i++)
        {
            StartTestProcess();
        }

        var identifiedPids = new List<int>();
        for (var i = 0; i < count; i++)
        {
            using var found = SpreadsheetCompare.WaitForProcess(_processName, snapshot);
            if (found != null)
            {
                identifiedPids.Add(found.Id);
            }
        }

        // All callers find *some* process (all are "new" relative to the snapshot)
        await Assert.That(identifiedPids.Count).IsEqualTo(count);

        // But they all grab the same one - the first returned by GetProcessesByName
        // that isn't in the snapshot. This is the race condition.
        await Assert.That(identifiedPids.Distinct().Count()).IsEqualTo(1);
    }
}
