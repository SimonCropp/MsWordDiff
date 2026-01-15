public class Test
{
    [Test]
    [Explicit]
    public void Launch() =>
        Word.Launch(
            ProjectFiles.input_temp_docx.FullPath,
            ProjectFiles.input_target_docx.FullPath);

    [Test]
    [Explicit]
    public void LaunchQuiet() =>
        Word.Launch(
            ProjectFiles.input_temp_docx.FullPath,
            ProjectFiles.input_target_docx.FullPath,
            quiet: true);

    [Test]
    [Explicit]
    public async Task LaunchViaProgram() =>
        await Program.Main([
            ProjectFiles.input_temp_docx.FullPath,
            ProjectFiles.input_target_docx.FullPath
        ]);

    [Test]
    [Explicit]
    public async Task WordProcessKilledWhenParentKilled()
    {
        // This test verifies that Word doesn't become a zombie process when
        // MsWordDiff is forcefully killed (e.g., by DiffEngineTray accepting changes)

        // Get existing Word PIDs before test
        var existingWordPids = Process.GetProcessesByName("WINWORD")
            .Select(p => p.Id)
            .ToHashSet();

        // Launch MsWordDiff in a separate process
        // Use the compiled executable directly instead of 'dotnet run' to avoid process hierarchy issues
        var exePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../MsWordDiff/bin/Debug/net10.0/MsWordDiff.exe"));

        Console.WriteLine($"Test diagnostics:");
        Console.WriteLine($"  Exe path: {exePath}");
        Console.WriteLine($"  Exe exists: {File.Exists(exePath)}");
        Console.WriteLine($"  Temp docx: {ProjectFiles.input_temp_docx.FullPath}");
        Console.WriteLine($"  Target docx: {ProjectFiles.input_target_docx.FullPath}");

        var output = new System.Text.StringBuilder();
        var error = new System.Text.StringBuilder();

        var msWordDiffProcess = new Process
        {
            StartInfo = new()
            {
                FileName = exePath,
                Arguments = $"\"{ProjectFiles.input_temp_docx.FullPath}\" \"{ProjectFiles.input_target_docx.FullPath}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            }
        };

        msWordDiffProcess.OutputDataReceived += (sender, e) =>
        {
            if (e.Data != null)
            {
                output.AppendLine(e.Data);
                Console.WriteLine($"[MsWordDiff OUT] {e.Data}");
            }
        };

        msWordDiffProcess.ErrorDataReceived += (sender, e) =>
        {
            if (e.Data != null)
            {
                error.AppendLine(e.Data);
                Console.WriteLine($"[MsWordDiff ERR] {e.Data}");
            }
        };

        msWordDiffProcess.Start();
        msWordDiffProcess.BeginOutputReadLine();
        msWordDiffProcess.BeginErrorReadLine();

        Console.WriteLine($"  MsWordDiff PID: {msWordDiffProcess.Id}");

        Process? wordProcess = null;
        try
        {
            // Wait for Word process to start (up to 10 seconds)
            var wordStarted = false;
            for (var i = 0; i < 100; i++)
            {
                // Check if MsWordDiff exited prematurely
                if (msWordDiffProcess.HasExited)
                {
                    Console.WriteLine($"MsWordDiff exited prematurely with code {msWordDiffProcess.ExitCode}");
                    Console.WriteLine($"Output: {output}");
                    Console.WriteLine($"Error: {error}");
                    break;
                }

                var newWordProcesses = Process.GetProcessesByName("WINWORD")
                    .Where(p => !existingWordPids.Contains(p.Id))
                    .ToList();

                if (newWordProcesses.Count > 0)
                {
                    wordProcess = newWordProcesses.First();
                    wordStarted = true;
                    Console.WriteLine($"  Word PID: {wordProcess.Id}");
                    break;
                }

                await Task.Delay(100);
            }

            if (!wordStarted)
            {
                Console.WriteLine($"Word did not start after 10 seconds");
                Console.WriteLine($"MsWordDiff still running: {!msWordDiffProcess.HasExited}");
                if (msWordDiffProcess.HasExited)
                {
                    Console.WriteLine($"MsWordDiff exit code: {msWordDiffProcess.ExitCode}");
                }
                Console.WriteLine($"Output: {output}");
                Console.WriteLine($"Error: {error}");
            }

            await Assert.That(wordStarted).IsTrue().Because("Word process should start");

            // Give Word a moment to initialize and be assigned to job object
            await Task.Delay(500);

            Console.WriteLine($"Killing MsWordDiff process...");
            // Kill the MsWordDiff process (simulating DiffEngineTray killing it)
            msWordDiffProcess.Kill();
            await msWordDiffProcess.WaitForExitAsync();
            Console.WriteLine($"MsWordDiff killed");

            // Verify Word process also exits within 10 seconds (due to Job Object)
            Console.WriteLine($"Waiting for Word to exit...");
            Console.WriteLine($"Word HasExited before wait: {wordProcess!.HasExited}");

            var wordExited = wordProcess.WaitForExit(10000);
            Console.WriteLine($"Word exited after WaitForExit: {wordExited}");
            Console.WriteLine($"Word HasExited property: {wordProcess.HasExited}");

            // Double-check by trying to access the process
            try
            {
                var stillRunning = Process.GetProcessById(wordProcess.Id);
                Console.WriteLine($"Word process still accessible by ID: {stillRunning.HasExited}");
            }
            catch (ArgumentException)
            {
                Console.WriteLine($"Word process no longer exists (good!)");
                wordExited = true; // Process doesn't exist anymore
            }

            await Assert.That(wordExited).IsTrue().Because("Word should be killed by Job Object when parent is killed");
        }
        finally
        {
            // Cleanup: ensure processes are terminated
            try
            {
                if (!msWordDiffProcess.HasExited)
                {
                    msWordDiffProcess.Kill();
                }
            }
            catch
            {
                // Ignore
            }

            msWordDiffProcess.Dispose();

            try
            {
                if (wordProcess is { HasExited: false })
                {
                    wordProcess.Kill();
                    wordProcess.WaitForExit(1000);
                }
            }
            catch
            {
                // Ignore
            }

            wordProcess?.Dispose();
        }
    }
}
