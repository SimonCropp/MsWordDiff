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
    public void LaunchWithWatch() =>
        // This test requires manual verification:
        // 1. Launch with --watch
        // 2. Edit one of the source files in another Word instance
        // 3. Save the file
        // 4. Verify comparison refreshes after 500ms
        // 5. Verify scroll position and zoom are preserved
        Word.Launch(
            ProjectFiles.input_temp_docx.FullPath,
            ProjectFiles.input_target_docx.FullPath,
            quiet: false,
            watch: true);
}
