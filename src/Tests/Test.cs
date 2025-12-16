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
    public void LaunchViaProgram() =>
        Program.Main([
                ProjectFiles.input_temp_docx.FullPath,
                ProjectFiles.input_target_docx.FullPath
            ]);
}