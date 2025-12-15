public class Test
{
    [Test]
    [Explicit]
    public void Launch() =>
        Program.Run(ProjectFiles.input_temp_docx, ProjectFiles.input_target_docx);
}