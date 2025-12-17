public class CommandParsingTests
{
    [Test]
    public async Task ValidPaths_InvokesLaunchAction()
    {
        var wasCalled = false;
        string? capturedPath1 = null;
        string? capturedPath2 = null;

        var command = Program.BuildCommand((path1, path2) =>
        {
            wasCalled = true;
            capturedPath1 = path1;
            capturedPath2 = path2;
        });

        var result = command.Parse([
            ProjectFiles.input_temp_docx.FullPath,
            ProjectFiles.input_target_docx.FullPath
        ]).Invoke();

        await Assert.That(result).IsEqualTo(0);
        await Assert.That(wasCalled).IsTrue();
        await Assert.That(capturedPath1).IsEqualTo(ProjectFiles.input_temp_docx.FullPath);
        await Assert.That(capturedPath2).IsEqualTo(ProjectFiles.input_target_docx.FullPath);
    }

    [Test]
    public async Task MissingBothArguments_ReturnsErrorWithArgumentNames()
    {
        var command = Program.BuildCommand((_, _) =>
        {
        });

        var result = command.Parse([]);
        await Verify(result.Errors.Select(_ => _.Message));
    }

    [Test]
    public async Task MissingSingleArgument_ReturnsErrorWithArgumentName()
    {
        var command = Program.BuildCommand((_, _) =>
        {
        });

        var result = command.Parse([ProjectFiles.input_temp_docx.FullPath]);
        await Verify(result.Errors.Select(_ => _.Message));
    }

    [Test]
    public async Task NonExistentFile_ReturnsErrorWithFilePath()
    {
        var command = Program.BuildCommand((_, _) =>
        {
        });

        var result = command.Parse(["nonexistent.docx", ProjectFiles.input_target_docx.FullPath]);

        await Verify(result.Errors.Select(_ => _.Message));
    }

    [Test]
    public async Task HelpOption_ShowsHelp()
    {
        var wasCalled = false;
        var command = Program.BuildCommand((_, _) => wasCalled = true);

        var result = command.Parse(["--help"]).Invoke();

        await Assert.That(result).IsEqualTo(0);
        await Assert.That(wasCalled).IsFalse();
    }
}