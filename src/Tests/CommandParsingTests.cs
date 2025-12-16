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

        var result = command.Invoke([
            ProjectFiles.input_temp_docx.FullPath,
            ProjectFiles.input_target_docx.FullPath
        ]);

        await Assert.That(result).IsEqualTo(0);
        await Assert.That(wasCalled).IsTrue();
        await Assert.That(capturedPath1).IsEqualTo(ProjectFiles.input_temp_docx.FullPath);
        await Assert.That(capturedPath2).IsEqualTo(ProjectFiles.input_target_docx.FullPath);
    }

    [Test]
    public async Task MissingArguments_ReturnsError()
    {
        var wasCalled = false;

        var command = Program.BuildCommand((_, _) => wasCalled = true);

        var result = command.Invoke([]);

        await Assert.That(result).IsNotEqualTo(0);
        await Assert.That(wasCalled).IsFalse();
    }

    [Test]
    public async Task SingleArgument_ReturnsError()
    {
        var wasCalled = false;

        var command = Program.BuildCommand((_, _) => wasCalled = true);

        var result = command.Invoke([ProjectFiles.input_temp_docx.FullPath]);

        await Assert.That(result).IsNotEqualTo(0);
        await Assert.That(wasCalled).IsFalse();
    }

    [Test]
    public async Task NonExistentFile_ReturnsError()
    {
        var wasCalled = false;

        var command = Program.BuildCommand((_, _) => wasCalled = true);

        var result = command.Invoke(["nonexistent.docx", ProjectFiles.input_target_docx.FullPath]);

        await Assert.That(result).IsNotEqualTo(0);
        await Assert.That(wasCalled).IsFalse();
    }

    [Test]
    public async Task HelpOption_ShowsHelp()
    {
        var wasCalled = false;

        var command = Program.BuildCommand((_, _) => wasCalled = true);

        var result = command.Invoke(["--help"]);

        await Assert.That(result).IsEqualTo(0);
        await Assert.That(wasCalled).IsFalse();
    }
}
