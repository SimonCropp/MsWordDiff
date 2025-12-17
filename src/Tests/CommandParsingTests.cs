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
        var wasCalled = false;
        var errorOutput = new StringWriter();
        var command = Program.BuildCommand((_, _) => wasCalled = true, errorOutput);

        var exitCode = command.Parse([]).Invoke();

        var errors = errorOutput.ToString();
        await Assert.That(exitCode).IsNotEqualTo(0);
        await Assert.That(errors).Contains("Required argument missing: <path1>");
        await Assert.That(errors).Contains("Required argument missing: <path2>");
        await Assert.That(wasCalled).IsFalse();
    }

    [Test]
    public async Task MissingSingleArgument_ReturnsErrorWithArgumentName()
    {
        var wasCalled = false;
        var errorOutput = new StringWriter();
        var command = Program.BuildCommand((_, _) => wasCalled = true, errorOutput);

        var exitCode = command.Parse([ProjectFiles.input_temp_docx.FullPath]).Invoke();

        var errors = errorOutput.ToString();
        await Assert.That(exitCode).IsNotEqualTo(0);
        await Assert.That(errors).DoesNotContain("Required argument missing: <path1>");
        await Assert.That(errors).Contains("Required argument missing: <path2>");
        await Assert.That(wasCalled).IsFalse();
    }

    [Test]
    public async Task NonExistentFile_ReturnsErrorWithFilePath()
    {
        var wasCalled = false;
        var errorOutput = new StringWriter();
        var command = Program.BuildCommand((_, _) => wasCalled = true, errorOutput);

        var exitCode = command.Parse(["nonexistent.docx", ProjectFiles.input_target_docx.FullPath]).Invoke();

        var errors = errorOutput.ToString();
        await Assert.That(exitCode).IsNotEqualTo(0);
        await Assert.That(errors).Contains("File not found:");
        await Assert.That(errors).Contains("nonexistent.docx");
        await Assert.That(wasCalled).IsFalse();
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
