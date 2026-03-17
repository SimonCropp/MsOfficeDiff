public class CommandParsingTests
{
    [Test]
    public async Task MissingBothArguments_ReturnsError()
    {
        using var console = new FakeInMemoryConsole();

        var app = Program.CreateBuilder()
            .UseConsole(console)
            .Build();

        var exitCode = await app.RunAsync([]);

        await Assert.That(exitCode).IsNotEqualTo(0);
        await Verify(new
        {
            Output = console.ReadOutputString(),
            Error = console.ReadErrorString()
        });
    }

    [Test]
    public async Task MissingSingleArgument_ReturnsError()
    {
        using var console = new FakeInMemoryConsole();

        var app = Program.CreateBuilder()
            .UseConsole(console)
            .Build();

        var exitCode = await app.RunAsync([ProjectFiles.input_temp_docx.FullPath]);

        await Assert.That(exitCode).IsNotEqualTo(0);
        await Verify(new
        {
            Output = console.ReadOutputString(),
            Error = console.ReadErrorString()
        });
    }

    [Test]
    public async Task NonExistentFile_ReturnsError()
    {
        using var console = new FakeInMemoryConsole();

        var app = Program.CreateBuilder()
            .UseConsole(console)
            .Build();

        var exitCode = await app.RunAsync(["nonexistent.docx", ProjectFiles.input_target_docx.FullPath]);

        await Assert.That(exitCode).IsNotEqualTo(0);
        await Verify(new
        {
            Output = console.ReadOutputString(),
            Error = console.ReadErrorString()
        });
    }

    [Test]
    public async Task HelpOption_ShowsHelp()
    {
        using var console = new FakeInMemoryConsole();

        var app = Program.CreateBuilder()
            .UseConsole(console)
            .Build();

        var exitCode = await app.RunAsync(["--help"]);

        await Assert.That(exitCode).IsEqualTo(0);
        await Verify(console.ReadOutputString());
    }
}
