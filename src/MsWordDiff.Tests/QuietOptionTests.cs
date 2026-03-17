public class QuietOptionTests
{
    [Test]
    public async Task QuietOption_ShowsInHelp()
    {
        using var console = new FakeInMemoryConsole();

        var app = Program
            .CreateBuilder()
            .UseConsole(console)
            .Build();

        var exitCode = await app.RunAsync(["--help"]);

        var output = console.ReadOutputString();

        await Assert.That(exitCode).IsEqualTo(0);
        await Verify(output);
    }

    [Test]
    public async Task SetQuietCommand_UpdatesSettings()
    {
        using var tempPath = TempFile.Create();

        using var console = new FakeInMemoryConsole();

        var app = Program
            .CreateBuilder(tempPath)
            .UseConsole(console)
            .Build();

        var exitCode = await app.RunAsync(["set-quiet", "true"]);

        await Assert.That(exitCode).IsEqualTo(0);

        var output = console.ReadOutputString();
        await Assert.That(output).Contains("Quiet mode default set to: True");
    }

    [Test]
    public async Task SettingsPathCommand_DisplaysPath()
    {
        using var console = new FakeInMemoryConsole();

        var app = Program
            .CreateBuilder()
            .UseConsole(console)
            .Build();

        var exitCode = await app.RunAsync(["settings"]);

        await Assert.That(exitCode).IsEqualTo(0);

        var output = console.ReadOutputString();
        await Assert.That(output).Contains(SettingsManager.DefaultSettingsPath);
    }
}
