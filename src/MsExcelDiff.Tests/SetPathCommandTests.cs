public class SetPathCommandTests
{
    [Test]
    public async Task SetPathCommand_WithExistingFile_UpdatesSettings()
    {
        var tempSettings = Path.GetTempFileName();
        var tempExe = Path.GetTempFileName();

        try
        {
            using var console = new FakeInMemoryConsole();

            var app = Program
                .CreateBuilder(tempSettings)
                .UseConsole(console)
                .Build();

            var exitCode = await app.RunAsync(["set-path", tempExe]);

            await Assert.That(exitCode).IsEqualTo(0);

            var output = console.ReadOutputString();
            await Assert.That(output).Contains("Spreadsheet Compare path set to:");
        }
        finally
        {
            File.Delete(tempSettings);
            File.Delete(tempExe);
        }
    }

    [Test]
    public async Task SetPathCommand_WithNonExistentFile_ReturnsError()
    {
        var tempSettings = Path.GetTempFileName();

        try
        {
            using var console = new FakeInMemoryConsole();

            var app = Program
                .CreateBuilder(tempSettings)
                .UseConsole(console)
                .Build();

            var exitCode = await app.RunAsync(["set-path", @"C:\nonexistent\SPREADSHEETCOMPARE.EXE"]);

            await Assert.That(exitCode).IsNotEqualTo(0);
        }
        finally
        {
            File.Delete(tempSettings);
        }
    }

    [Test]
    public async Task SettingsCommand_DisplaysPath()
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
