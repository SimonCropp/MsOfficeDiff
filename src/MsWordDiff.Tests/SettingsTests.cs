public class SettingsTests
{
    [Test]
    [Explicit]
    public Task Setup()
    {
        var tempPath = TempFile.Create();

        File.Delete(tempPath);
        var manager = new SettingsManager(tempPath);
        return manager.Setup();
    }

    [Test]
    public async Task ReadNonExistentSettings_ReturnsDefaultSettings()
    {
        var tempPath = TempFile.Create();

        var manager = new SettingsManager(tempPath);
        var settings = await manager.Read();

        await Assert.That(settings.Quiet).IsFalse();
    }

    [Test]
    public async Task WriteAndReadSettings_PreservesValues()
    {
        var tempPath = TempFile.Create();

        var manager = new SettingsManager(tempPath);
        var settings = new Settings {Quiet = true};

        await manager.Write(settings);

        var readSettings = await manager.Read();

        await Assert.That(readSettings.Quiet).IsTrue();
    }

    [Test]
    public async Task ReadCorruptedSettings_ReturnsDefaultSettings()
    {
        var tempPath = TempFile.Create();

        await File.WriteAllTextAsync(tempPath, "{ invalid json }");

        var manager = new SettingsManager(tempPath);
        var settings = await manager.Read();

        await Assert.That(settings.Quiet).IsFalse();
    }
}