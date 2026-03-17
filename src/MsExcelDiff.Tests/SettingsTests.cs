public class SettingsTests
{
    [Test]
    public async Task ReadNonExistentSettings_ReturnsDefaultSettings()
    {
        var tempPath = Path.GetTempFileName();
        File.Delete(tempPath);

        var manager = new SettingsManager(tempPath);
        var settings = await manager.Read();

        await Assert.That(settings.SpreadsheetComparePath).IsNull();
    }

    [Test]
    public async Task WriteAndReadSettings_PreservesValues()
    {
        var tempPath = Path.GetTempFileName();

        try
        {
            var manager = new SettingsManager(tempPath);
            var settings = new Settings { SpreadsheetComparePath = @"C:\Custom\SPREADSHEETCOMPARE.EXE" };

            await manager.Write(settings);

            var readSettings = await manager.Read();

            await Assert.That(readSettings.SpreadsheetComparePath).IsEqualTo(@"C:\Custom\SPREADSHEETCOMPARE.EXE");
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Test]
    public async Task ReadCorruptedSettings_ReturnsDefaultSettings()
    {
        var tempPath = Path.GetTempFileName();

        try
        {
            await File.WriteAllTextAsync(tempPath, "{ invalid json }");

            var manager = new SettingsManager(tempPath);
            var settings = await manager.Read();

            await Assert.That(settings.SpreadsheetComparePath).IsNull();
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Test]
    public async Task SetSpreadsheetComparePath_UpdatesSettings()
    {
        var tempPath = Path.GetTempFileName();

        try
        {
            var manager = new SettingsManager(tempPath);
            await manager.SetSpreadsheetComparePath(@"C:\Custom\SPREADSHEETCOMPARE.EXE");

            var settings = await manager.Read();

            await Assert.That(settings.SpreadsheetComparePath).IsEqualTo(@"C:\Custom\SPREADSHEETCOMPARE.EXE");
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Test]
    public async Task SetSpreadsheetComparePath_NullClearsPath()
    {
        var tempPath = Path.GetTempFileName();

        try
        {
            var manager = new SettingsManager(tempPath);
            await manager.SetSpreadsheetComparePath(@"C:\Custom\SPREADSHEETCOMPARE.EXE");
            await manager.SetSpreadsheetComparePath(null);

            var settings = await manager.Read();

            await Assert.That(settings.SpreadsheetComparePath).IsNull();
        }
        finally
        {
            File.Delete(tempPath);
        }
    }
}
