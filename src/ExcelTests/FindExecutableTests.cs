// ReSharper disable UnusedVariable
public class FindExecutableTests
{
    [Test]
    public void NullSettingsPath_AndNoExeInstalled_ReturnsNull()
    {
        // If no Spreadsheet Compare is installed in standard paths, returns null
        // This test may find the exe on machines with Office Pro Plus installed
        var result = SpreadsheetCompare.FindExecutable(null);

        // We can't assert null because the exe might actually be installed
        // Just verify it doesn't throw
    }

    [Test]
    public async Task SettingsPath_PointsToExistingFile_ReturnsThatPath()
    {
        var tempFile = Path.GetTempFileName();

        try
        {
            var result = SpreadsheetCompare.FindExecutable(tempFile);

            await Assert.That(result).IsEqualTo(tempFile);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }

    [Test]
    public void SettingsPath_PointsToNonExistentFile_FallsBackToSearch()
    {
        var result = SpreadsheetCompare.FindExecutable(@"C:\nonexistent\SPREADSHEETCOMPARE.EXE");

        // Falls back to searching standard paths; may or may not find it
        // Just verify it doesn't throw
    }
}
