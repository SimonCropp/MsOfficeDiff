public class TempFilesTests
{
    string tempDir = null!;

    [Before(Test)]
    public void Setup() =>
        tempDir = Path.Combine(Path.GetTempPath(), $"MsExcelDiff_Test_{Guid.NewGuid():N}");

    [After(Test)]
    public void Cleanup()
    {
        try
        {
            Directory.Delete(tempDir, true);
        }
        catch
        {
        }
    }

    [Test]
    public async Task Create_WritesContent()
    {
        var path = TempFiles.Create(tempDir, "some content");

        await Assert.That(File.ReadAllText(path)).IsEqualTo("some content");
    }

    [Test]
    public async Task Create_CreatesFileInDirectory()
    {
        var path = TempFiles.Create(tempDir, "test");

        await Assert.That(Path.GetDirectoryName(path)).IsEqualTo(tempDir);
    }

    [Test]
    public async Task Create_UsesUniqueFilenames()
    {
        var path1 = TempFiles.Create(tempDir, "a");
        var path2 = TempFiles.Create(tempDir, "b");

        await Assert.That(path1).IsNotEqualTo(path2);
    }

    [Test]
    public async Task Create_CreatesDirectory()
    {
        TempFiles.Create(tempDir, "test");

        await Assert.That(Directory.Exists(tempDir)).IsTrue();
    }

    [Test]
    public async Task TryDelete_DeletesFile()
    {
        var path = TempFiles.Create(tempDir, "test");

        TempFiles.TryDelete(path);

        await Assert.That(File.Exists(path)).IsFalse();
    }

    [Test]
    public async Task TryDelete_ReturnsFalse()
    {
        var path = TempFiles.Create(tempDir, "test");

        var result = TempFiles.TryDelete(path);

        await Assert.That(result).IsFalse();
    }

    [Test]
    public void TryDelete_DoesNotThrow_WhenFileDoesNotExist() =>
        TempFiles.TryDelete(Path.Combine(tempDir, "nonexistent.txt"));

    [Test]
    public async Task CleanOld_DeletesOldFiles()
    {
        var path = TempFiles.Create(tempDir, "old");
        File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddDays(-2));

        TempFiles.CleanOld(tempDir);

        await Assert.That(File.Exists(path)).IsFalse();
    }

    [Test]
    public async Task CleanOld_KeepsRecentFiles()
    {
        var path = TempFiles.Create(tempDir, "recent");

        TempFiles.CleanOld(tempDir);

        await Assert.That(File.Exists(path)).IsTrue();
    }
}
