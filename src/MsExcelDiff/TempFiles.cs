static class TempFiles
{
    static readonly string directory = Path.Combine(Path.GetTempPath(), "MsExcelDiff");

    public static string Create(string content) =>
        Create(directory, content);

    internal static string Create(string directory, string content)
    {
        Directory.CreateDirectory(directory);
        CleanOld(directory);
        var path = Path.Combine(directory, $"{Guid.NewGuid()}.txt");
        File.WriteAllText(path, content);
        return path;
    }

    public static bool TryDelete(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch
        {
            // Best effort; CleanOld will get it next time
        }

        return false;
    }

    internal static void CleanOld(string directory)
    {
        try
        {
            var cutoff = DateTime.UtcNow.AddDays(-1);
            foreach (var file in Directory.EnumerateFiles(directory))
            {
                try
                {
                    if (File.GetLastWriteTimeUtc(file) < cutoff)
                    {
                        File.Delete(file);
                    }
                }
                catch
                {
                    // File may be in use
                }
            }
        }
        catch
        {
            // Best effort cleanup
        }
    }
}
