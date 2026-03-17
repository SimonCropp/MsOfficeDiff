public class LaunchTests
{
    [Test]
    [Explicit]
    public void Launch() =>
        SpreadsheetCompare.Launch(
            ProjectFiles.input_source_xlsx.FullPath,
            ProjectFiles.input_target_xlsx.FullPath);

    [Test]
    [Explicit]
    public async Task LaunchViaProgram() =>
        await Program.Main([
            ProjectFiles.input_source_xlsx.FullPath,
            ProjectFiles.input_target_xlsx.FullPath
        ]);
}
