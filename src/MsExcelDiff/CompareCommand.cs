[Command(Description = "Compare two Excel workbooks using Microsoft Spreadsheet Compare")]
public class CompareCommand(SettingsManager settingsManager) : ICommand
{
    [CommandParameter(0, Name = "path1", Description = "Path to the first Excel workbook")]
    public required FileInfo Path1 { get; init; }

    [CommandParameter(1, Name = "path2", Description = "Path to the second Excel workbook")]
    public required FileInfo Path2 { get; init; }

    public async ValueTask ExecuteAsync(IConsole console)
    {
        if (!Path1.Exists)
        {
            throw new CommandException($"File does not exist: {Path1.FullName}");
        }

        if (!Path2.Exists)
        {
            throw new CommandException($"File does not exist: {Path2.FullName}");
        }

        var settings = await settingsManager.Read();

        try
        {
            SpreadsheetCompare.Launch(Path1.FullName, Path2.FullName, settings.SpreadsheetComparePath);
        }
        catch (Exception exception)
        {
            Log.Error(exception, "Failed to launch");
            throw new CommandException("Failed to launch", 1, innerException: exception);
        }
    }
}
