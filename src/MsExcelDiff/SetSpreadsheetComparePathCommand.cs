[Command("set-path", Description = "Set the path to SPREADSHEETCOMPARE.EXE. Settings are saved in the user configuration file.")]
public class SetSpreadsheetComparePathCommand(SettingsManager settingsManager) : ICommand
{
    [CommandParameter(0, Name = "value", Description = "Path to SPREADSHEETCOMPARE.EXE, or empty to clear and revert to auto-detection")]
    public string? Value { get; init; }

    public async ValueTask ExecuteAsync(IConsole console)
    {
        if (!string.IsNullOrWhiteSpace(Value) && !File.Exists(Value))
        {
            throw new CommandException($"File does not exist: {Value}");
        }

        var path = string.IsNullOrWhiteSpace(Value) ? null : Value;

        await settingsManager.SetSpreadsheetComparePath(path);

        if (path == null)
        {
            await console.Output.WriteLineAsync("Spreadsheet Compare path cleared. Auto-detection will be used.");
        }
        else
        {
            await console.Output.WriteLineAsync($"Spreadsheet Compare path set to: {path}");
        }

        await console.Output.WriteLineAsync($"Settings saved to: {settingsManager.SettingsPath}");
    }
}
