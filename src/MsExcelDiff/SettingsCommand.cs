[Command("settings", Description = "Display the settings")]
public class SettingsPathCommand(SettingsManager settingsManager) : ICommand
{
    public async ValueTask ExecuteAsync(IConsole console)
    {
        await console.Output.WriteLineAsync(settingsManager.SettingsPath);
        if (File.Exists(settingsManager.SettingsPath))
        {
            await console.Output.WriteLineAsync(await File.ReadAllTextAsync(settingsManager.SettingsPath));
            return;
        }

        await console.Output.WriteLineAsync("No settings file found.");
    }
}
