public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void InitializeOther()
    {
        VerifierSettings.InitializePlugins();
        VerifierSettings.ScrubLinesContaining("ExcelTests v");
    }
}
