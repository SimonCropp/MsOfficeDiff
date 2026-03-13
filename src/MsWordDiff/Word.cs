public static partial class Word
{
    public static void Launch(string path1, string path2, bool quiet = false)
    {
        var wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            throw new("Microsoft Word is not installed");
        }

        var job = CreateJobToKillChildProcessesWhenThisProcessExits();

        dynamic word = Activator.CreateInstance(wordType)!;

        // WdAlertLevel.wdAlertsNone = 0
        word.DisplayAlerts = 0;

        // Disable AutoRecover to prevent "serious error" recovery dialogs
        word.Options.SaveInterval = 0;

        string tempPath1;
        var doc1 = Open(word, path1, out tempPath1);

        // Get process from Word's window handle and assign to job
        var hwnd = (IntPtr) word.ActiveWindow.Hwnd;
        GetWindowThreadProcessId(hwnd, out var processId);
        using var process = Process.GetProcessById(processId);
        AssignProcessToJobObject(job, process.Handle);

        string tempPath2;
        var doc2 = Open(word, path2, out tempPath2);

        var compare = LaunchCompare(word, doc1, doc2, tempPath1, tempPath2);

        word.Visible = true;

        ApplyQuiet(quiet, word);

        HideNavigationPane(word);

        MinimizeRibbon(word);

        // Bring Word to the foreground
        SetForegroundWindow(hwnd);

        process.WaitForExit();

        Marshal.ReleaseComObject(compare);
        Marshal.ReleaseComObject(word);
        CloseHandle(job);

        RestoreRibbon(wordType);
    }

    internal static dynamic LaunchCompare(dynamic word, dynamic doc1, dynamic doc2, string tempPath1, string tempPath2)
    {
        // WdCompareDestination.wdCompareDestinationNew = 2
        // WdGranularity.wdGranularityWordLevel = 1
        var compare = word.CompareDocuments(
            doc1,
            doc2,
            Destination: 2,
            Granularity: 1,
            CompareFormatting: true,
            CompareCaseChanges: true,
            CompareWhitespace: true,
            CompareTables: true,
            CompareHeaders: true,
            CompareFootnotes: true,
            CompareTextboxes: true,
            CompareFields: true,
            CompareComments: true,
            CompareMoves: true,
            RevisedAuthor: "",
            IgnoreAllComparisonWarnings: true);

        doc1.Close(SaveChanges: false);
        doc2.Close(SaveChanges: false);

        File.Delete(tempPath1);
        File.Delete(tempPath2);

        // Mark as saved so Word won't prompt to save on close
        compare.Saved = true;

        compare.AutoSaveOn = false;
        compare.ShowSpellingErrors = false;
        compare.ShowGrammaticalErrors = false;
        return compare;
    }

    static IntPtr CreateJobToKillChildProcessesWhenThisProcessExits()
    {
        var job = CreateJobObject(IntPtr.Zero, null);
        var info = new JOBOBJECT_EXTENDED_LIMIT_INFORMATION
        {
            BasicLimitInformation = new()
            {
                LimitFlags = jobObjectLimitKillOnJobClose
            }
        };
        SetInformationJobObject(job, jobObjectExtendedLimitInformation, ref info, (uint)Marshal.SizeOf(info));
        return job;
    }

    internal static void ApplyQuiet(bool quiet, dynamic word)
    {
        if (quiet)
        {
            // WdShowSourceDocuments.wdShowSourceDocumentsNone = 0
            // Hides the source documents, showing only the comparison
            word.ActiveWindow.ShowSourceDocuments = 0;
        }
        else
        {
            // WdShowSourceDocuments.wdShowSourceDocumentsBoth = 3
            // Shows the original and revised documents alongside the comparison
            word.ActiveWindow.ShowSourceDocuments = 3;
        }
    }

    internal static dynamic Open(dynamic word, string path, out string tempFilePath)
    {
        // Copy to a temp file so the original is never locked by Word
        tempFilePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}{Path.GetExtension(path)}");
        File.Copy(path, tempFilePath);

        var doc = word.Documents.Open(
            tempFilePath,
            ConfirmConversions: false,
            ReadOnly: true,
            AddToRecentFiles: false,
            OpenAndRepair: false,
            NoEncodingDialog: true);
        // Hide document window to prevent flickering while preparing comparison
        doc.ActiveWindow.Visible = false;
        return doc;
    }

static void HideNavigationPane(dynamic word) =>
        word.ActiveWindow.DocumentMap = false;

    static void MinimizeRibbon(dynamic word)
    {
        if (!word.CommandBars.GetPressedMso("MinimizeRibbon"))
        {
            word.CommandBars.ExecuteMso("MinimizeRibbon");
        }
    }

    static void RestoreRibbon(Type wordType)
    {
        dynamic word = Activator.CreateInstance(wordType)!;
        try
        {
            word.DisplayAlerts = 0;

            // Must be visible for settings to persist, but minimize to reduce flash
            // WdWindowState.wdWindowStateMinimize = 2
            word.WindowState = 2;
            word.Visible = true;

            if (word.CommandBars.GetPressedMso("MinimizeRibbon"))
            {
                word.CommandBars.ExecuteMso("MinimizeRibbon");
            }

            word.Quit();
        }
        finally
        {
            Marshal.ReleaseComObject(word);
        }
    }

    const uint jobObjectLimitKillOnJobClose = 0x2000;
    const int jobObjectExtendedLimitInformation = 9;

    [StructLayout(LayoutKind.Sequential)]
    struct JOBOBJECT_BASIC_LIMIT_INFORMATION
    {
        public long PerProcessUserTimeLimit;
        public long PerJobUserTimeLimit;
        public uint LimitFlags;
        public nuint MinimumWorkingSetSize;
        public nuint MaximumWorkingSetSize;
        public uint ActiveProcessLimit;
        public nuint Affinity;
        public uint PriorityClass;
        public uint SchedulingClass;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct IO_COUNTERS
    {
        public ulong ReadOperationCount;
        public ulong WriteOperationCount;
        public ulong OtherOperationCount;
        public ulong ReadTransferCount;
        public ulong WriteTransferCount;
        public ulong OtherTransferCount;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct JOBOBJECT_EXTENDED_LIMIT_INFORMATION
    {
        public JOBOBJECT_BASIC_LIMIT_INFORMATION BasicLimitInformation;
        public IO_COUNTERS IoInfo;
        public nuint ProcessMemoryLimit;
        public nuint JobMemoryLimit;
        public nuint PeakProcessMemoryUsed;
        public nuint PeakJobMemoryUsed;
    }

    [LibraryImport("user32.dll")]
    internal static partial uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool SetForegroundWindow(IntPtr hWnd);

    [LibraryImport("kernel32.dll", EntryPoint = "CreateJobObjectW", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
    private static partial IntPtr CreateJobObject(IntPtr lpJobAttributes, string? lpName);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool SetInformationJobObject(IntPtr hJob, int jobObjectInfoClass, ref JOBOBJECT_EXTENDED_LIMIT_INFORMATION lpJobObjectInfo, uint cbJobObjectInfoLength);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool AssignProcessToJobObject(IntPtr hJob, IntPtr hProcess);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool CloseHandle(IntPtr hObject);
}
