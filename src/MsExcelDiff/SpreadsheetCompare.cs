public static partial class SpreadsheetCompare
{
    static readonly string[] programFolders =
    [
        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
    ];

    static readonly string[] searchRelativePaths =
    [
        // Office 16 (Microsoft 365 / Office 2016+) - most common
        @"Microsoft Office\root\Office16\DCF\SPREADSHEETCOMPARE.EXE",
        // Click-to-Run installs place the exe inside a virtual filesystem (vfs) directory
        @"Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE",
        @"Microsoft Office\root\vfs\ProgramFilesX64\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE",
        // Office 15 (Office 2013)
        @"Microsoft Office\root\Office15\DCF\SPREADSHEETCOMPARE.EXE",
        @"Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office15\DCF\SPREADSHEETCOMPARE.EXE",
        @"Microsoft Office\root\vfs\ProgramFilesX64\Microsoft Office\Office15\DCF\SPREADSHEETCOMPARE.EXE"
    ];

    public static string? FindExecutable(string? settingsPath = null)
    {
        if (settingsPath != null &&
            File.Exists(settingsPath))
        {
            return settingsPath;
        }

        foreach (var folder in programFolders)
        {
            foreach (var relative in searchRelativePaths)
            {
                var path = Path.Combine(folder, relative);
                if (File.Exists(path))
                {
                    return path;
                }
            }
        }

        return null;
    }

    static string? FindAppVlp()
    {
        foreach (var folder in programFolders)
        {
            var path = Path.Combine(folder, @"Microsoft Office\root\Client\AppVLP.exe");
            if (File.Exists(path))
            {
                return path;
            }
        }

        return null;
    }

    public static async Task Launch(string path1, string path2, string? exePath = null)
    {
        var exe = FindExecutable(exePath);
        if (exe == null)
        {
            throw new(
                """
                Spreadsheet Compare (SPREADSHEETCOMPARE.EXE) was not found.
                It is included with Office Professional Plus / Microsoft 365 Apps for Enterprise.
                If installed in a custom location, use the 'set-path' command to configure the path.
                """);
        }

        // SPREADSHEETCOMPARE.EXE takes a single argument: a path to a file
        // containing the two workbook paths (one per line)
        var tempFile = TempFiles.Create($"{path1}{Environment.NewLine}{path2}");

        var job = JobObject.Create();

        try
        {
            using var process = await LaunchProcess(exe, tempFile);

            JobObject.AssignProcess(job, process.Handle);
            await MaximizeWindow(process);
            await process.WaitForExitAsync();
        }
        catch when (TempFiles.TryDelete(tempFile))
        {
            // unreachable: TryDeleteTempFile always returns false
            throw;
        }
        finally
        {
            JobObject.Close(job);
        }
    }

    static async Task<Process> LaunchProcess(string exe, string tempFile)
    {
        // Click-to-Run Office installs require launching via AppVLP.exe (the App-V
        // virtualization layer). SPREADSHEETCOMPARE.EXE crashes if launched directly.
        var appVlp = FindAppVlp();

        if (appVlp == null)
        {
            return LaunchDirect(exe, tempFile);
        }

        return await LaunchViaAppVlp(appVlp, exe, tempFile);
    }

    static Process LaunchDirect(string exe, string tempFile) =>
        Process.Start(
            new ProcessStartInfo(exe, tempFile)
            {
                UseShellExecute = true
            })
        ?? throw new("Failed to start Spreadsheet Compare process");

    static readonly string lockFilePath = Path.Combine(TempFiles.TempDirectory, ".lock");

    static async Task<Process> LaunchViaAppVlp(string appVlp, string exe, string tempFile)
    {
        // Serialize the snapshot-launch-identify sequence across concurrent
        // diffexcel instances. Without this, concurrent instances snapshot the
        // same PID set, race to claim the same SPREADSHEETCOMPARE process, and
        // leave others orphaned (not in any job object, so they survive when
        // diffexcel is killed).
        // Uses a file lock instead of a Mutex because file locks are not
        // thread-affine, allowing async code within the critical section.
        using (await AcquireFileLock())
        {
            var existingPids = GetSpreadsheetComparePids();

            using var launcher = Process.Start(
                                     new ProcessStartInfo(appVlp, $"\"{exe}\" {tempFile}")
                                     {
                                         UseShellExecute = false
                                     })
                                 ?? throw new("Failed to start AppVLP process");

            // AppVLP.exe is a launcher that exits after starting the real process.
            // Find the actual SPREADSHEETCOMPARE process and wait on it.
            await launcher.WaitForExitAsync();

            return await WaitForProcess(existingPids)
                   ?? throw new("Spreadsheet Compare did not start. Ensure the application is installed correctly.");
        }
    }

    static async Task<FileStream> AcquireFileLock()
    {
        for (var i = 0; i < 300; i++)
        {
            try
            {
                return new(lockFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                await Task.Delay(100);
            }
        }

        throw new IOException($"Failed to acquire lock file: {lockFilePath}");
    }

    static HashSet<int> GetSpreadsheetComparePids() =>
        GetProcessPids("SPREADSHEETCOMPARE");

    internal static HashSet<int> GetProcessPids(string processName)
    {
        var processes = Process.GetProcessesByName(processName);
        var pids = processes.Select(_ => _.Id).ToHashSet();
        foreach (var process in processes)
        {
            process.Dispose();
        }

        return pids;
    }

    static Task<Process?> WaitForProcess(HashSet<int> existingPids) =>
        WaitForProcess("SPREADSHEETCOMPARE", existingPids);

    internal static async Task<Process?> WaitForProcess(string processName, HashSet<int> existingPids, int maxAttempts = 100)
    {
        for (var i = 0; i < maxAttempts; i++)
        {
            var processes = Process.GetProcessesByName(processName);
            Process? result = null;
            foreach (var process in processes)
            {
                if (result == null && !existingPids.Contains(process.Id))
                {
                    result = process;
                }
                else
                {
                    process.Dispose();
                }
            }

            if (result != null)
            {
                return result;
            }

            await Task.Delay(100);
        }

        return null;
    }

    static async Task MaximizeWindow(Process process)
    {
        // Wait for the main window to appear
        for (var i = 0; i < 100; i++)
        {
            process.Refresh();
            if (process.MainWindowHandle != IntPtr.Zero)
            {
                // SW_MAXIMIZE = 3
                ShowWindow(process.MainWindowHandle, 3);
                SetForegroundWindow(process.MainWindowHandle);

                // Wait briefly for the window to finish layout after maximize
                await Task.Delay(500);
                CenterVerticalSplit(process.MainWindowHandle);
                return;
            }

            await Task.Delay(100);
        }
    }

    static void CenterVerticalSplit(IntPtr mainWindow)
    {
        // Collect all child windows with their parent, class name, and rect
        var children = new List<(IntPtr Handle, IntPtr Parent, string ClassName, RECT Rect)>();
        EnumChildWindows(mainWindow, (hwnd, _) =>
        {
            GetWindowRect(hwnd, out var rect);
            var className = GetWindowClassName(hwnd);
            children.Add((hwnd, GetParent(hwnd), className, rect));
            return true;
        }, IntPtr.Zero);

        // Log child window hierarchy for diagnostics
        Log.Information("CenterVerticalSplit: found {Count} child windows", children.Count);
        foreach (var child in children)
        {
            var w = child.Rect.Right - child.Rect.Left;
            var h = child.Rect.Bottom - child.Rect.Top;
            Log.Information(
                "  hwnd={Handle} parent={Parent} class={ClassName} pos=({Left},{Top}) size={Width}x{Height}",
                child.Handle, child.Parent, child.ClassName,
                child.Rect.Left, child.Rect.Top, w, h);
        }

        // Find the vertical splitter: look for pairs of side-by-side siblings
        // with similar height that together span most of their parent's width.
        // Pick the pair with the largest combined area.
        var bestArea = 0;
        var bestLeftRect = default(RECT);
        var bestRightRect = default(RECT);
        var bestParent = IntPtr.Zero;

        foreach (var group in children.GroupBy(c => c.Parent))
        {
            var siblings = group.ToList();

            for (var i = 0; i < siblings.Count; i++)
            {
                for (var j = i + 1; j < siblings.Count; j++)
                {
                    var a = siblings[i];
                    var b = siblings[j];
                    var heightA = a.Rect.Bottom - a.Rect.Top;
                    var heightB = b.Rect.Bottom - b.Rect.Top;

                    var widthA = a.Rect.Right - a.Rect.Left;
                    var widthB = b.Rect.Right - b.Rect.Left;

                    if (heightA < 100 || heightB < 100 ||
                        widthA <= 0 || widthB <= 0)
                    {
                        continue;
                    }

                    if (Math.Abs(heightA - heightB) > 20 ||
                        Math.Abs(a.Rect.Top - b.Rect.Top) > 20)
                    {
                        continue;
                    }

                    GetClientRect(group.Key, out var parentClient);
                    var totalSpan = Math.Max(a.Rect.Right, b.Rect.Right) - Math.Min(a.Rect.Left, b.Rect.Left);
                    if (totalSpan < parentClient.Right * 0.8)
                    {
                        continue;
                    }

                    var area = (a.Rect.Right - a.Rect.Left) * heightA +
                               (b.Rect.Right - b.Rect.Left) * heightB;
                    if (area <= bestArea)
                    {
                        continue;
                    }

                    bestArea = area;
                    bestParent = group.Key;
                    if (a.Rect.Left <= b.Rect.Left)
                    {
                        bestLeftRect = a.Rect;
                        bestRightRect = b.Rect;
                    }
                    else
                    {
                        bestLeftRect = b.Rect;
                        bestRightRect = a.Rect;
                    }
                }
            }
        }

        if (bestArea == 0)
        {
            Log.Information("CenterVerticalSplit: no matching split panel pair found");
            return;
        }

        // The splitter bar sits in the gap between the two panels.
        // Convert splitter screen position to parent client coordinates and
        // send mouse messages directly to the parent (SplitContainer) window.
        var splitterScreenX = (bestLeftRect.Right + bestRightRect.Left) / 2;
        var splitterScreenY = (bestLeftRect.Top + bestLeftRect.Bottom) / 2;

        var splitterPoint = new POINT { X = splitterScreenX, Y = splitterScreenY };
        ScreenToClient(bestParent, ref splitterPoint);

        GetClientRect(bestParent, out var client);
        var targetClientX = client.Right / 2;

        Log.Information(
            "CenterVerticalSplit: sending drag from client ({FromX},{FromY}) to ({ToX},{ToY})",
            splitterPoint.X, splitterPoint.Y, targetClientX, splitterPoint.Y);

        var downLParam = MakeLParam(splitterPoint.X, splitterPoint.Y);
        var moveLParam = MakeLParam(targetClientX, splitterPoint.Y);

        // WM_LBUTTONDOWN = 0x0201, WM_MOUSEMOVE = 0x0200, WM_LBUTTONUP = 0x0202
        // MK_LBUTTON = 0x0001
        SendMessage(bestParent, 0x0201, (IntPtr)0x0001, downLParam);
        SendMessage(bestParent, 0x0200, (IntPtr)0x0001, moveLParam);
        SendMessage(bestParent, 0x0202, IntPtr.Zero, moveLParam);
    }

    static IntPtr MakeLParam(int x, int y) =>
        (IntPtr)((y << 16) | (x & 0xFFFF));

    delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    struct RECT
    {
        public int Left, Top, Right, Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct POINT
    {
        public int X, Y;
    }

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool SetForegroundWindow(IntPtr hWnd);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool GetClientRect(IntPtr hWnd, out RECT lpRect);

    [LibraryImport("user32.dll")]
    private static partial IntPtr GetParent(IntPtr hWnd);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);

    [LibraryImport("user32.dll", EntryPoint = "SendMessageW")]
    private static partial IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    static string GetWindowClassName(IntPtr hWnd)
    {
        var buffer = new System.Text.StringBuilder(256);
        GetClassName(hWnd, buffer, buffer.Capacity);
        return buffer.ToString();
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

}
