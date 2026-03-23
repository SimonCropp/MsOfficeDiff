static partial class WindowLayout
{
    /// <summary>
    /// Maximizes the window and centers all split containers.
    /// </summary>
    internal static async Task MaximizeAndCenterSplits(Process process)
    {
        for (var i = 0; i < 100; i++)
        {
            process.Refresh();
            if (process.MainWindowHandle != IntPtr.Zero)
            {
                // SW_MAXIMIZE = 3
                // ShowWindow is synchronous — WinForms processes WM_SIZE and
                // lays out child controls before it returns, so no delay needed.
                ShowWindow(process.MainWindowHandle, 3);
                await Task.Delay(500);
                CenterSplits(process.MainWindowHandle);
                return;
            }

            await Task.Delay(100);
        }
    }

    static void CenterSplits(IntPtr mainWindow)
    {
        var children = new List<(IntPtr Handle, IntPtr Parent, string ClassName, RECT Rect)>();
        EnumChildWindows(
            mainWindow,
            (hwnd, _) =>
            {
                GetWindowRect(hwnd, out var rect);
                var className = GetWindowClassName(hwnd);
                children.Add((hwnd, GetParent(hwnd), className, rect));
                return true;
            },
            IntPtr.Zero);

        Log.Debug("CenterSplits: found {Count} child windows", children.Count);
        foreach (var child in children)
        {
            var w = child.Rect.Right - child.Rect.Left;
            var h = child.Rect.Bottom - child.Rect.Top;
            Log.Debug(
                "  hwnd={Handle} parent={Parent} class={ClassName} pos=({Left},{Top}) size={Width}x{Height}",
                child.Handle, child.Parent, child.ClassName,
                child.Rect.Left, child.Rect.Top, w, h);
        }

        CenterSplit(children, SplitOrientation.Vertical);
        CenterSplit(children, SplitOrientation.Horizontal);
    }

    enum SplitOrientation
    {
        Vertical,
        Horizontal
    }

    /// <summary>
    /// Finds all WinForms SplitContainer pairs in the given orientation and centers each splitter.
    /// Identifies split panels by looking for sibling window pairs that:
    ///   - have matching dimensions on the shared axis (height for vertical, width for horizontal)
    ///   - together span most of their parent's extent
    ///   - have a gap between them (the splitter bar)
    /// Uses PostMessage to simulate a mouse drag on each splitter, which goes through the
    /// target app's message queue so SetCapture works correctly for the drag operation.
    /// </summary>
    static void CenterSplit(
        List<(IntPtr Handle, IntPtr Parent, string ClassName, RECT Rect)> children,
        SplitOrientation orientation)
    {
        var matches = new List<(RECT First, RECT Second, IntPtr Parent)>();

        foreach (var group in children.GroupBy(c => c.Parent))
        {
            var siblings = group.ToList();

            for (var i = 0; i < siblings.Count; i++)
            {
                for (var j = i + 1; j < siblings.Count; j++)
                {
                    var a = siblings[i];
                    var b = siblings[j];
                    var widthA = a.Rect.Right - a.Rect.Left;
                    var widthB = b.Rect.Right - b.Rect.Left;
                    var heightA = a.Rect.Bottom - a.Rect.Top;
                    var heightB = b.Rect.Bottom - b.Rect.Top;

                    if (widthA <= 0 || widthB <= 0 ||
                        heightA <= 0 || heightB <= 0)
                    {
                        continue;
                    }

                    GetClientRect(group.Key, out var parentClient);

                    bool isMatch;
                    if (orientation == SplitOrientation.Vertical)
                    {
                        // Side-by-side: same height/top, span parent width
                        isMatch = Math.Abs(heightA - heightB) <= 20 &&
                                  Math.Abs(a.Rect.Top - b.Rect.Top) <= 20 &&
                                  Math.Max(a.Rect.Right, b.Rect.Right) - Math.Min(a.Rect.Left, b.Rect.Left) >= parentClient.Right * 0.8;
                    }
                    else
                    {
                        // Stacked: same width/left, span parent height
                        isMatch = Math.Abs(widthA - widthB) <= 20 &&
                                  Math.Abs(a.Rect.Left - b.Rect.Left) <= 20 &&
                                  Math.Max(a.Rect.Bottom, b.Rect.Bottom) - Math.Min(a.Rect.Top, b.Rect.Top) >= parentClient.Bottom * 0.8;
                    }

                    if (!isMatch)
                    {
                        continue;
                    }

                    // Require a gap between the panels (the splitter bar).
                    // Adjacent windows without a gap (e.g. ribbon/content) are not splits.
                    int gap;
                    RECT first, second;
                    if (orientation == SplitOrientation.Vertical)
                    {
                        first = a.Rect.Left <= b.Rect.Left ? a.Rect : b.Rect;
                        second = a.Rect.Left <= b.Rect.Left ? b.Rect : a.Rect;
                        gap = second.Left - first.Right;
                    }
                    else
                    {
                        first = a.Rect.Top <= b.Rect.Top ? a.Rect : b.Rect;
                        second = a.Rect.Top <= b.Rect.Top ? b.Rect : a.Rect;
                        gap = second.Top - first.Bottom;
                    }

                    if (gap <= 0)
                    {
                        continue;
                    }

                    matches.Add((first, second, group.Key));
                }
            }
        }

        if (matches.Count == 0)
        {
            Log.Debug("CenterSplit({Orientation}): no matching pairs found", orientation);
            return;
        }

        foreach (var match in matches)
        {
            // Convert screen coordinates to client coordinates of the parent (SplitContainer)
            var fromScreen = new POINT();
            GetClientRect(match.Parent, out var client);

            if (orientation == SplitOrientation.Vertical)
            {
                fromScreen.X = (match.First.Right + match.Second.Left) / 2;
                fromScreen.Y = (match.First.Top + match.First.Bottom) / 2;
            }
            else
            {
                fromScreen.X = (match.First.Left + match.First.Right) / 2;
                fromScreen.Y = (match.First.Bottom + match.Second.Top) / 2;
            }

            ScreenToClient(match.Parent, ref fromScreen);

            var toClient = new POINT { X = fromScreen.X, Y = fromScreen.Y };
            if (orientation == SplitOrientation.Vertical)
            {
                toClient.X = client.Right / 2;
            }
            else
            {
                toClient.Y = client.Bottom / 2;
            }

            Log.Debug(
                "CenterSplit({Orientation}): PostMessage drag client ({FromX},{FromY}) to ({ToX},{ToY})",
                orientation, fromScreen.X, fromScreen.Y, toClient.X, toClient.Y);

            var downLParam = MakeLParam(fromScreen.X, fromScreen.Y);
            var moveLParam = MakeLParam(toClient.X, toClient.Y);

            // WM_LBUTTONDOWN=0x0201 WM_MOUSEMOVE=0x0200 WM_LBUTTONUP=0x0202 MK_LBUTTON=0x0001
            PostMessage(match.Parent, 0x0201, 0x0001, downLParam);
            Thread.Sleep(50);
            PostMessage(match.Parent, 0x0200, 0x0001, moveLParam);
            Thread.Sleep(50);
            PostMessage(match.Parent, 0x0202, IntPtr.Zero, moveLParam);
            Thread.Sleep(100);
        }
    }

    static IntPtr MakeLParam(int x, int y) =>
        (y << 16) | (x & 0xFFFF);

    static string GetWindowClassName(IntPtr hWnd)
    {
        var buffer = new StringBuilder(256);
        GetClassName(hWnd, buffer, buffer.Capacity);
        return buffer.ToString();
    }

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

    [LibraryImport("user32.dll", EntryPoint = "PostMessageW")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
}
