using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

[assembly: System.Reflection.AssemblyTitle("演示倒计时")]
[assembly: System.Reflection.AssemblyDescription("PowerPoint、WPS 演示和 WPS PDF 全屏倒计时")]
[assembly: System.Reflection.AssemblyVersion("2.0.0.1")]

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        NativeWindowApi.SetProcessDPIAware();
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        bool isFirstInstance;
        using (var instance = new System.Threading.Mutex(true, @"Local\WpsSlideCountdown", out isFirstInstance))
        {
            if (isFirstInstance)
                Application.Run(new CountdownOverlay());
            else
                MessageBox.Show("倒计时工具已打开。右键点击倒计时框可打开菜单。", "演示倒计时");
        }
    }
}

internal sealed class SettingsWindow : Form
{
    private readonly NumericUpDown secondsField;

    internal int Seconds { get { return Decimal.ToInt32(secondsField.Value); } }

    internal SettingsWindow(int seconds)
    {
        Text = "倒计时设置";
        StartPosition = FormStartPosition.CenterScreen;
        AutoScaleDimensions = new SizeF(96, 96);
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = new Font("Microsoft YaHei UI", 10f);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(306, 128);

        var durationLabel = new Label
        {
            Text = "倒计时时长（秒）",
            AutoSize = true,
            Location = new Point(18, 17)
        };
        secondsField = new NumericUpDown
        {
            Location = new Point(21, 45),
            Size = new Size(264, 30),
            Font = new Font("Segoe UI", 13f),
            Minimum = 1,
            Maximum = CountdownOverlay.MaximumSeconds,
            Value = seconds,
            AccessibleName = durationLabel.Text
        };
        var confirm = new Button
        {
            Text = "确定", DialogResult = DialogResult.OK,
            Location = new Point(121, 83), Size = new Size(78, 30)
        };
        var cancel = new Button
        {
            Text = "取消", DialogResult = DialogResult.Cancel,
            Location = new Point(207, 83), Size = new Size(78, 30)
        };

        Controls.AddRange(new Control[] { durationLabel, secondsField, confirm, cancel });
        AcceptButton = confirm;
        CancelButton = cancel;
        ActiveControl = secondsField;
        Shown += (sender, args) => secondsField.Select(0, secondsField.Text.Length);
    }
}

internal sealed class CountdownOverlay : Form
{
    private static readonly string SettingsFile = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "WpsSlideCountdown", "seconds.txt");
    private static readonly string PositionFile = Path.Combine(Path.GetDirectoryName(SettingsFile), "position.txt");
    private static readonly Color NormalBackground = Color.FromArgb(21, 33, 52);
    private static readonly Color ExpiredBackground = Color.FromArgb(163, 38, 46);
    private const int DefaultSeconds = 60;
    internal const int MaximumSeconds = 5999;
    private const string AboutText = "演示倒计时2.0 （支持ppt和pdf）\r\n开发：曾峻峰（53106995@qq.com）";

    private readonly Stopwatch elapsed = new Stopwatch();
    private readonly Timer refresh = new Timer { Interval = 150 };
    private readonly Timer dragTimer = new Timer { Interval = 16 };
    private readonly ContextMenuStrip menu = new ContextMenuStrip { ShowImageMargin = false };
    private IntPtr activeShow;
    private Point? savedPosition;
    private Point dragStartMouse;
    private Point dragStartWindow;
    private bool dragMoved;
    private int savedSeconds;
    private int runningSeconds;
    private string timeText = "00:00";

    internal CountdownOverlay()
    {
        Text = "演示倒计时";
        TopMost = true;
        ShowInTaskbar = false;
        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.Manual;
        AutoScaleDimensions = new SizeF(96, 96);
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(176, 64);
        BackColor = NormalBackground;
        Opacity = 0.95;
        DoubleBuffered = true;
        AccessibleName = Text;
        Cursor = Cursors.SizeAll;

        menu.Items.Add("设置", null, (sender, args) => OpenSettings());
        menu.Items.Add("关于", null, (sender, args) => MessageBox.Show(this, AboutText, "关于"));
        menu.Items.Add("退出", null, (sender, args) => Close());
        ContextMenuStrip = menu;

        savedSeconds = LoadDuration();
        savedPosition = LoadPosition();
        refresh.Tick += (sender, args) => RefreshCountdown();
        dragTimer.Tick += (sender, args) =>
        {
            if (!NativeWindowApi.IsPrimaryMouseButtonDown())
                EndDrag(Cursor.Position);
            else
                MoveDrag(Cursor.Position);
        };
        RefreshCountdown();
    }

    protected override void OnShown(EventArgs args)
    {
        base.OnShown(args);
        RefreshCountdown();
        refresh.Start();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            refresh.Dispose();
            dragTimer.Dispose();
            menu.Dispose();
        }
        base.Dispose(disposing);
    }

    private static int LoadDuration()
    {
        try
        {
            int seconds;
            if (int.TryParse(File.ReadAllText(SettingsFile), NumberStyles.None, CultureInfo.InvariantCulture, out seconds)
                && seconds >= 1 && seconds <= MaximumSeconds)
                return seconds;
        }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
        return DefaultSeconds;
    }

    private static Point? LoadPosition()
    {
        try
        {
            string[] parts = File.ReadAllText(PositionFile).Split(',');
            int x, y;
            if (parts.Length == 2
                && int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out x)
                && int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out y))
                return new Point(x, y);
        }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
        return null;
    }

    protected override void OnMouseDown(MouseEventArgs args)
    {
        base.OnMouseDown(args);
        if (args.Button != MouseButtons.Left)
            return;
        BeginDrag(PointToScreen(args.Location));
    }

    protected override void OnMouseUp(MouseEventArgs args)
    {
        base.OnMouseUp(args);
        if (args.Button == MouseButtons.Left && dragTimer.Enabled)
            EndDrag(Cursor.Position);
    }

    private void BeginDrag(Point pointer)
    {
        dragStartMouse = pointer;
        dragStartWindow = Location;
        dragMoved = false;
        dragTimer.Start();
    }

    private void MoveDrag(Point pointer)
    {
        int dx = pointer.X - dragStartMouse.X;
        int dy = pointer.Y - dragStartMouse.Y;
        Size threshold = SystemInformation.DragSize;
        if (!dragMoved && Math.Abs(dx) < threshold.Width / 2 && Math.Abs(dy) < threshold.Height / 2)
            return;

        dragMoved = true;
        NativeWindowApi.PlaceOverlay(Handle, dragStartWindow.X + dx, dragStartWindow.Y + dy);
    }

    private void EndDrag(Point pointer)
    {
        MoveDrag(pointer);
        dragTimer.Stop();
        if (!dragMoved)
            return;

        Point position = FitOnScreen(Location);
        savedPosition = position;
        NativeWindowApi.PlaceOverlay(Handle, position.X, position.Y);
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(PositionFile));
            File.WriteAllText(PositionFile, string.Format(CultureInfo.InvariantCulture, "{0},{1}", position.X, position.Y));
        }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    private Point FitOnScreen(Point location)
    {
        Rectangle bounds = Screen.FromRectangle(new Rectangle(location, Size)).Bounds;
        int x = Math.Max(bounds.Left, Math.Min(location.X, bounds.Right - Width));
        int y = Math.Max(bounds.Top, Math.Min(location.Y, bounds.Bottom - Height));
        return new Point(x, y);
    }

    private void OpenSettings()
    {
        using (var settings = new SettingsWindow(savedSeconds))
        {
            if (settings.ShowDialog(this) != DialogResult.OK)
                return;
            savedSeconds = settings.Seconds;
        }

        RefreshCountdown();
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsFile));
            string pendingFile = SettingsFile + ".tmp";
            File.WriteAllText(pendingFile, savedSeconds.ToString(CultureInfo.InvariantCulture));
            if (File.Exists(SettingsFile))
                File.Replace(pendingFile, SettingsFile, null);
            else
                File.Move(pendingFile, SettingsFile);
        }
        catch (IOException) { ShowSaveNotice(); }
        catch (UnauthorizedAccessException) { ShowSaveNotice(); }
    }

    private void ShowSaveNotice()
    {
        MessageBox.Show(this, "新时长已应用，本次退出后将恢复为此前保存的设置。", Text);
    }

    private void RefreshCountdown()
    {
        IntPtr found = NativeWindowApi.FindSlideshow(activeShow);
        if (found != activeShow)
        {
            activeShow = found;
            if (found == IntPtr.Zero)
                elapsed.Reset();
            else
            {
                runningSeconds = savedSeconds;
                elapsed.Restart();
            }
        }

        NativeWindowApi.Rect rect;
        if (activeShow != IntPtr.Zero && NativeWindowApi.GetWindowRect(activeShow, out rect))
        {
            int remaining = Math.Max(0, (int)Math.Ceiling(runningSeconds - elapsed.Elapsed.TotalSeconds));
            Display(Rectangle.FromLTRB(rect.Left, rect.Top, rect.Right, rect.Bottom), remaining);
        }
        else
            Display(Screen.PrimaryScreen.WorkingArea, savedSeconds);
    }

    protected override bool ShowWithoutActivation { get { return true; } }

    protected override CreateParams CreateParams
    {
        get
        {
            CreateParams parameters = base.CreateParams;
            parameters.ExStyle |= NativeWindowApi.OverlayStyles;
            return parameters;
        }
    }

    protected override void OnSizeChanged(EventArgs args)
    {
        base.OnSizeChanged(args);
        if (ClientSize.Width < 24 || ClientSize.Height < 24)
            return;
        int diameter = Math.Max(12, ClientSize.Height / 4);
        int right = ClientSize.Width - diameter;
        int bottom = ClientSize.Height - diameter;
        using (var path = new GraphicsPath())
        {
            path.AddArc(0, 0, diameter, diameter, 180, 90);
            path.AddArc(right, 0, diameter, diameter, 270, 90);
            path.AddArc(right, bottom, diameter, diameter, 0, 90);
            path.AddArc(0, bottom, diameter, diameter, 90, 90);
            path.CloseFigure();
            Region previous = Region;
            Region = new Region(path);
            if (previous != null)
                previous.Dispose();
        }
    }

    private void Display(Rectangle slideBounds, int remaining)
    {
        string time = string.Format(CultureInfo.InvariantCulture, "{0:00}:{1:00}", remaining / 60, remaining % 60);
        if (timeText != time || BackColor != (remaining == 0 ? ExpiredBackground : NormalBackground))
        {
            timeText = time;
            AccessibleDescription = time;
            BackColor = remaining == 0 ? ExpiredBackground : NormalBackground;
            Invalidate();
        }

        if (!Visible)
        {
            Location = GetOverlayLocation(slideBounds);
            return;
        }

        if (!menu.Visible && !dragTimer.Enabled)
        {
            Point position = GetOverlayLocation(slideBounds);
            NativeWindowApi.PlaceOverlay(Handle, position.X, position.Y);
        }
    }

    private Point GetOverlayLocation(Rectangle slideBounds)
    {
        if (savedPosition.HasValue)
            return FitOnScreen(savedPosition.Value);

        int left = slideBounds.Left + (slideBounds.Width - Width) / 2;
        int top = slideBounds.Top + Math.Max(10, Height / 4);
        return new Point(left, top);
    }

    protected override void OnPaint(PaintEventArgs args)
    {
        base.OnPaint(args);
        args.Graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
        // Keep the five-character time within the available pixels at any DPI.
        float fontSize = Math.Max(1f, Math.Min(ClientSize.Height * 0.60f, ClientSize.Width * 0.25f));
        using (var font = new Font("Consolas", fontSize, FontStyle.Bold, GraphicsUnit.Pixel))
        using (var alignment = new StringFormat(StringFormat.GenericTypographic)
        {
            Alignment = StringAlignment.Center,
            LineAlignment = StringAlignment.Center,
            FormatFlags = StringFormatFlags.NoWrap
        })
        {
            args.Graphics.DrawString(timeText, font, Brushes.White, ClientRectangle, alignment);
        }
    }
}

internal static class NativeWindowApi
{
    private const int ToolWindow = 0x00000080;
    private const int NoActivate = 0x08000000;
    internal const int OverlayStyles = ToolWindow | NoActivate;
    private static readonly IntPtr Topmost = new IntPtr(-1);
    private const uint NoSize = 0x0001;
    private const uint PositionWithoutActivation = 0x0010;
    private const uint NoOwnerZOrder = 0x0200;

    internal static IntPtr FindSlideshow(IntPtr previous)
    {
        IntPtr foreground = GetForegroundWindow();
        if (foreground != previous && IsSlideshow(foreground))
            return foreground;
        if (IsSlideshow(previous))
        {
            // The editor or presenter view can have focus on a second display.
            uint foregroundOwner, previousOwner;
            GetWindowThreadProcessId(foreground, out foregroundOwner);
            GetWindowThreadProcessId(previous, out previousOwner);
            if (foregroundOwner == 0 || foregroundOwner == previousOwner)
                return previous;

            IntPtr foregroundShow = FindSlideshowForProcess(foregroundOwner);
            return foregroundShow != IntPtr.Zero ? foregroundShow : previous;
        }

        return FindSlideshowForProcess(0);
    }

    private static IntPtr FindSlideshowForProcess(uint processId)
    {
        IntPtr candidate = IntPtr.Zero;
        EnumWindows((window, unused) =>
        {
            if (processId != 0)
            {
                uint owner;
                GetWindowThreadProcessId(window, out owner);
                if (owner != processId)
                    return true;
            }
            if (!IsSlideshow(window))
                return true;
            candidate = window;
            return false;
        }, IntPtr.Zero);
        return candidate;
    }

    private static bool IsSlideshow(IntPtr window)
    {
        if (window == IntPtr.Zero || !IsWindowVisible(window))
            return false;

        var windowText = new StringBuilder(512);
        GetClassName(window, windowText, windowText.Capacity);
        bool powerpointWindow = string.Equals(windowText.ToString(), "screenClass", StringComparison.OrdinalIgnoreCase);
        bool pdfWindow = false;
        if (!powerpointWindow)
        {
            if (GetWindowText(window, windowText, windowText.Capacity) == 0)
                return false;

            string title = windowText.ToString();
            if (!IsWpsSlideshow(title))
            {
                pdfWindow = IsFullscreenPdf(window, title);
                if (!pdfWindow)
                    return false;
            }
        }

        uint processId;
        GetWindowThreadProcessId(window, out processId);
        try
        {
            using (Process owner = Process.GetProcessById((int)processId))
            {
                if (powerpointWindow)
                    return string.Equals(owner.ProcessName, "POWERPNT", StringComparison.OrdinalIgnoreCase);
                return string.Equals(owner.ProcessName, "wpp", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(owner.ProcessName, "wps", StringComparison.OrdinalIgnoreCase)
                    || (pdfWindow && string.Equals(owner.ProcessName, "wpspdf", StringComparison.OrdinalIgnoreCase));
            }
        }
        catch (ArgumentException) { return false; }
        catch (InvalidOperationException) { return false; }
    }

    private static bool IsWpsSlideshow(string title)
    {
        return title.StartsWith("WPS Presentation Slide Show", StringComparison.OrdinalIgnoreCase)
            || title.StartsWith("WPS 演示幻灯片放映", StringComparison.OrdinalIgnoreCase)
            || title.StartsWith("WPS 演示 幻灯片放映", StringComparison.OrdinalIgnoreCase)
            || title.StartsWith("WPS 幻灯片放映", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsFullscreenPdf(IntPtr window, string title)
    {
        bool pdfTitle = title.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)
            || title.IndexOf(".pdf - ", StringComparison.OrdinalIgnoreCase) >= 0
            || title.IndexOf(".pdf [", StringComparison.OrdinalIgnoreCase) >= 0
            || title.IndexOf(".pdf（", StringComparison.OrdinalIgnoreCase) >= 0;
        if (!pdfTitle)
            return false;

        // WPS reuses the document window and switches it to a borderless popup for full screen.
        const int windowStyle = -16;
        const uint popup = 0x80000000;
        Rect windowRect;
        if ((GetWindowLong(window, windowStyle) & popup) == 0 || !GetWindowRect(window, out windowRect))
            return false;

        Rectangle display = Screen.FromHandle(window).Bounds;
        const int pixelTolerance = 2;
        return Math.Abs(windowRect.Left - display.Left) <= pixelTolerance
            && Math.Abs(windowRect.Top - display.Top) <= pixelTolerance
            && Math.Abs(windowRect.Right - display.Right) <= pixelTolerance
            && Math.Abs(windowRect.Bottom - display.Bottom) <= pixelTolerance;
    }

    internal static void PlaceOverlay(IntPtr handle, int left, int top)
    {
        SetWindowPos(handle, Topmost, left, top, 0, 0, NoSize | PositionWithoutActivation | NoOwnerZOrder);
    }

    internal static bool IsPrimaryMouseButtonDown()
    {
        const int leftButton = 0x01;
        const int rightButton = 0x02;
        int button = SystemInformation.MouseButtonsSwapped ? rightButton : leftButton;
        return (GetAsyncKeyState(button) & 0x8000) != 0;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct Rect { internal int Left, Top, Right, Bottom; }

    private delegate bool EnumerateWindow(IntPtr window, IntPtr parameter);

    [DllImport("user32.dll")] internal static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] internal static extern bool GetWindowRect(IntPtr window, out Rect rect);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumerateWindow callback, IntPtr parameter);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int virtualKey);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr window, StringBuilder title, int capacity);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr window, StringBuilder className, int capacity);
    [DllImport("user32.dll", EntryPoint = "GetWindowLongW")]
    private static extern uint GetWindowLong(IntPtr window, int index);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(IntPtr window, IntPtr order, int x, int y, int width, int height, uint flags);
}
