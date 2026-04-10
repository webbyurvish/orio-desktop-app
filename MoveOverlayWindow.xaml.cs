using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

namespace AiInterviewAssistant;

public partial class MoveOverlayWindow : Window
{
    private readonly Rectangle _screenBoundsPx;
    private readonly double _dipPerPx;
    private bool _chordConsumed;
    private StartupWindowSlot? _highlightedSlot;

    public StartupWindowSlot? SelectedSlot { get; private set; }

    private MoveOverlayWindow(Rectangle screenBoundsPx, double dipPerPx)
    {
        InitializeComponent();
        _screenBoundsPx = screenBoundsPx;
        _dipPerPx = dipPerPx;

        Left = _screenBoundsPx.Left * _dipPerPx;
        Top = _screenBoundsPx.Top * _dipPerPx;
        Width = _screenBoundsPx.Width * _dipPerPx;
        Height = _screenBoundsPx.Height * _dipPerPx;
        WindowStartupLocation = WindowStartupLocation.Manual;
    }

    public static StartupWindowSlot? PickSlot(Window owner)
    {
        var handle = new WindowInteropHelper(owner).Handle;
        if (!TryGetMonitorBoundsPx(handle, out var boundsPx, out var dpi))
            boundsPx = GetVirtualScreenBoundsPx();

        var dipPerPx = 96.0 / Math.Max(96u, dpi);
        var win = new MoveOverlayWindow(boundsPx, dipPerPx)
        {
            Owner = owner,
            ShowActivated = true
        };

        owner.Activate();
        var result = win.ShowDialog();
        return result == true ? win.SelectedSlot : null;
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        try
        {
            BackgroundImage.Source = CaptureScreenToBitmapSource(_screenBoundsPx);
        }
        catch
        {
            // If capture fails (permissions/drivers), keep the dark veil only.
        }

        Activate();
        Focus();
        Keyboard.Focus(this);
        UpdateHighlight();
    }

    private static BitmapSource CaptureScreenToBitmapSource(Rectangle boundsPx)
    {
        using var bmp = new Bitmap(boundsPx.Width, boundsPx.Height, System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
        using (var g = Graphics.FromImage(bmp))
        {
            g.CopyFromScreen(boundsPx.Left, boundsPx.Top, 0, 0, boundsPx.Size, CopyPixelOperation.SourceCopy);
        }

        var hBitmap = bmp.GetHbitmap();
        try
        {
            var src = Imaging.CreateBitmapSourceFromHBitmap(
                hBitmap,
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
            src.Freeze();
            return src;
        }
        finally
        {
            DeleteObject(hBitmap);
        }
    }

    [DllImport("gdi32.dll")]
    private static extern bool DeleteObject(IntPtr hObject);

    private static Rectangle GetVirtualScreenBoundsPx()
    {
        int left = GetSystemMetrics(SM_XVIRTUALSCREEN);
        int top = GetSystemMetrics(SM_YVIRTUALSCREEN);
        int width = GetSystemMetrics(SM_CXVIRTUALSCREEN);
        int height = GetSystemMetrics(SM_CYVIRTUALSCREEN);
        return new Rectangle(left, top, width, height);
    }

    private static bool TryGetMonitorBoundsPx(IntPtr hwnd, out Rectangle boundsPx, out uint dpi)
    {
        boundsPx = default;
        dpi = 96;

        try
        {
            var hMon = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
            if (hMon == IntPtr.Zero) return false;

            var mi = new MONITORINFO();
            mi.cbSize = Marshal.SizeOf<MONITORINFO>();
            if (!GetMonitorInfo(hMon, ref mi)) return false;

            dpi = GetDpiForWindow(hwnd);
            boundsPx = Rectangle.FromLTRB(mi.rcMonitor.Left, mi.rcMonitor.Top, mi.rcMonitor.Right, mi.rcMonitor.Bottom);
            return boundsPx.Width > 0 && boundsPx.Height > 0;
        }
        catch
        {
            return false;
        }
    }

    private const int SM_XVIRTUALSCREEN = 76;
    private const int SM_YVIRTUALSCREEN = 77;
    private const int SM_CXVIRTUALSCREEN = 78;
    private const int SM_CYVIRTUALSCREEN = 79;

    private const uint MONITOR_DEFAULTTONEAREST = 2;

    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int nIndex);

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hwnd);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MONITORINFO
    {
        public int cbSize;
        public RECT rcMonitor;
        public RECT rcWork;
        public uint dwFlags;
    }

    private static bool IsCtrlDown() =>
        (Keyboard.Modifiers & ModifierKeys.Control) != 0;

    private static bool IsKeyActive(KeyEventArgs e, Key k) =>
        e.Key == k || Keyboard.IsKeyDown(k);

    private static bool IsUp(KeyEventArgs e) =>
        IsKeyActive(e, Key.Up) || IsKeyActive(e, Key.NumPad8);

    private static bool IsDown(KeyEventArgs e) =>
        IsKeyActive(e, Key.Down) || IsKeyActive(e, Key.NumPad2);

    private static bool IsLeft(KeyEventArgs e) =>
        IsKeyActive(e, Key.Left) || IsKeyActive(e, Key.NumPad4);

    private static bool IsRight(KeyEventArgs e) =>
        IsKeyActive(e, Key.Right) || IsKeyActive(e, Key.NumPad6);

    private static bool ArrowLeftHeld() =>
        Keyboard.IsKeyDown(Key.Left) || Keyboard.IsKeyDown(Key.NumPad4);

    private static bool ArrowRightHeld() =>
        Keyboard.IsKeyDown(Key.Right) || Keyboard.IsKeyDown(Key.NumPad6);

    private static bool ArrowUpHeld() =>
        Keyboard.IsKeyDown(Key.Up) || Keyboard.IsKeyDown(Key.NumPad8);

    private static bool ArrowDownHeld() =>
        Keyboard.IsKeyDown(Key.Down) || Keyboard.IsKeyDown(Key.NumPad2);

    private void Select(StartupWindowSlot slot)
    {
        _chordConsumed = true;
        SelectedSlot = slot;
        DialogResult = true;
        Close();
    }

    private StartupWindowSlot? ComputeSlotFromCurrentKeys(KeyEventArgs? e = null)
    {
        if (!IsCtrlDown()) return null;

        // Prefer corners when both arrows are down.
        bool up = e != null ? IsUp(e) : ArrowUpHeld();
        bool down = e != null ? IsDown(e) : ArrowDownHeld();
        bool left = e != null ? IsLeft(e) : ArrowLeftHeld();
        bool right = e != null ? IsRight(e) : ArrowRightHeld();

        if (up && left) return StartupWindowSlot.TopLeft;
        if (up && right) return StartupWindowSlot.TopRight;
        if (down && left) return StartupWindowSlot.BottomLeft;
        if (down && right) return StartupWindowSlot.BottomRight;

        if (up) return StartupWindowSlot.TopCenter;
        if (down) return StartupWindowSlot.BottomCenter;

        return null;
    }

    private Border? SlotBorderFor(StartupWindowSlot slot) => slot switch
    {
        StartupWindowSlot.TopLeft => SlotTopLeft,
        StartupWindowSlot.TopCenter => SlotTopCenter,
        StartupWindowSlot.TopRight => SlotTopRight,
        StartupWindowSlot.BottomLeft => SlotBottomLeft,
        StartupWindowSlot.BottomCenter => SlotBottomCenter,
        StartupWindowSlot.BottomRight => SlotBottomRight,
        _ => null
    };

    private void ClearHighlight()
    {
        if (_highlightedSlot == null) return;
        var b = SlotBorderFor(_highlightedSlot.Value);
        if (b != null)
        {
            b.BorderThickness = new Thickness(1);
            b.BorderBrush = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#40A78BFA")!;
            b.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#241A1F2E")!;
        }
        _highlightedSlot = null;
    }

    private void ApplyHighlight(StartupWindowSlot slot)
    {
        var b = SlotBorderFor(slot);
        if (b == null) return;

        b.BorderThickness = new Thickness(2);
        b.BorderBrush = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#A5B4FC")!;
        b.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#3A1E2436")!;
        _highlightedSlot = slot;
    }

    private void UpdateHighlight(KeyEventArgs? e = null)
    {
        var next = ComputeSlotFromCurrentKeys(e);
        if (next == _highlightedSlot) return;

        ClearHighlight();
        if (next != null)
            ApplyHighlight(next.Value);
    }

    private void Slot1_Click(object sender, MouseButtonEventArgs e) => Select(StartupWindowSlot.TopLeft);
    private void Slot2_Click(object sender, MouseButtonEventArgs e) => Select(StartupWindowSlot.TopCenter);
    private void Slot3_Click(object sender, MouseButtonEventArgs e) => Select(StartupWindowSlot.TopRight);
    private void Slot4_Click(object sender, MouseButtonEventArgs e) => Select(StartupWindowSlot.BottomLeft);
    private void Slot5_Click(object sender, MouseButtonEventArgs e) => Select(StartupWindowSlot.BottomCenter);
    private void Slot6_Click(object sender, MouseButtonEventArgs e) => Select(StartupWindowSlot.BottomRight);

    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        UpdateHighlight(e);

        if (e.Key == Key.Escape)
        {
            e.Handled = true;
            DialogResult = false;
            Close();
            return;
        }

        if (!IsCtrlDown()) return;

        // Diagonal corners: Ctrl + two arrows (either order once both are down).
        if (IsUp(e) && IsLeft(e))
        {
            e.Handled = true;
            Select(StartupWindowSlot.TopLeft);
            return;
        }

        if (IsUp(e) && IsRight(e))
        {
            e.Handled = true;
            Select(StartupWindowSlot.TopRight);
            return;
        }

        if (IsDown(e) && IsLeft(e))
        {
            e.Handled = true;
            Select(StartupWindowSlot.BottomLeft);
            return;
        }

        if (IsDown(e) && IsRight(e))
        {
            e.Handled = true;
            Select(StartupWindowSlot.BottomRight);
            return;
        }
    }

    private void Window_PreviewKeyUp(object sender, KeyEventArgs e)
    {
        UpdateHighlight(e);

        if (_chordConsumed) return;
        if (!IsCtrlDown()) return;

        switch (e.Key)
        {
            case Key.Up:
            case Key.NumPad8:
                e.Handled = true;
                if (ArrowLeftHeld()) Select(StartupWindowSlot.TopLeft);
                else if (ArrowRightHeld()) Select(StartupWindowSlot.TopRight);
                else Select(StartupWindowSlot.TopCenter);
                return;
            case Key.Down:
            case Key.NumPad2:
                e.Handled = true;
                if (ArrowLeftHeld()) Select(StartupWindowSlot.BottomLeft);
                else if (ArrowRightHeld()) Select(StartupWindowSlot.BottomRight);
                else Select(StartupWindowSlot.BottomCenter);
                return;
            case Key.Left:
            case Key.NumPad4:
                if (ArrowUpHeld() && !ArrowDownHeld())
                {
                    e.Handled = true;
                    Select(StartupWindowSlot.TopLeft);
                }
                else if (ArrowDownHeld() && !ArrowUpHeld())
                {
                    e.Handled = true;
                    Select(StartupWindowSlot.BottomLeft);
                }

                return;
            case Key.Right:
            case Key.NumPad6:
                if (ArrowUpHeld() && !ArrowDownHeld())
                {
                    e.Handled = true;
                    Select(StartupWindowSlot.TopRight);
                }
                else if (ArrowDownHeld() && !ArrowUpHeld())
                {
                    e.Handled = true;
                    Select(StartupWindowSlot.BottomRight);
                }

                return;
            case Key.LeftCtrl:
            case Key.RightCtrl:
                ClearHighlight();
                return;
        }
    }
}
