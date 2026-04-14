using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace AiInterviewAssistant;

internal static class WindowPrivacy
{
    // Windows 10 2004+ (19041+) supports WDA_EXCLUDEFROMCAPTURE.
    // On older versions this call may fail; we treat it as best-effort.
    private const uint WDA_NONE = 0;
    private const uint WDA_EXCLUDEFROMCAPTURE = 0x11;

    private const int GWL_EXSTYLE = -20;
    private const int WS_EX_TOOLWINDOW = 0x00000080; // hides from Alt-Tab
    private const int WS_EX_APPWINDOW = 0x00040000;  // forces Alt-Tab

    public static void Apply(Window window)
    {
        if (window == null) return;

        IntPtr hwnd;
        try
        {
            hwnd = new WindowInteropHelper(window).Handle;
        }
        catch
        {
            return;
        }

        ApplyToHwnd(hwnd);
    }

    public static void ApplyToHwnd(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero) return;

        try
        {
            // 1) Exclude from most screen capture APIs (screenshots, Teams/Zoom/window capture).
            // If it fails, do nothing (no hard dependency).
            _ = SetWindowDisplayAffinity(hwnd, WDA_EXCLUDEFROMCAPTURE) ||
                SetWindowDisplayAffinity(hwnd, WDA_NONE);
        }
        catch
        {
            // ignore
        }

        try
        {
            // 2) Ensure it never appears in Alt-Tab / Win+Tab.
            // ShowInTaskbar=false alone usually does this, but ex-style makes it consistent.
            var exStyle = GetWindowLongPtr(hwnd, GWL_EXSTYLE);
            var exStyleInt = unchecked((int)exStyle.ToInt64());
            exStyleInt |= WS_EX_TOOLWINDOW;
            exStyleInt &= ~WS_EX_APPWINDOW;
            _ = SetWindowLongPtr(hwnd, GWL_EXSTYLE, new IntPtr(exStyleInt));
        }
        catch
        {
            // ignore
        }
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowDisplayAffinity(IntPtr hWnd, uint dwAffinity);

    // GetWindowLongPtr/SetWindowLongPtr need dual entry-points for x86/x64.
    [DllImport("user32.dll", EntryPoint = "GetWindowLong")]
    private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")]
    private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

    private static IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex) =>
        IntPtr.Size == 8 ? GetWindowLongPtr64(hWnd, nIndex) : new IntPtr(GetWindowLong32(hWnd, nIndex));

    [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
    private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
    private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    private static IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong) =>
        IntPtr.Size == 8 ? SetWindowLongPtr64(hWnd, nIndex, dwNewLong) : new IntPtr(SetWindowLong32(hWnd, nIndex, dwNewLong.ToInt32()));
}

