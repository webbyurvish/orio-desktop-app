using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace AiInterviewAssistant;

public enum StartupWindowSlot
{
    TopLeft,
    TopCenter,
    TopRight,
    BottomLeft,
    BottomCenter,
    BottomRight
}

public partial class StartupLoginView : UserControl
{
    private Window? _hostWindow;
    private bool _isLoginBusy;

    public StartupLoginView()
    {
        InitializeComponent();
        Loaded += OnLoaded;
        Unloaded += OnUnloaded;
    }

    public event RoutedEventHandler? LoginRequested;
    public event RoutedEventHandler? CloseRequested;
    public event RoutedEventHandler? MinimizeRequested;
    public event Action<StartupWindowSlot>? WindowSlotRequested;
    public event RoutedEventHandler? DashboardRequested;
    public event RoutedEventHandler? LogoutRequested;

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        _hostWindow = Window.GetWindow(this);
        if (_hostWindow != null)
            _hostWindow.PreviewMouseLeftButtonDown += OnWindowPreviewMouseLeftButtonDown;
        MoreOptionsPopup.PlacementTarget = SharedHeader.MoreMenuAnchorElement;
        MoveOptionsPopup.PlacementTarget = SharedHeader.MoveMenuAnchorElement;
    }

    private void OnUnloaded(object sender, RoutedEventArgs e)
    {
        if (_hostWindow != null)
        {
            _hostWindow.PreviewMouseLeftButtonDown -= OnWindowPreviewMouseLeftButtonDown;
            _hostWindow = null;
        }
    }

    private void OnWindowPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (!MoreOptionsPopup.IsOpen && !MoveOptionsPopup.IsOpen) return;
        if (e.OriginalSource is not DependencyObject src) return;
        if (IsWithin(src, SharedHeader.MoreMenuAnchorElement) || IsWithin(src, SharedHeader.MoveMenuAnchorElement)) return;
        if (MoreOptionsPopup.Child is DependencyObject morePopupRoot && IsWithin(src, morePopupRoot)) return;
        if (MoveOptionsPopup.Child is DependencyObject movePopupRoot && IsWithin(src, movePopupRoot)) return;
        MoreOptionsPopup.IsOpen = false;
        MoveOptionsPopup.IsOpen = false;
    }

    private static bool IsWithin(DependencyObject? child, DependencyObject? ancestor)
    {
        while (child != null)
        {
            if (ReferenceEquals(child, ancestor)) return true;
            child = VisualTreeHelper.GetParent(child);
        }
        return false;
    }

    private void SharedHeader_MoreClicked(object? sender, RoutedEventArgs e)
    {
        MoveOptionsPopup.IsOpen = false;
        MoreOptionsPopup.IsOpen = !MoreOptionsPopup.IsOpen;
    }

    private void SharedHeader_MoveClicked(object? sender, RoutedEventArgs e)
    {
        MoreOptionsPopup.IsOpen = false;
        MoveOptionsPopup.IsOpen = false;

        var owner = Window.GetWindow(this);
        if (owner == null) return;

        var picked = MoveOverlayWindow.PickSlot(owner);
        if (picked != null)
            WindowSlotRequested?.Invoke(picked.Value);
    }

    private void EmitWindowSlot(StartupWindowSlot slot, MouseButtonEventArgs e)
    {
        e.Handled = true;
        MoveOptionsPopup.IsOpen = false;
        WindowSlotRequested?.Invoke(slot);
    }

    private void MoveTopLeft_Click(object sender, MouseButtonEventArgs e) => EmitWindowSlot(StartupWindowSlot.TopLeft, e);
    private void MoveTopCenter_Click(object sender, MouseButtonEventArgs e) => EmitWindowSlot(StartupWindowSlot.TopCenter, e);
    private void MoveTopRight_Click(object sender, MouseButtonEventArgs e) => EmitWindowSlot(StartupWindowSlot.TopRight, e);
    private void MoveBottomLeft_Click(object sender, MouseButtonEventArgs e) => EmitWindowSlot(StartupWindowSlot.BottomLeft, e);
    private void MoveBottomCenter_Click(object sender, MouseButtonEventArgs e) => EmitWindowSlot(StartupWindowSlot.BottomCenter, e);
    private void MoveBottomRight_Click(object sender, MouseButtonEventArgs e) => EmitWindowSlot(StartupWindowSlot.BottomRight, e);

    private void MenuDashboard_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        MoreOptionsPopup.IsOpen = false;
        DashboardRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void MenuLogout_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        MoreOptionsPopup.IsOpen = false;
        LogoutRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void MenuNextScreen_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        MoreOptionsPopup.IsOpen = false;
    }

    private void MenuNextScreenArrow_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        MoreOptionsPopup.IsOpen = false;
    }

    private void MenuZoomPlus_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void MenuZoomMinus_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void MenuZoomReset_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private void LoginBorder_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        if (_isLoginBusy) return;
        LoginRequested?.Invoke(this, new RoutedEventArgs());
    }

    public void SetLoginBusy(bool isBusy, string? message = null)
    {
        _isLoginBusy = isBusy;
        LoginBorder.Opacity = isBusy ? 0.7 : 1.0;
        LoginBorder.Cursor = isBusy ? Cursors.Wait : Cursors.Hand;
        LoginButtonText.Text = isBusy ? "Connecting..." : "Login";
        if (message != null)
            LoginStatusText.Text = message;
    }

    public void SetLoginStatus(string? message)
    {
        LoginStatusText.Text = message ?? string.Empty;
    }

    private void CloseBorder_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        CloseRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void SharedHeader_MinimizeClicked(object? sender, RoutedEventArgs e)
    {
        MinimizeRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void SharedHeader_CloseClicked(object? sender, RoutedEventArgs e)
    {
        CloseRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void SharedHeader_HeaderDragRequested(object? sender, MouseButtonEventArgs e)
    {
        if (e.ButtonState != MouseButtonState.Pressed) return;
        if (Window.GetWindow(this) is Window w)
            w.DragMove();
    }
}
