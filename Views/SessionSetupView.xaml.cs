using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace AiInterviewAssistant;

public partial class SessionSetupView : UserControl
{
    private Window? _hostWindow;

    public SessionSetupView()
    {
        InitializeComponent();
        Loaded += OnLoaded;
        Unloaded += OnUnloaded;
    }

    public event RoutedEventHandler? FullSessionRequested;
    public event RoutedEventHandler? BuyCreditsRequested;
    public event RoutedEventHandler? FreeSessionRequested;
    public event RoutedEventHandler? CloseRequested;
    public event RoutedEventHandler? MinimizeRequested;
    public event Action<StartupWindowSlot>? WindowSlotRequested;
    public event RoutedEventHandler? PastSessionsRequested;
    public event RoutedEventHandler? DashboardRequested;
    public event RoutedEventHandler? LogoutRequested;

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        _hostWindow = Window.GetWindow(this);
        if (_hostWindow != null)
            _hostWindow.PreviewMouseLeftButtonDown += OnWindowPreviewMouseLeftButtonDown;

        MoreOptionsPopup.PlacementTarget = SharedHeader.MoreMenuAnchorElement;
        MoveOptionsPopup.PlacementTarget = SharedHeader.MoveMenuAnchorElement;

        CreditsState.Current.PropertyChanged += CreditsState_OnPropertyChanged;
        RefreshPaidSessionOptionFromCredits();
    }

    private void OnUnloaded(object sender, RoutedEventArgs e)
    {
        CreditsState.Current.PropertyChanged -= CreditsState_OnPropertyChanged;

        if (_hostWindow != null)
        {
            _hostWindow.PreviewMouseLeftButtonDown -= OnWindowPreviewMouseLeftButtonDown;
            _hostWindow = null;
        }
    }

    private void CreditsState_OnPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(CreditsState.CreditsText))
            RefreshPaidSessionOptionFromCredits();
    }

    private void RefreshPaidSessionOptionFromCredits()
    {
        var buy = CreditsState.Current.ShouldShowBuyCreditsInsteadOfFullSession();
        PaidSessionOptionTitle.Text = buy ? "Buy credits" : "Full session";
        PaidSessionOptionBorder.ToolTip = buy
            ? "You need interview credits for a full session. Opens the website pricing section in your browser."
            : "Creating a full session does not spend credits. When you activate the interview, 0.5 credit is used for the first block of time; the session can auto-extend near the end.";
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

    private void Header_MoreClicked(object? sender, RoutedEventArgs e)
    {
        MoveOptionsPopup.IsOpen = false;
        MoreOptionsPopup.IsOpen = !MoreOptionsPopup.IsOpen;
    }

    private void Header_MoveClicked(object? sender, RoutedEventArgs e)
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

    private void PaidSessionOption_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        if (CreditsState.Current.ShouldShowBuyCreditsInsteadOfFullSession())
        {
            BuyCreditsRequested?.Invoke(this, new RoutedEventArgs());
            return;
        }

        FullSessionRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void FreeSession_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        FreeSessionRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void PastSessionsTab_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        PastSessionsRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void Header_CloseClicked(object? sender, RoutedEventArgs e)
    {
        CloseRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void Header_MinimizeClicked(object? sender, RoutedEventArgs e)
    {
        MinimizeRequested?.Invoke(this, new RoutedEventArgs());
    }
}
