using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace AiInterviewAssistant;

public partial class PastSessionsView : UserControl
{
    private Window? _hostWindow;
    private HttpClient? _apiClient;

    public PastSessionsView()
    {
        InitializeComponent();
        Loaded += OnLoaded;
        Unloaded += OnUnloaded;
    }

    public event RoutedEventHandler? CreateRequested;
    public event RoutedEventHandler? ViewAllRequested;
    public event RoutedEventHandler? CloseRequested;
    public event RoutedEventHandler? MinimizeRequested;
    public event Action<StartupWindowSlot>? WindowSlotRequested;
    public event RoutedEventHandler? DashboardRequested;
    public event RoutedEventHandler? LogoutRequested;
    public event Action<Guid, bool, string?>? ActivateNotActivatedRequested;

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        _hostWindow = Window.GetWindow(this);
        if (_hostWindow != null)
            _hostWindow.PreviewMouseLeftButtonDown += OnWindowPreviewMouseLeftButtonDown;

        MoreOptionsPopup.PlacementTarget = SharedHeader.MoreMenuAnchorElement;
        MoveOptionsPopup.PlacementTarget = SharedHeader.MoveMenuAnchorElement;
        _ = ReloadSessionsAsync();
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

    private void Header_MoreClicked(object? sender, RoutedEventArgs e)
    {
        MoveOptionsPopup.IsOpen = false;
        MoreOptionsPopup.IsOpen = !MoreOptionsPopup.IsOpen;
    }

    private void Header_MoveClicked(object? sender, RoutedEventArgs e)
    {
        MoreOptionsPopup.IsOpen = false;
        MoveOptionsPopup.IsOpen = !MoveOptionsPopup.IsOpen;
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

    private void CreateTab_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        CreateRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void ViewAll_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        ViewAllRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void SessionRow_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        if (sender is not FrameworkElement fe) return;
        if (fe.DataContext is not CallSessionListItem item) return;

        // Only Not Activated sessions can be started from here.
        if (!string.Equals(item.EndsIn?.Trim(), "Not Activated", StringComparison.OrdinalIgnoreCase))
            return;

        ActivateNotActivatedRequested?.Invoke(item.Id, item.IsFreeSession, item.Language);
    }

    public void ConfigureApiClient(HttpClient apiClient)
    {
        _apiClient = apiClient;
    }

    private const int MaxPastSessionsToShow = 10;

    public async Task ReloadSessionsAsync()
    {
        try
        {
            SessionsStatusText.Text = "Loading sessions...";
            SessionsStatusText.Visibility = Visibility.Visible;
            SessionsListView.ItemsSource = null;

            if (_apiClient == null)
            {
                SessionsStatusText.Text = "API client is not configured.";
                SessionsStatusText.Visibility = Visibility.Visible;
                return;
            }

            var response = await _apiClient.GetAsync(
                $"callsessions?page=1&pageSize={MaxPastSessionsToShow}");
            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync();
                SessionsStatusText.Text = $"Failed to load sessions ({(int)response.StatusCode}).";
                SessionsStatusText.Visibility = Visibility.Visible;
                DesktopLogger.Warn($"Past sessions load failed status={(int)response.StatusCode} body={body}");
                return;
            }

            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            var payload = await response.Content.ReadFromJsonAsync<PagedResult<CallSessionListItem>>(options);
            var items = payload?.Items ?? new List<CallSessionListItem>();
            // Always cap to latest N by date (defensive if API returns extra rows).
            items = items
                .OrderByDescending(i => i.CreatedAt)
                .Take(MaxPastSessionsToShow)
                .ToList();

            foreach (var item in items)
            {
                if (string.IsNullOrWhiteSpace(item.Title))
                    item.Title = "(Untitled session)";
                if (string.IsNullOrWhiteSpace(item.Language))
                    item.Language = "English";
                if (string.IsNullOrWhiteSpace(item.EndsIn))
                    item.EndsIn = "Not Activated";
                item.CreatedAtDisplay = item.CreatedAt.ToLocalTime().ToString("MMM dd, yyyy");
            }

            SessionsListView.ItemsSource = items;
            SessionsStatusText.Text = items.Count == 0
                ? "No past sessions found."
                : $"{items.Count} session(s) loaded.";
            SessionsStatusText.Visibility = items.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }
        catch (Exception ex)
        {
            SessionsStatusText.Text = "Unable to load sessions.";
            SessionsStatusText.Visibility = Visibility.Visible;
            DesktopLogger.Error($"Past sessions load exception: {ex}");
        }
    }

    private void Header_CloseClicked(object? sender, RoutedEventArgs e)
    {
        CloseRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void Header_MinimizeClicked(object? sender, RoutedEventArgs e)
    {
        MinimizeRequested?.Invoke(this, new RoutedEventArgs());
    }

    private sealed class PagedResult<T>
    {
        public List<T> Items { get; set; } = new();
        public int TotalCount { get; set; }
    }

    private sealed class CallSessionListItem
    {
        public Guid Id { get; set; }
        public string Title { get; set; } = string.Empty;
        public string Language { get; set; } = "English";
        public string EndsIn { get; set; } = "Not Activated";
        public bool IsFreeSession { get; set; } = true;
        public DateTime CreatedAt { get; set; }
        public string CreatedAtDisplay { get; set; } = string.Empty;
    }
}
