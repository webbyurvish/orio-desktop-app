using System;
using System.Collections.Generic;
using System.Diagnostics;
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

public partial class CreateSessionDetailsView : UserControl
{
    /// <summary>Last item in the resume list; choosing it opens the web CV upload page and does not submit as a resume ID.</summary>
    private static readonly Guid UploadResumeSentinelId = Guid.Parse("ffffffff-ffff-4fff-bfff-fffffffffffe");

    private Window? _hostWindow;
    private HttpClient? _apiClient;

    public CreateSessionDetailsView()
    {
        InitializeComponent();
        ResumeComboBox.SelectionChanged += ResumeComboBox_SelectionChanged;
        Loaded += OnLoaded;
        Unloaded += OnUnloaded;
    }

    public event RoutedEventHandler? BackRequested;
    public event RoutedEventHandler? NextRequested;
    public event RoutedEventHandler? PastSessionsRequested;
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
        _ = ReloadResumesAsync();
        UpdateNextButtonState();
    }

    private void OnUnloaded(object sender, RoutedEventArgs e)
    {
        if (_hostWindow != null)
        {
            _hostWindow.PreviewMouseLeftButtonDown -= OnWindowPreviewMouseLeftButtonDown;
            _hostWindow = null;
        }
    }

    private void ResumeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ResumeComboBox.SelectedItem is not ResumeOption ro || ro.Id != UploadResumeSentinelId)
        {
            UpdateNextButtonState();
            return;
        }

        OpenResumeUploadPageInBrowser();

        ResumeComboBox.SelectionChanged -= ResumeComboBox_SelectionChanged;
        try
        {
            if (ResumeComboBox.ItemsSource is IEnumerable<ResumeOption> list)
            {
                var firstReal = list.FirstOrDefault(x => x.Id != UploadResumeSentinelId);
                if (firstReal != null)
                    ResumeComboBox.SelectedItem = firstReal;
                else
                    ResumeComboBox.SelectedIndex = -1;
            }
        }
        finally
        {
            ResumeComboBox.SelectionChanged += ResumeComboBox_SelectionChanged;
        }

        UpdateNextButtonState();
    }

    private static string GetWebAppOrigin()
    {
        const string fallback = "http://localhost:5173";
        var authorizeUrl = App.Settings.DesktopAuth.AuthorizeUrl?.Trim();
        if (!string.IsNullOrWhiteSpace(authorizeUrl) && Uri.TryCreate(authorizeUrl, UriKind.Absolute, out var authUri))
            return $"{authUri.Scheme}://{authUri.Authority}";
        return fallback;
    }

    private static void OpenResumeUploadPageInBrowser()
    {
        var url = $"{GetWebAppOrigin().TrimEnd('/')}/dashboard/cvs";
        try
        {
            Process.Start(new ProcessStartInfo { FileName = url, UseShellExecute = true });
            DesktopLogger.Info($"Opened resume upload page: {url}");
        }
        catch (Exception ex)
        {
            DesktopLogger.Error($"Failed to open resume upload page: {ex}");
            MessageBox.Show("Unable to open the resume upload page in your browser.", "Upload resume", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private static void AppendUploadResumeOption(ICollection<ResumeOption> resumes)
    {
        resumes.Add(new ResumeOption
        {
            Id = UploadResumeSentinelId,
            DisplayName = "+ Upload a resume",
        });
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

    private void PastSessionsTab_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        PastSessionsRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void Back_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        BackRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void Next_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        if (!IsCompanyValid())
        {
            UpdateNextButtonState(showValidation: true);
            return;
        }
        NextRequested?.Invoke(this, new RoutedEventArgs());
    }

    public void ConfigureApiClient(HttpClient apiClient)
    {
        _apiClient = apiClient;
    }

    public async Task ReloadResumesAsync()
    {
        try
        {
            ResumeLoadStatusText.Text = "Loading resumes...";
            ResumeComboBox.ItemsSource = null;
            ResumeComboBox.IsEnabled = false;

            if (_apiClient == null)
            {
                var uploadOnly = new List<ResumeOption>();
                AppendUploadResumeOption(uploadOnly);
                ResumeComboBox.ItemsSource = uploadOnly;
                ResumeComboBox.IsEnabled = true;
                ResumeComboBox.SelectedIndex = -1;
                ResumeLoadStatusText.Text = "API client is not configured. Use \"+ Upload a resume\" on the web.";
                return;
            }

            var response = await _apiClient.GetAsync("resumes");
            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync();
                var uploadOnly = new List<ResumeOption>();
                AppendUploadResumeOption(uploadOnly);
                ResumeComboBox.ItemsSource = uploadOnly;
                ResumeComboBox.IsEnabled = true;
                ResumeComboBox.SelectedIndex = -1;
                ResumeLoadStatusText.Text = $"Failed to load resumes ({(int)response.StatusCode}). You can upload a resume on the web.";
                DesktopLogger.Warn($"CreateSessionDetailsView resumes load failed status={(int)response.StatusCode} body={body}");
                return;
            }

            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            var resumes = await response.Content.ReadFromJsonAsync<List<ResumeOption>>(options) ?? new List<ResumeOption>();
            foreach (var resume in resumes)
            {
                if (string.IsNullOrWhiteSpace(resume.DisplayName))
                    resume.DisplayName = !string.IsNullOrWhiteSpace(resume.Title) ? resume.Title! : (resume.FileName ?? "Untitled resume");
            }

            AppendUploadResumeOption(resumes);
            ResumeComboBox.ItemsSource = resumes;
            ResumeComboBox.IsEnabled = true;
            var realCount = resumes.Count - 1;
            ResumeComboBox.SelectedIndex = realCount > 0 ? 0 : -1;
            ResumeLoadStatusText.Text = realCount > 0
                ? $"{realCount} resume(s) found. Use \"+ Upload a resume\" to add another on the web."
                : "No resumes yet. Choose \"+ Upload a resume\" to open the web upload page.";
            UpdateNextButtonState();
        }
        catch (Exception ex)
        {
            var uploadOnly = new List<ResumeOption>();
            AppendUploadResumeOption(uploadOnly);
            ResumeComboBox.ItemsSource = uploadOnly;
            ResumeComboBox.IsEnabled = true;
            ResumeComboBox.SelectedIndex = -1;
            ResumeLoadStatusText.Text = "Unable to load resumes. You can still open the web upload page.";
            DesktopLogger.Error($"CreateSessionDetailsView ReloadResumesAsync error: {ex}");
            UpdateNextButtonState();
        }
    }

    public string Company => (CompanyTextBox.Text ?? string.Empty).Trim();
    public string JobDescription => (JobDescriptionTextBox.Text ?? string.Empty).Trim();

    private bool IsCompanyValid() =>
        !string.IsNullOrWhiteSpace(CompanyTextBox.Text);

    private void UpdateNextButtonState(bool showValidation = false)
    {
        var companyOk = IsCompanyValid();
        var enabled = companyOk;

        if (CompanyValidationText != null)
            CompanyValidationText.Visibility = (!companyOk && showValidation) ? Visibility.Visible : Visibility.Collapsed;

        if (CompanyFieldHostBorder != null)
        {
            CompanyFieldHostBorder.BorderBrush = (!companyOk && showValidation)
                ? (Brush)new BrushConverter().ConvertFromString("#F87171")!
                : (Brush)new BrushConverter().ConvertFromString("#3352525E")!;
            CompanyFieldHostBorder.BorderThickness = (!companyOk && showValidation) ? new Thickness(1.5) : new Thickness(1);
        }

        if (NextButtonBorder != null)
        {
            NextButtonBorder.Opacity = enabled ? 1.0 : 0.45;
            NextButtonBorder.Cursor = enabled ? Cursors.Hand : Cursors.Arrow;
            NextButtonBorder.IsHitTestVisible = enabled;
        }

        if (NextButtonText != null)
            NextButtonText.Text = SelectedResumeId == null ? "Continue without resume" : "Next";
    }

    private void CompanyTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        // As soon as the user starts typing, clear any “required” error state.
        UpdateNextButtonState(showValidation: false);
    }

    public void ResetForNewSession()
    {
        CompanyTextBox.Text = string.Empty;
        JobDescriptionTextBox.Text = string.Empty;

        // "Fresh start": don't implicitly carry over last resume selection.
        // The next ReloadResumesAsync may pick index 0; we reset again after reload in the host when needed.
        ResumeComboBox.SelectedIndex = -1;
        ResumeLoadStatusText.Text = string.Empty;
        UpdateNextButtonState(showValidation: false);
    }

    public Guid? SelectedResumeId
    {
        get
        {
            var value = ResumeComboBox.SelectedValue;
            if (value is Guid g && g != UploadResumeSentinelId)
                return g;
            if (value is string s && Guid.TryParse(s, out var parsed) && parsed != UploadResumeSentinelId)
                return parsed;
            return null;
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

    private sealed class ResumeOption
    {
        public Guid Id { get; set; }
        public string? Title { get; set; }
        public string? FileName { get; set; }
        public string DisplayName { get; set; } = string.Empty;
    }
}
