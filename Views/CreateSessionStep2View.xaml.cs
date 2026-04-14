using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace AiInterviewAssistant;

public partial class CreateSessionStep2View : UserControl
{
    private Window? _hostWindow;
    private bool _isFreeSession = true;
    private bool _languagesLoaded;

    private sealed record SpeechLocaleOption(string Locale, string Label);
    // BCP-47 locales for interviews: Azure Speech-to-Text + Azure OpenAI (keep in sync with
    // orio-web-app/src/constants/azureSpeechSttLocales.ts). Excludes variants we don’t use.
    private static readonly SpeechLocaleOption[] AzureSpeechSttLocales =
    [
        // Pinned at top
        new("en-IN", "English (India)"),
        new("hi-IN", "Hindi (India)"),

        // Commonly used global / regional
        new("en-US", "English (United States)"),
        new("en-GB", "English (United Kingdom)"),
        new("es-ES", "Spanish (Spain)"),
        new("es-MX", "Spanish (Mexico)"),
        new("fr-FR", "French (France)"),
        new("fr-CA", "French (Canada)"),
        new("pt-BR", "Portuguese (Brazil)"),
        new("pt-PT", "Portuguese (Portugal)"),
        new("de-DE", "German (Germany)"),
        new("it-IT", "Italian (Italy)"),
        new("nl-NL", "Dutch (Netherlands)"),
        new("pl-PL", "Polish (Poland)"),
        new("ru-RU", "Russian (Russia)"),
        new("tr-TR", "Turkish (Türkiye)"),
        new("uk-UA", "Ukrainian (Ukraine)"),

        // Asia (commonly used)
        new("ja-JP", "Japanese (Japan)"),
        new("ko-KR", "Korean (Korea)"),
        new("zh-CN", "Chinese (Mandarin, Simplified)"),
        new("zh-TW", "Chinese (Taiwanese Mandarin, Traditional)"),
        new("zh-HK", "Chinese (Cantonese, Traditional)"),
        new("id-ID", "Indonesian (Indonesia)"),
        new("ms-MY", "Malay (Malaysia)"),
        new("th-TH", "Thai (Thailand)"),
        new("vi-VN", "Vietnamese (Vietnam)"),

        // Middle East / Africa (commonly used)
        new("ar-SA", "Arabic (Saudi Arabia)"),
        new("ar-EG", "Arabic (Egypt)"),
        new("fa-IR", "Persian (Iran)"),
        new("he-IL", "Hebrew (Israel)"),
        new("sw-KE", "Kiswahili (Kenya)"),
        new("af-ZA", "Afrikaans (South Africa)"),
        new("zu-ZA", "Zulu (South Africa)"),

        // India (commonly used)
        new("bn-IN", "Bengali (India)"),
        new("gu-IN", "Gujarati (India)"),
        new("kn-IN", "Kannada (India)"),
        new("ml-IN", "Malayalam (India)"),
        new("mr-IN", "Marathi (India)"),
        new("pa-IN", "Punjabi (India)"),
        new("ta-IN", "Tamil (India)"),
        new("te-IN", "Telugu (India)"),
        new("ur-IN", "Urdu (India)")
    ];

    private void EnsureLanguageOptionsLoaded()
    {
        if (_languagesLoaded) return;
        _languagesLoaded = true;

        LanguageComboBox.Items.Clear();
        foreach (var opt in AzureSpeechSttLocales)
        {
            LanguageComboBox.Items.Add(new ComboBoxItem
            {
                Content = opt.Label,
                Tag = opt.Locale
            });
        }
        LanguageComboBox.SelectedIndex = 0; // default en-IN label
    }

    public CreateSessionStep2View()
    {
        InitializeComponent();
        Loaded += OnLoaded;
        Unloaded += OnUnloaded;
    }

    public event RoutedEventHandler? BackRequested;
    public event RoutedEventHandler? CreateSessionRequested;
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
        EnsureLanguageOptionsLoaded();
        RefreshExtraContextPlaceholder();
    }

    private void ExtraContextTextBox_TextChanged(object sender, TextChangedEventArgs e) =>
        RefreshExtraContextPlaceholder();

    private void RefreshExtraContextPlaceholder()
    {
        var empty = string.IsNullOrWhiteSpace(ExtraContextTextBox.Text);
        ExtraContextPlaceholder.Visibility = empty ? Visibility.Visible : Visibility.Collapsed;
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

    private void CreateSession_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        CreateSessionRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void Header_CloseClicked(object? sender, RoutedEventArgs e)
    {
        CloseRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void Header_MinimizeClicked(object? sender, RoutedEventArgs e)
    {
        MinimizeRequested?.Invoke(this, new RoutedEventArgs());
    }

    public string SelectedLanguage
    {
        get
        {
            if (LanguageComboBox.SelectedItem is ComboBoxItem item)
            {
                var locale = (item.Tag?.ToString() ?? string.Empty).Trim();
                if (!string.IsNullOrWhiteSpace(locale)) return locale;
                return (item.Content?.ToString() ?? "en-US").Trim();
            }
            return "en-IN";
        }
    }

    public bool SimpleLanguage => SimpleLanguageToggle.IsChecked ?? false;
    public bool NaturalSpeakingMode => NaturalSpeakingModeToggle.IsChecked ?? false;
    public string ExtraContext => (ExtraContextTextBox.Text ?? string.Empty).Trim();

    public bool SaveTranscript => SaveTranscriptToggle.IsChecked ?? false;

    public void ResetForNewSession()
    {
        EnsureLanguageOptionsLoaded();
        LanguageComboBox.SelectedIndex = 0; // English (India)

        SimpleLanguageToggle.IsChecked = true;
        NaturalSpeakingModeToggle.IsChecked = false;
        ExtraContextTextBox.Text = string.Empty;
        RefreshExtraContextPlaceholder();
        SaveTranscriptToggle.IsChecked = true;
    }

    public void SetSessionMode(bool isFreeSession)
    {
        _isFreeSession = isFreeSession;
        CreateSessionButtonText.Text = _isFreeSession ? "Create Free Session" : "Create Session";
    }
}
