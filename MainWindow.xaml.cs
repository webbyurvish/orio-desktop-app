using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Text.Json;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using System.Windows.Controls;
using System.Windows.Documents;
using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;

namespace AiInterviewAssistant;

public partial class MainWindow : Window
{
    private enum AudioQuestionSource
    {
        Interviewer,
        SelfMic
    }

    private static readonly HashSet<string> CodeKeywords = new(StringComparer.OrdinalIgnoreCase)
    {
        "select","from","where","join","left","right","inner","outer","group","by","order","limit","insert","into","values","update","set","delete","create","table",
        "if","else","for","foreach","while","return","class","public","private","protected","static","void","string","int","bool","var","new","using","async","await","try","catch","finally","switch","case","break","continue","true","false","null"
    };

    private static readonly Regex InterviewerQuestionLeadInRegex = new(
        @"^(what|why|how|when|where|who|whom|which|whose)\b|" +
        @"^(can|could|would|should|will|do|does|did|is|are|was|were|have|has|had)\s+(you|we|they|i|she|he|it|there|this|that)\b|" +
        @"^(tell|describe|explain|outline)\s+me\b|" +
        @"^(walk)\s+me\s+through\b|" +
        @"^(give|name|list)\s+(me\s+)?(a|an|the|some|three|your)?\b",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly string AnswerFromTranscriptSystemPrompt =
        "You are role-playing as the job candidate whose resume is provided in the context. " +
        "Always answer in FIRST PERSON as that candidate (for example: 'My name is ...', 'I have 5 years of experience ...'). " +
        "Never say you are an AI or language model. " +
        "The user will give you text that was transcribed from speech. " +
        "Answer the question fully and in detail, using the resume details whenever relevant. " +
        "For questions like 'what is your name' or 'introduce yourself', answer using the candidate's real name and background from the resume.";

    private static string MapSessionLanguageToAzureSpeechLocale(string? sessionLanguage)
    {
        var lang = (sessionLanguage ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(lang))
            return "en-IN";

        // If user already provided a locale (e.g. "hi-IN"), pass through.
        if (lang.Length == 5 && lang[2] == '-' &&
            char.IsLetter(lang[0]) && char.IsLetter(lang[1]) &&
            char.IsLetter(lang[3]) && char.IsLetter(lang[4]))
        {
            return lang;
        }

        return lang.ToLowerInvariant() switch
        {
            "english" => "en-IN",
            "hindi" => "hi-IN",
            // Common variants / aliases
            "en" => "en-IN",
            "hi" => "hi-IN",
            _ => "en-IN"
        };
    }

    private static string GetOutputLanguageNameForPrompt(string? sessionLanguageOrLocale)
    {
        var locale = MapSessionLanguageToAzureSpeechLocale(sessionLanguageOrLocale);
        var prefix = locale.Split('-')[0].ToLowerInvariant();
        return prefix switch
        {
            "en" => "English",
            "hi" => "Hindi",
            "bn" => "Bengali",
            "gu" => "Gujarati",
            "kn" => "Kannada",
            "ml" => "Malayalam",
            "mr" => "Marathi",
            "pa" => "Punjabi",
            "ta" => "Tamil",
            "te" => "Telugu",
            "ur" => "Urdu",
            "ar" => "Arabic",
            "de" => "German",
            "es" => "Spanish",
            "fr" => "French",
            "he" => "Hebrew",
            "id" => "Indonesian",
            "it" => "Italian",
            "ja" => "Japanese",
            "ko" => "Korean",
            "ms" => "Malay",
            "nl" => "Dutch",
            "pl" => "Polish",
            "pt" => "Portuguese",
            "ru" => "Russian",
            "sw" => "Swahili",
            "th" => "Thai",
            "tr" => "Turkish",
            "uk" => "Ukrainian",
            "vi" => "Vietnamese",
            "zh" => "Chinese",
            "fa" => "Persian",
            _ => "English"
        };
    }
    private sealed class CreateCallSessionRequest
    {
        public string? Title { get; set; }
        public string? Description { get; set; }
        public Guid? ResumeId { get; set; }
        public string? Language { get; set; }
        public bool SimpleLanguage { get; set; }
        public string? ExtraContext { get; set; }
        public string? AiModel { get; set; }
        public bool SaveTranscript { get; set; }
        public bool IsFreeSession { get; set; }
    }

    private sealed class CallSessionDto
    {
        public Guid Id { get; set; }
        public Guid? ResumeId { get; set; }
    }

    private sealed class CurrentUserDto
    {
        public decimal CallCredits { get; set; }
        public string? Email { get; set; }
    }

    private const int WDA_NONE = 0;
    private const int WDA_EXCLUDEFROMCAPTURE = 0x11;

    [DllImport("user32.dll")]
    private static extern bool SetWindowDisplayAffinity(IntPtr hWnd, int affinity);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowPos(
        IntPtr hWnd,
        IntPtr hWndInsertAfter,
        int X,
        int Y,
        int cx,
        int cy,
        uint uFlags);

    private const int GWL_EXSTYLE = -20;
    private const int GWL_STYLE = -16;
    private const int WS_EX_CLIENTEDGE = 0x200;
    private const int WS_EX_STATICEDGE = 0x20000;
    private const int WS_EX_TOOLWINDOW = 0x80;
    private const int WS_EX_APPWINDOW = 0x40000;
    private const int WS_BORDER = 0x00800000;

    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_NOMOVE = 0x0002;
    private const uint SWP_NOZORDER = 0x0004;
    private const uint SWP_FRAMECHANGED = 0x0020;

    [DllImport("dwmapi.dll", SetLastError = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
    private const int DWMWA_BORDER_COLOR = 34;

    private SpeechRecognizer? _speechRecognizer; // microphone
    private SpeechRecognizer? _systemSpeechRecognizer; // system audio (loopback)
    private PushAudioInputStream? _systemPushStream;
    private WasapiLoopbackCapture? _loopbackCapture;
    private object _systemPushLock = new();

    private readonly StringBuilder _finalTranscript = new();
    private string _partialMic = string.Empty;
    private string _partialSystem = string.Empty;
    private bool _isListening;
    private bool _micOn = true;
    private bool _speakerOn = true;

    private readonly string _deploymentName;
    private readonly string _systemPrompt;
    private sealed class DesktopAiAnswerRequest
    {
        public string? UserContent { get; set; }
        public string? SystemPrompt { get; set; }
        public string? ResumeContext { get; set; }
    }

    private sealed class DesktopScreenshotAnswerRequest
    {
        public string? ImageBase64 { get; set; }
        public string? MimeType { get; set; }
        public string? SystemPrompt { get; set; }
        public string? ResumeContext { get; set; }
    }

    private sealed class DesktopSpeechTokenResponse
    {
        public string Region { get; set; } = string.Empty;
        public string Token { get; set; } = string.Empty;
        public int ExpiresInSeconds { get; set; }
    }

    private sealed class CallSessionMessageDto
    {
        public Guid Id { get; set; }
        public string Role { get; set; } = string.Empty;
        public string Content { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; }
    }

    private sealed class AnswerHistoryItem
    {
        public string? Heading { get; set; }
        public string Content { get; set; } = string.Empty;
        public Guid? ServerMessageId { get; set; }
    }

    private readonly HttpClient _apiClient;
    private Guid _callSessionId;
    private readonly Guid? _resumeId;
    private string? _resumeContext;
    private bool _mainUiStarted;
    private RestoreChipWindow? _restoreChip;
    private bool _startupPositionInitialized;
    private readonly DesktopWebAuthService _desktopWebAuthService;

    /// <summary>Live text target while Azure streams tokens; replaced when <see cref="RenderAiAnswer"/> runs.</summary>
    private Run? _streamingAnswerRun;

    /// <summary>Bold heading shown above the answer (question sent to the user); cleared when answer is cleared.</summary>
    private string? _currentAnswerDisplayHeading;
    private readonly SemaphoreSlim _loginFlowLock = new(1, 1);
    private readonly object _loginSync = new();
    private CancellationTokenSource? _activeLoginCts;
    private int _loginAttemptVersion;
    private bool _isFreeSessionFlow = true;
    private bool _pendingProtocolActivation;

    private DispatcherTimer? _sessionTimer;
    private DateTimeOffset _sessionStartUtc;
    private DateTimeOffset _nextFullExtensionUtc;
    private bool _sessionActive;
    private bool _activeSessionIsFree;
    private int _fullExtensionsApplied;
    private bool _pendingActivateSync;
    private int _pendingExtendSyncCount;
    private bool _pendingEndSync;
    private DateTimeOffset _lastServerSyncAttemptUtc;
    private bool _saveTranscriptEnabled = true;

    /// <summary>Bumped when a session ends or a new interview starts so in-flight AI answer streams cannot repaint the UI for the wrong session.</summary>
    private int _answerUiEpoch;

    /// <summary>Matches <see cref="_answerUiEpoch"/> snapshot for the answer stream that disabled the AI buttons; cleared on bump or release.</summary>
    private int _answerStreamLeaseEpoch;

    private readonly SemaphoreSlim _creditsRefreshLock = new(1, 1);
    private DateTimeOffset _lastCreditsRefreshUtc = DateTimeOffset.MinValue;

    private readonly List<AnswerHistoryItem> _sessionAnswerHistory = new();
    private int _answerHistoryViewIndex = -1;
    private int _lastAppendedAnswerHistoryIndex = -1;
    private bool _answerGenerationInFlight;
    private bool _speechInitInFlight;

    private DispatcherTimer? _callAutoAnswerDebounceTimer;
    private DispatcherTimer? _micAutoAnswerDebounceTimer;
    private string _callAutoAnswerBuffer = string.Empty;
    private string _micAutoAnswerBuffer = string.Empty;
    private string _lastAutoAnswerNormKey = string.Empty;
    private DateTimeOffset _lastAutoAnswerUtc = DateTimeOffset.MinValue;

    public MainWindow()
    {
        InitializeComponent();

        var settings = App.Settings;
        _desktopWebAuthService = new DesktopWebAuthService(settings.DesktopAuth);

        _deploymentName = settings.AzureOpenAI.DeploymentName;
        _systemPrompt = settings.AzureOpenAI.SystemPrompt;

        // HTTP client for logging conversation to dashboard API
        _apiClient = new HttpClient();
        var baseUrl = settings.ApiBaseUrl?.Trim();
        if (string.IsNullOrWhiteSpace(baseUrl))
            baseUrl = "http://localhost:5050/api/";
        _apiClient.BaseAddress = new Uri(baseUrl.EndsWith("/", StringComparison.Ordinal) ? baseUrl : baseUrl + "/");
        if (!string.IsNullOrWhiteSpace(settings.ApiBearerToken))
        {
            _apiClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", settings.ApiBearerToken.Trim());
        }
        PastSessionsView.ConfigureApiClient(_apiClient);
        CreateSessionDetailsView.ConfigureApiClient(_apiClient);
        DesktopAnalytics.Configure(_apiClient);
        if (!Guid.TryParse(settings.CallSessionId, out _callSessionId))
        {
            // If parsing fails, use the fixed session id you requested for now
            _callSessionId = Guid.Parse("AB589C99-0980-4467-8AF5-ADAB340FE1A0");
        }
        _pendingProtocolActivation = App.LaunchedViaProtocol && _callSessionId != Guid.Empty;

        if (Guid.TryParse(settings.ResumeId, out var rid))
        {
            _resumeId = rid;
        }

        // Show which session/API we are logging to (helps debug protocol launch).
        StatusTextBlock.Text =
            $"Logging → {_apiClient.BaseAddress} | Session {_callSessionId}" + (App.LaunchedViaProtocol ? " (protocol)" : "") +
            Environment.NewLine +
            $"Log → {DesktopLogger.LogFilePath}" +
            Environment.NewLine +
            $"Fallback Log → {DesktopLogger.FallbackLogFilePath}";
        DesktopLogger.Info($"MainWindow init. baseAddress={_apiClient.BaseAddress} sessionId={_callSessionId} tokenPresent={_apiClient.DefaultRequestHeaders.Authorization != null} launchedViaProtocol={App.LaunchedViaProtocol}");
        DesktopLogger.Info($"Log file: {DesktopLogger.LogFilePath}");

        // Populate credits chip as soon as we have a token (appsettings/protocol may include one).
        _ = RefreshCreditsFromServerAsync(force: true, showLoading: !CreditsState.Current.IsKnown);

        // Startup screen only; main interview UI appears after Login.
        ApplyStartupChrome();
        StartupLoginView.LoginRequested += async (_, _) => await HandleLoginRequestedAsync();
        StartupLoginView.DashboardRequested += (_, _) => OpenWebDashboardHome();
        StartupLoginView.LogoutRequested += async (_, _) => await PerformLogoutAsync();
        StartupLoginView.CloseRequested += (_, _) => Close();
        StartupLoginView.MinimizeRequested += (_, _) => MinimizeToRestoreChip();
        StartupLoginView.WindowSlotRequested += ApplyStartupWindowSlot;
        SessionSetupView.FullSessionRequested += (_, _) =>
        {
            _isFreeSessionFlow = false;
            ResetCreateSessionDraftForNewFlow();
            ShowCreateSessionDetailsView();
        };
        SessionSetupView.FreeSessionRequested += (_, _) =>
        {
            _isFreeSessionFlow = true;
            ResetCreateSessionDraftForNewFlow();
            ShowCreateSessionDetailsView();
        };
        SessionSetupView.CloseRequested += (_, _) => Close();
        SessionSetupView.MinimizeRequested += (_, _) => MinimizeToRestoreChip();
        SessionSetupView.WindowSlotRequested += ApplyStartupWindowSlot;
        SessionSetupView.PastSessionsRequested += (_, _) => ShowPastSessionsView();
        PastSessionsView.CreateRequested += (_, _) =>
        {
            ResetCreateSessionDraftForNewFlow();
            ShowSessionSetupView();
        };
        PastSessionsView.ViewAllRequested += (_, _) => OpenPastSessionsInBrowser();
        PastSessionsView.ActivateNotActivatedRequested += (sessionId, isFree, language) =>
        {
            _callSessionId = sessionId;
            _isFreeSessionFlow = isFree;
            App.Settings.SessionLanguage = language;
            ShowActivateSessionView();
        };
        PastSessionsView.CloseRequested += (_, _) => Close();
        PastSessionsView.MinimizeRequested += (_, _) => MinimizeToRestoreChip();
        PastSessionsView.WindowSlotRequested += ApplyStartupWindowSlot;
        CreateSessionDetailsView.BackRequested += (_, _) => ShowSessionSetupView();
        CreateSessionDetailsView.NextRequested += (_, _) => ShowCreateSessionStep2View();
        CreateSessionDetailsView.PastSessionsRequested += (_, _) => ShowPastSessionsView();
        CreateSessionDetailsView.CloseRequested += (_, _) => Close();
        CreateSessionDetailsView.MinimizeRequested += (_, _) => MinimizeToRestoreChip();
        CreateSessionDetailsView.WindowSlotRequested += ApplyStartupWindowSlot;
        CreateSessionStep2View.BackRequested += (_, _) => ShowCreateSessionDetailsView();
        CreateSessionStep2View.CreateSessionRequested += async (_, _) => await HandleCreateFreeSessionRequestedAsync();
        CreateSessionStep2View.PastSessionsRequested += (_, _) => ShowPastSessionsView();
        CreateSessionStep2View.CloseRequested += (_, _) => Close();
        CreateSessionStep2View.MinimizeRequested += (_, _) => MinimizeToRestoreChip();
        CreateSessionStep2View.WindowSlotRequested += ApplyStartupWindowSlot;
        ActivateSessionView.BackRequested += (_, _) => ShowCreateSessionStep2View();
        ActivateSessionView.ActivateRequested += (_, _) => StartInterviewSessionAsync();
        ActivateSessionView.CloseRequested += (_, _) => Close();
        ActivateSessionView.MinimizeRequested += (_, _) => MinimizeToRestoreChip();
        ActivateSessionView.WindowSlotRequested += ApplyStartupWindowSlot;
        SessionSetupView.DashboardRequested += (_, _) => OpenWebDashboardHome();
        SessionSetupView.LogoutRequested += async (_, _) => await PerformLogoutAsync();
        PastSessionsView.DashboardRequested += (_, _) => OpenWebDashboardHome();
        PastSessionsView.LogoutRequested += async (_, _) => await PerformLogoutAsync();
        CreateSessionDetailsView.DashboardRequested += (_, _) => OpenWebDashboardHome();
        CreateSessionDetailsView.LogoutRequested += async (_, _) => await PerformLogoutAsync();
        CreateSessionStep2View.DashboardRequested += (_, _) => OpenWebDashboardHome();
        CreateSessionStep2View.LogoutRequested += async (_, _) => await PerformLogoutAsync();
        ActivateSessionView.DashboardRequested += (_, _) => OpenWebDashboardHome();
        ActivateSessionView.LogoutRequested += async (_, _) => await PerformLogoutAsync();

        // If app is opened via protocol and auth already exists, skip startup/login and go straight to activation.
        NavigateAfterAuthentication(source: "startup");

        SizeChanged += (_, _) => UpdateAiAnswerBodyMaxHeight();
        LocationChanged += (_, _) => UpdateAiAnswerBodyMaxHeight();

        _callAutoAnswerDebounceTimer = new DispatcherTimer(DispatcherPriority.Background)
        {
            Interval = TimeSpan.FromMilliseconds(250)
        };
        _callAutoAnswerDebounceTimer.Tick += (_, _) => OnCallAutoAnswerDebounceTick();
        _micAutoAnswerDebounceTimer = new DispatcherTimer(DispatcherPriority.Background)
        {
            Interval = TimeSpan.FromMilliseconds(250)
        };
        _micAutoAnswerDebounceTimer.Tick += (_, _) => OnMicAutoAnswerDebounceTick();
    }

    /// <summary>
    /// Caps the answer host height to the visible desktop below it so the main RichTextBox gets a finite height
    /// and shows its vertical scrollbar instead of clipping content.
    /// </summary>
    private void UpdateAiAnswerBodyMaxHeight()
    {
        if (!_mainUiStarted || AnswerSectionPanel.Visibility != Visibility.Visible)
        {
            AiAnswerBodyBorder.ClearValue(FrameworkElement.MaxHeightProperty);
            return;
        }

        try
        {
            var wa = SystemParameters.WorkArea;
            var topScreen = AiAnswerBodyBorder.PointToScreen(new Point(0, 0));
            const double bottomPad = 16;
            var maxH = wa.Bottom - topScreen.Y - bottomPad;
            if (double.IsNaN(maxH) || double.IsInfinity(maxH))
                maxH = 360;
            AiAnswerBodyBorder.MaxHeight = Math.Max(120, maxH);
        }
        catch
        {
            AiAnswerBodyBorder.MaxHeight = Math.Max(200, Math.Min(520, SystemParameters.WorkArea.Height * 0.55));
        }
    }

    /// <summary>
    /// Hides the main window and shows a small accent circle; click to restore (login / app UI).
    /// </summary>
    private void MinimizeToRestoreChip()
    {
        if (_restoreChip == null)
        {
            _restoreChip = new RestoreChipWindow();
            _restoreChip.RestoreRequested += OnRestoreChipRestoreRequested;
        }

        var wa = SystemParameters.WorkArea;
        const double topMargin = 10;
        double chipW = _restoreChip.Width;
        double chipH = _restoreChip.Height;
        // Always show restore chip at top-center of the current work area.
        double left = wa.Left + (wa.Width - chipW) / 2.0;
        double top = wa.Top + topMargin;

        _restoreChip.Left = left;
        _restoreChip.Top = top;
        _restoreChip.Show();
        _restoreChip.Activate();
        Hide();
    }

    /// <summary>
    /// Restores the main window if it was hidden/minimized and hides the restore chip.
    /// Used when returning from external flows (browser / protocol activation).
    /// </summary>
    private void RestoreFromRestoreChipIfNeeded()
    {
        try
        {
            // If we're hidden, Show() brings us back. If we're minimized, normalize.
            Show();
            if (WindowState == WindowState.Minimized)
                WindowState = WindowState.Normal;
            Activate();
            try { Focus(); } catch { /* ignore */ }
            _restoreChip?.Hide();
        }
        catch
        {
            // best-effort only
        }
    }

    private void OnRestoreChipRestoreRequested(object? sender, EventArgs e)
    {
        RestoreFromRestoreChipIfNeeded();
    }

    private void ApplyStartupWindowSlot(StartupWindowSlot slot)
    {
        var wa = SystemParameters.WorkArea;
        // Ensure we can position exactly; startup view still controls content.
        SizeToContent = SizeToContent.WidthAndHeight;
        UpdateLayout();

        var w = ActualWidth > 0 ? ActualWidth : Width;
        var h = ActualHeight > 0 ? ActualHeight : Height;
        if (double.IsNaN(w) || w <= 0) w = 520;
        if (double.IsNaN(h) || h <= 0) h = 260;

        double left = wa.Left;
        double top = wa.Top;

        switch (slot)
        {
            case StartupWindowSlot.TopLeft:
                left = wa.Left + 10;
                top = wa.Top + 10;
                break;
            case StartupWindowSlot.TopCenter:
                left = wa.Left + (wa.Width - w) / 2;
                top = wa.Top + 10;
                break;
            case StartupWindowSlot.TopRight:
                left = wa.Right - w - 10;
                top = wa.Top + 10;
                break;
            case StartupWindowSlot.BottomLeft:
                left = wa.Left + 10;
                top = wa.Bottom - h - 10;
                break;
            case StartupWindowSlot.BottomCenter:
                left = wa.Left + (wa.Width - w) / 2;
                top = wa.Bottom - h - 10;
                break;
            case StartupWindowSlot.BottomRight:
                left = wa.Right - w - 10;
                top = wa.Bottom - h - 10;
                break;
        }

        Left = Math.Max(wa.Left, Math.Min(left, wa.Right - w));
        Top = Math.Max(wa.Top, Math.Min(top, wa.Bottom - h));
    }

    private void ApplyStartupChrome()
    {
        ResizeMode = ResizeMode.NoResize;
        StartupLoginView.Visibility = Visibility.Visible;
        SessionSetupView.Visibility = Visibility.Collapsed;
        MainContentGrid.Visibility = Visibility.Collapsed;
        // No fill behind the login card — avoids light “halo” around the rounded box; show-through = real transparency
        RootChromeBorder.Background = Brushes.Transparent;
        RootChromeBorder.Padding = new Thickness(0);
        RootChromeBorder.CornerRadius = new CornerRadius(0);
        // Hug the card so the window isn’t wider/taller than the rounded content
        Width = double.NaN;
        SizeToContent = SizeToContent.WidthAndHeight;
        TrySetDwmBorderColor(0x00FAFBFC); // BGR tint near card #FAFBFC (subtle frame when transparency shows DWM edge)

        // Initial launch position: top-center (2nd column, 1st row in the 3x2 move grid).
        if (!_startupPositionInitialized)
        {
            _startupPositionInitialized = true;
            Dispatcher.BeginInvoke(new Action(() => ApplyStartupWindowSlot(StartupWindowSlot.TopCenter)));
        }
    }

    private async Task HandleLoginRequestedAsync()
    {
        int attemptVersion = Interlocked.Increment(ref _loginAttemptVersion);
        string authTraceId = $"{DateTime.UtcNow:HHmmss}-{attemptVersion}";
        var timeoutSeconds = Math.Max(30, App.Settings.DesktopAuth.LoginTimeoutSeconds);
        var attemptCts = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutSeconds));

        CancellationTokenSource? previousCts;
        lock (_loginSync)
        {
            previousCts = _activeLoginCts;
            _activeLoginCts = attemptCts;
        }

        if (previousCts != null)
        {
            StartupLoginView.SetLoginStatus("Restarting login...");
            DesktopLogger.Warn($"[AUTH:{authTraceId}] Previous login attempt cancelled because a new click arrived.");
            previousCts.Cancel();
            previousCts.Dispose();
        }

        await _loginFlowLock.WaitAsync();
        try
        {
            if (attemptVersion != _loginAttemptVersion)
            {
                DesktopLogger.Warn($"[AUTH:{authTraceId}] Aborted before start because another newer attempt exists.");
                return;
            }

            StartupLoginView.SetLoginBusy(true, "Opening browser for secure login...");
            DesktopLogger.Info($"[AUTH:{authTraceId}] Desktop login attempt started timeout={timeoutSeconds}s authorizeUrl={App.Settings.DesktopAuth.AuthorizeUrl} exchangeUrl={App.Settings.DesktopAuth.ExchangeUrl}");

            // User is switching to the browser flow; keep desktop out of the way.
            MinimizeToRestoreChip();

            DesktopAuthTokens tokens = await _desktopWebAuthService.AuthenticateAsync(attemptCts.Token, authTraceId);
            _apiClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", tokens.AccessToken);
            DesktopLogger.Info($"[AUTH:{authTraceId}] Desktop login successful. API bearer updated accessLen={tokens.AccessToken.Length} refreshLen={tokens.RefreshToken.Length}");
            StartupLoginView.SetLoginStatus("Login successful.");
            // The login finished; bring the window back to the foreground.
            RestoreFromRestoreChipIfNeeded();
            await RefreshCreditsFromServerAsync(force: true, showLoading: false).ConfigureAwait(true);
            NavigateAfterAuthentication(source: $"login:{authTraceId}");
        }
        catch (OperationCanceledException)
        {
            var cancelledByNewAttempt = attemptVersion != _loginAttemptVersion;
            var message = cancelledByNewAttempt
                ? "Login restarted..."
                : "Login not completed (timed out). Please try again.";
            StartupLoginView.SetLoginStatus(message);
            DesktopLogger.Warn($"[AUTH:{authTraceId}] Desktop login cancelled cancelledByNewAttempt={cancelledByNewAttempt} elapsedTimeout={timeoutSeconds}s");
        }
        catch (Exception ex)
        {
            StartupLoginView.SetLoginStatus($"Login failed: {ex.Message}");
            DesktopLogger.Error($"[AUTH:{authTraceId}] Desktop login failed: {ex}");
        }
        finally
        {
            lock (_loginSync)
            {
                if (ReferenceEquals(_activeLoginCts, attemptCts))
                    _activeLoginCts = null;
            }

            attemptCts.Dispose();
            StartupLoginView.SetLoginBusy(false);
            DesktopLogger.Info($"[AUTH:{authTraceId}] Desktop login attempt finished and lock released.");
            _loginFlowLock.Release();
        }
    }

    private bool IsApiAuthenticated() =>
        _apiClient.DefaultRequestHeaders.Authorization != null;

    private void NavigateAfterAuthentication(string source)
    {
        if (!IsApiAuthenticated())
            return;

        // If the app was hidden while the user interacted with the browser, restore it.
        RestoreFromRestoreChipIfNeeded();

        if (_pendingProtocolActivation)
        {
            _pendingProtocolActivation = false;
            DesktopLogger.Info($"Protocol launch session detected. Routing to ActivateSessionView source={source} sessionId={_callSessionId}");
            ShowActivateSessionView();
            return;
        }

        ShowSessionSetupView();
    }

    private void ShowSessionSetupView()
    {
        ResizeMode = ResizeMode.NoResize;
        StartupLoginView.Visibility = Visibility.Collapsed;
        SessionSetupView.Visibility = Visibility.Visible;
        PastSessionsView.Visibility = Visibility.Collapsed;
        CreateSessionDetailsView.Visibility = Visibility.Collapsed;
        CreateSessionStep2View.Visibility = Visibility.Collapsed;
        ActivateSessionView.Visibility = Visibility.Collapsed;
        MainContentGrid.Visibility = Visibility.Collapsed;
        RootChromeBorder.Background = Brushes.Transparent;
        RootChromeBorder.Padding = new Thickness(0);
        RootChromeBorder.CornerRadius = new CornerRadius(0);
        Width = double.NaN;
        SizeToContent = SizeToContent.WidthAndHeight;
        TrySetDwmBorderColor(0x00FAFBFC);
    }

    private void ShowPastSessionsView()
    {
        ResizeMode = ResizeMode.NoResize;
        StartupLoginView.Visibility = Visibility.Collapsed;
        SessionSetupView.Visibility = Visibility.Collapsed;
        PastSessionsView.Visibility = Visibility.Visible;
        CreateSessionDetailsView.Visibility = Visibility.Collapsed;
        CreateSessionStep2View.Visibility = Visibility.Collapsed;
        ActivateSessionView.Visibility = Visibility.Collapsed;
        MainContentGrid.Visibility = Visibility.Collapsed;
        RootChromeBorder.Background = Brushes.Transparent;
        RootChromeBorder.Padding = new Thickness(0);
        RootChromeBorder.CornerRadius = new CornerRadius(0);
        Width = double.NaN;
        SizeToContent = SizeToContent.WidthAndHeight;
        TrySetDwmBorderColor(0x00FAFBFC);
        _ = PastSessionsView.ReloadSessionsAsync();
    }

    private static string GetWebAppOrigin()
    {
        var fallback = "http://localhost:5173";
        var authorizeUrl = App.Settings.DesktopAuth.AuthorizeUrl?.Trim();
        if (!string.IsNullOrWhiteSpace(authorizeUrl) && Uri.TryCreate(authorizeUrl, UriKind.Absolute, out var authUri))
            return $"{authUri.Scheme}://{authUri.Authority}";
        return fallback;
    }

    private void OpenWebDashboardHome()
    {
        var targetUrl = $"{GetWebAppOrigin().TrimEnd('/')}/dashboard";
        try
        {
            MinimizeToRestoreChip();
            Process.Start(new ProcessStartInfo { FileName = targetUrl, UseShellExecute = true });
            DesktopLogger.Info($"Opened web dashboard: {targetUrl}");
        }
        catch (Exception ex)
        {
            DesktopLogger.Error($"Failed to open web dashboard: {ex}");
            MessageBox.Show("Unable to open web dashboard.", "Dashboard", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OpenPastSessionsInBrowser()
    {
        var targetUrl = $"{GetWebAppOrigin().TrimEnd('/')}/dashboard/call-sessions";

        try
        {
            MinimizeToRestoreChip();
            Process.Start(new ProcessStartInfo
            {
                FileName = targetUrl,
                UseShellExecute = true
            });
            DesktopLogger.Info($"Opened past sessions web page: {targetUrl}");
        }
        catch (Exception ex)
        {
            DesktopLogger.Error($"Failed to open past sessions web page: {ex}");
            MessageBox.Show("Unable to open web dashboard.", "Open Dashboard", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task PerformLogoutAsync()
    {
        DesktopLogger.Info("User signed out from desktop app.");
        try
        {
            if (_mainUiStarted)
            {
                await StopSpeechSessionAsync();
                _mainUiStarted = false;
            }

            SecureTokenStore.ClearPersistedTokens();
            _apiClient.DefaultRequestHeaders.Authorization = null;
            CreditsState.Current.SetUnknown();
            DesktopUserState.Current.Clear();

            _finalTranscript.Clear();
            _partialMic = string.Empty;
            _partialSystem = string.Empty;
            TranscriptTextBlock.Text = string.Empty;

            ResetSessionAnswerHistoryForInterview();
            ResetAutoAnswerTransientState();
            BumpAnswerUiEpoch();
            ResetInterviewAnswerUi();

            if (Guid.TryParse(App.Settings.CallSessionId, out var sid))
                _callSessionId = sid;
            else
                _callSessionId = Guid.Parse("AB589C99-0980-4467-8AF5-ADAB340FE1A0");

            StatusTextBlock.Text =
                $"Logging → {_apiClient.BaseAddress} | Session {_callSessionId}" + (App.LaunchedViaProtocol ? " (protocol)" : "") +
                Environment.NewLine +
                $"Log → {DesktopLogger.LogFilePath}" +
                Environment.NewLine +
                $"Fallback Log → {DesktopLogger.FallbackLogFilePath}";

            StartupLoginView.SetLoginBusy(false);
            StartupLoginView.SetLoginStatus("Signed out. Sign in again to continue.");
            ApplyStartupChrome();
        }
        catch (Exception ex)
        {
            DesktopLogger.Error($"Logout error: {ex}");
            MessageBox.Show("Sign out could not complete. You can close the app and try again.", "Logout", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private async Task StopSpeechSessionAsync()
    {
        try
        {
            _loopbackCapture?.StopRecording();
            _loopbackCapture?.Dispose();
        }
        catch { /* ignore */ }

        _loopbackCapture = null;

        if (_systemSpeechRecognizer != null)
        {
            try { await _systemSpeechRecognizer.StopContinuousRecognitionAsync(); } catch { /* ignore */ }
            try { _systemSpeechRecognizer.Dispose(); } catch { /* ignore */ }
            _systemSpeechRecognizer = null;
        }

        try { _systemPushStream?.Close(); } catch { /* ignore */ }
        _systemPushStream = null;

        if (_speechRecognizer != null)
        {
            try
            {
                if (_isListening)
                    await _speechRecognizer.StopContinuousRecognitionAsync();
            }
            catch { /* ignore */ }

            try { _speechRecognizer.Dispose(); } catch { /* ignore */ }
            _speechRecognizer = null;
        }

        _isListening = false;
    }

    private void ShowCreateSessionDetailsView()
    {
        ResizeMode = ResizeMode.NoResize;
        StartupLoginView.Visibility = Visibility.Collapsed;
        SessionSetupView.Visibility = Visibility.Collapsed;
        PastSessionsView.Visibility = Visibility.Collapsed;
        CreateSessionDetailsView.Visibility = Visibility.Visible;
        CreateSessionStep2View.Visibility = Visibility.Collapsed;
        ActivateSessionView.Visibility = Visibility.Collapsed;
        MainContentGrid.Visibility = Visibility.Collapsed;
        RootChromeBorder.Background = Brushes.Transparent;
        RootChromeBorder.Padding = new Thickness(0);
        RootChromeBorder.CornerRadius = new CornerRadius(0);
        Width = double.NaN;
        SizeToContent = SizeToContent.WidthAndHeight;
        TrySetDwmBorderColor(0x00FAFBFC);
        _ = CreateSessionDetailsView.ReloadResumesAsync();
    }

    private void ShowCreateSessionStep2View()
    {
        ResizeMode = ResizeMode.NoResize;
        StartupLoginView.Visibility = Visibility.Collapsed;
        SessionSetupView.Visibility = Visibility.Collapsed;
        PastSessionsView.Visibility = Visibility.Collapsed;
        CreateSessionDetailsView.Visibility = Visibility.Collapsed;
        CreateSessionStep2View.Visibility = Visibility.Visible;
        ActivateSessionView.Visibility = Visibility.Collapsed;
        MainContentGrid.Visibility = Visibility.Collapsed;
        RootChromeBorder.Background = Brushes.Transparent;
        RootChromeBorder.Padding = new Thickness(0);
        RootChromeBorder.CornerRadius = new CornerRadius(0);
        Width = double.NaN;
        SizeToContent = SizeToContent.WidthAndHeight;
        TrySetDwmBorderColor(0x00FAFBFC);
        CreateSessionStep2View.SetSessionMode(_isFreeSessionFlow);
    }

    private void ResetCreateSessionDraftForNewFlow()
    {
        CreateSessionDetailsView.ResetForNewSession();
        CreateSessionStep2View.ResetForNewSession();
    }

    private async Task HandleCreateFreeSessionRequestedAsync()
    {
        var company = CreateSessionDetailsView.Company;
        var jobDescription = CreateSessionDetailsView.JobDescription;
        if (string.IsNullOrWhiteSpace(company) || string.IsNullOrWhiteSpace(jobDescription))
        {
            MessageBox.Show("Company and Job Description are required.", "Create Session", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            var payload = new CreateCallSessionRequest
            {
                Title = company,
                Description = jobDescription,
                ResumeId = CreateSessionDetailsView.SelectedResumeId,
                Language = CreateSessionStep2View.SelectedLanguage,
                SimpleLanguage = CreateSessionStep2View.SimpleLanguage,
                ExtraContext = CreateSessionStep2View.ExtraContext,
                AiModel = CreateSessionStep2View.SelectedAiModel,
                SaveTranscript = CreateSessionStep2View.SaveTranscript,
                IsFreeSession = _isFreeSessionFlow
            };
            App.Settings.SessionLanguage = payload.Language;
            _saveTranscriptEnabled = payload.SaveTranscript;

            DesktopLogger.Info($"Create session request titleLen={payload.Title?.Length ?? 0} hasResume={payload.ResumeId.HasValue} language={payload.Language} saveTranscript={payload.SaveTranscript} isFreeSession={payload.IsFreeSession}");
            using var response = await _apiClient.PostAsJsonAsync("callsessions", payload);
            var responseBody = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
            {
                DesktopLogger.Warn($"Create free session failed status={(int)response.StatusCode} body={responseBody}");
                MessageBox.Show($"Create session failed: {(int)response.StatusCode}", "Create Session", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            var created = JsonSerializer.Deserialize<CallSessionDto>(responseBody, options);
            if (created == null || created.Id == Guid.Empty)
            {
                MessageBox.Show("Create session failed: invalid server response.", "Create Session", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            _callSessionId = created.Id;
            DesktopLogger.Info($"Create session success sessionId={_callSessionId} isFreeSession={payload.IsFreeSession}");
            if (created.ResumeId.HasValue)
                _ = LoadResumeContextAsync(created.ResumeId.Value);

            StatusTextBlock.Text =
                $"Logging -> {_apiClient.BaseAddress} | Session {_callSessionId}" + (App.LaunchedViaProtocol ? " (protocol)" : "") +
                Environment.NewLine +
                $"Log -> {DesktopLogger.LogFilePath}" +
                Environment.NewLine +
                $"Fallback Log -> {DesktopLogger.FallbackLogFilePath}";

            ShowActivateSessionView();
        }
        catch (Exception ex)
        {
            DesktopLogger.Error($"Create free session exception: {ex}");
            MessageBox.Show($"Create session failed: {ex.Message}", "Create Session", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ShowActivateSessionView()
    {
        ActivateSessionView.PrepareForDisplay();

        ResizeMode = ResizeMode.NoResize;
        StartupLoginView.Visibility = Visibility.Collapsed;
        SessionSetupView.Visibility = Visibility.Collapsed;
        PastSessionsView.Visibility = Visibility.Collapsed;
        CreateSessionDetailsView.Visibility = Visibility.Collapsed;
        CreateSessionStep2View.Visibility = Visibility.Collapsed;
        ActivateSessionView.Visibility = Visibility.Visible;
        MainContentGrid.Visibility = Visibility.Collapsed;
        RootChromeBorder.Background = Brushes.Transparent;
        RootChromeBorder.Padding = new Thickness(0);
        RootChromeBorder.CornerRadius = new CornerRadius(0);
        Width = double.NaN;
        SizeToContent = SizeToContent.WidthAndHeight;
        TrySetDwmBorderColor(0x00FAFBFC);
    }

    private async Task RefreshCreditsFromServerAsync(bool force = false, bool showLoading = false)
    {
        try
        {
            if (!IsApiAuthenticated())
            {
                await Dispatcher.InvokeAsync(() => CreditsState.Current.SetUnknown());
                await Dispatcher.InvokeAsync(() => DesktopUserState.Current.Clear());
                return;
            }

            var now = DateTimeOffset.UtcNow;
            if (!force && (now - _lastCreditsRefreshUtc) < TimeSpan.FromSeconds(20))
                return;

            if (!await _creditsRefreshLock.WaitAsync(0).ConfigureAwait(false))
                return;

            try
            {
                now = DateTimeOffset.UtcNow;
                if (!force && (now - _lastCreditsRefreshUtc) < TimeSpan.FromSeconds(20))
                    return;

                if (showLoading)
                    await Dispatcher.InvokeAsync(() => CreditsState.Current.SetLoading());

                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var me = await _apiClient.GetFromJsonAsync<CurrentUserDto>("auth/me", options).ConfigureAwait(false);
                if (me == null)
                {
                    await Dispatcher.InvokeAsync(() => CreditsState.Current.SetUnknown());
                    await Dispatcher.InvokeAsync(() => DesktopUserState.Current.Clear());
                    return;
                }

                _lastCreditsRefreshUtc = DateTimeOffset.UtcNow;
                await Dispatcher.InvokeAsync(() => CreditsState.Current.SetCredits(me.CallCredits));
                await Dispatcher.InvokeAsync(() => DesktopUserState.Current.SetEmail(me.Email));
            }
            finally
            {
                _creditsRefreshLock.Release();
            }
        }
        catch (Exception ex)
        {
            DesktopLogger.Warn($"RefreshCreditsFromServerAsync failed: {ex.Message}");
            try { await Dispatcher.InvokeAsync(() => CreditsState.Current.SetUnknown()); } catch { /* ignore */ }
        }
    }

    private void ApplyMainChrome()
    {
        ResizeMode = ResizeMode.CanResizeWithGrip;
        StartupLoginView.Visibility = Visibility.Collapsed;
        SessionSetupView.Visibility = Visibility.Collapsed;
        PastSessionsView.Visibility = Visibility.Collapsed;
        CreateSessionDetailsView.Visibility = Visibility.Collapsed;
        CreateSessionStep2View.Visibility = Visibility.Collapsed;
        ActivateSessionView.Visibility = Visibility.Collapsed;
        MainContentGrid.Visibility = Visibility.Visible;
        SizeToContent = SizeToContent.Height;
        Width = 720;
        RootChromeBorder.Padding = new Thickness(12);
        RootChromeBorder.CornerRadius = new CornerRadius(16);
        RootChromeBorder.BorderThickness = new Thickness(1);
        Opacity = 1.0;
        ApplyChromeTranslucencyFromSlider();
        TrySetDwmBorderColor(0x00ECE7E5);
    }

    /// <summary>
    /// Translucency is applied only to chrome backgrounds (alpha on brushes). Window.Opacity is kept at 1 so
    /// answer text, code, and syntax colors stay fully readable.
    /// </summary>
    private void ApplyChromeTranslucencyFromSlider()
    {
        if (!_mainUiStarted || RootChromeBorder == null) return;

        var pct = WindowOpacitySlider?.Value ?? 90;
        var t = Math.Clamp((pct - 55.0) / 45.0, 0, 1);
        // More transparent panel at low slider, stronger dim at high end (55% .. 100%).
        byte bgA = (byte)Math.Round(40 + t * (210 - 40));
        byte borderA = (byte)Math.Round(36 + t * (160 - 36));

        // Aurora void + cool edge (matches SharedDesktopChrome aurora stealth)
        RootChromeBorder.Background = new SolidColorBrush(Color.FromArgb(bgA, 0x05, 0x05, 0x08));
        RootChromeBorder.BorderBrush = new SolidColorBrush(Color.FromArgb(borderA, 0xC7, 0xD2, 0xFE));
    }

    private void WindowOpacitySlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (!_mainUiStarted) return;
        Opacity = 1.0;
        ApplyChromeTranslucencyFromSlider();
    }

    private void TrySetDwmBorderColor(int colorRefBgr)
    {
        try
        {
            var helper = new WindowInteropHelper(this);
            if (helper.Handle == IntPtr.Zero) return;
            DwmSetWindowAttribute(helper.Handle, DWMWA_BORDER_COLOR, ref colorRefBgr, sizeof(int));
        }
        catch { /* ignore if DWM attribute not supported */ }
    }

    private async Task StartMainUiIfNeededAsync()
    {
        if (_mainUiStarted) return;
        _mainUiStarted = true;
        ApplyMainChrome();
        await EnsureSpeechConfiguredAsync().ConfigureAwait(true);
        if (_resumeId.HasValue)
        {
            _ = LoadResumeContextAsync(_resumeId.Value);
        }
    }

    private async void StartInterviewSessionAsync()
    {
        await StartMainUiIfNeededAsync().ConfigureAwait(true);
        await EnsureSpeechConfiguredAsync().ConfigureAwait(true);
        BumpAnswerUiEpoch();
        ResetSessionAnswerHistoryForInterview();
        ResetAutoAnswerTransientState();
        ResetInterviewAnswerUi();
        try
        {
            await LoadAssistantAnswerHistoryFromServerAsync().ConfigureAwait(true);
        }
        catch (Exception ex)
        {
            DesktopLogger.Warn($"StartInterviewSessionAsync history load: {ex.Message}");
        }

        _activeSessionIsFree = _isFreeSessionFlow;
        _sessionStartUtc = DateTimeOffset.UtcNow;
        _nextFullExtensionUtc = _sessionStartUtc.AddMinutes(30);
        _fullExtensionsApplied = 0;
        _sessionActive = true;

        EndSessionButton.IsEnabled = true;
        SessionTimerTextBlock.Text = "00:00";
        _pendingActivateSync = true;
        _pendingExtendSyncCount = 0;
        _pendingEndSync = false;
        _lastServerSyncAttemptUtc = DateTimeOffset.MinValue;

        if (_sessionTimer == null)
        {
            _sessionTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _sessionTimer.Tick += SessionTimer_Tick;
        }

        _sessionTimer.Stop();
        _sessionTimer.Start();

        var mode = _activeSessionIsFree ? "Free" : "Full";
        StatusTextBlock.Text = $"Session started ({mode}).";

        var cid = _callSessionId == Guid.Empty ? (Guid?)null : _callSessionId;
        DesktopAnalytics.Track(
            DesktopAnalyticsEventTypes.SessionActivated,
            JsonSerializer.Serialize(new { mode, free = _activeSessionIsFree }),
            cid);
    }

    private void ResetSessionAnswerHistoryForInterview()
    {
        _sessionAnswerHistory.Clear();
        _answerHistoryViewIndex = -1;
        _lastAppendedAnswerHistoryIndex = -1;
        if (Dispatcher.CheckAccess())
            UpdateAnswerHistoryNav();
        else
            _ = Dispatcher.InvokeAsync(UpdateAnswerHistoryNav);
    }

    private async Task LoadAssistantAnswerHistoryFromServerAsync()
    {
        if (_callSessionId == Guid.Empty) return;

        using var res = await _apiClient.GetAsync($"callsessions/{_callSessionId}/messages").ConfigureAwait(false);
        if (!res.IsSuccessStatusCode)
        {
            DesktopLogger.Warn($"GET callsessions/messages failed status={(int)res.StatusCode}");
            return;
        }

        var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        var messages = await res.Content.ReadFromJsonAsync<List<CallSessionMessageDto>>(options).ConfigureAwait(false);
        if (messages == null) return;

        var items = messages
            .Where(m => string.Equals(m.Role, "Assistant", StringComparison.OrdinalIgnoreCase))
            .OrderBy(m => m.CreatedAt)
            .Select(m => new AnswerHistoryItem
            {
                Heading = null,
                Content = m.Content ?? string.Empty,
                ServerMessageId = m.Id
            })
            .ToList();

        await Dispatcher.InvokeAsync(() =>
        {
            _sessionAnswerHistory.Clear();
            foreach (var it in items)
                _sessionAnswerHistory.Add(it);
            _answerHistoryViewIndex = items.Count > 0 ? items.Count - 1 : -1;
            _lastAppendedAnswerHistoryIndex = -1;

            if (items.Count > 0)
            {
                if (FindName("AnswerSectionPanel") is System.Windows.UIElement panel)
                {
                    panel.Visibility = Visibility.Visible;
                    SetAnswerSectionRowHeight(collapsed: false);
                }

                ApplyAnswerHistoryView();
            }
            else
            {
                ClearAiAnswer();
                if (FindName("AnswerSectionPanel") is System.Windows.UIElement emptyPanel)
                {
                    emptyPanel.Visibility = Visibility.Collapsed;
                    SetAnswerSectionRowHeight(collapsed: true);
                }

                UpdateAiAnswerBodyMaxHeight();
            }

            UpdateAnswerHistoryNav();
        });
    }

    private void UpdateAnswerHistoryNav()
    {
        if (AnswerHistoryPrevButton == null || AnswerHistoryNextButton == null || AnswerHistoryPositionText == null)
            return;

        var panelVisible = FindName("AnswerSectionPanel") is System.Windows.UIElement pan && pan.Visibility == Visibility.Visible;
        var n = _sessionAnswerHistory.Count;
        var showChrome = panelVisible && n > 1;

        AnswerHistoryPrevButton.Visibility = showChrome ? Visibility.Visible : Visibility.Collapsed;
        AnswerHistoryNextButton.Visibility = showChrome ? Visibility.Visible : Visibility.Collapsed;
        AnswerHistoryPositionText.Visibility = showChrome ? Visibility.Visible : Visibility.Collapsed;

        if (!showChrome)
            return;

        var idx = n == 0 ? -1 : Math.Clamp(_answerHistoryViewIndex, 0, n - 1);
        AnswerHistoryPositionText.Text = $"{idx + 1} / {n}";
        AnswerHistoryPrevButton.IsEnabled = idx > 0 && !_answerGenerationInFlight;
        AnswerHistoryNextButton.IsEnabled = idx < n - 1 && !_answerGenerationInFlight;
    }

    private void RegisterSuccessfulAnswerInHistory(string? completion)
    {
        if (string.IsNullOrWhiteSpace(completion))
            return;

        _sessionAnswerHistory.Add(new AnswerHistoryItem
        {
            Heading = _currentAnswerDisplayHeading,
            Content = completion,
            ServerMessageId = null
        });
        _lastAppendedAnswerHistoryIndex = _sessionAnswerHistory.Count - 1;
        _answerHistoryViewIndex = _lastAppendedAnswerHistoryIndex;
    }

    private void ApplyAnswerHistoryView()
    {
        if (_answerHistoryViewIndex < 0 || _answerHistoryViewIndex >= _sessionAnswerHistory.Count)
            return;

        var item = _sessionAnswerHistory[_answerHistoryViewIndex];
        RenderAiAnswer(item.Content, item.Heading);
        try
        {
            AiAnswerTextBlock.ScrollToHome();
        }
        catch
        {
            // ignore
        }

        UpdateAiAnswerBodyMaxHeight();
        UpdateAnswerHistoryNav();
    }

    private void AnswerHistoryPrev_Click(object sender, RoutedEventArgs e)
    {
        e.Handled = true;
        if (_answerGenerationInFlight || _sessionAnswerHistory.Count == 0) return;
        if (_answerHistoryViewIndex <= 0) return;
        _answerHistoryViewIndex--;
        ApplyAnswerHistoryView();
    }

    private void AnswerHistoryNext_Click(object sender, RoutedEventArgs e)
    {
        e.Handled = true;
        if (_answerGenerationInFlight || _sessionAnswerHistory.Count == 0) return;
        if (_answerHistoryViewIndex >= _sessionAnswerHistory.Count - 1) return;
        _answerHistoryViewIndex++;
        ApplyAnswerHistoryView();
    }

    private async Task AttachServerIdToAppendedAssistantAsync(Guid streamSessionId)
    {
        try
        {
            if (_callSessionId != streamSessionId || !_sessionActive)
                return;

            var idx = _lastAppendedAnswerHistoryIndex;
            if (idx < 0 || idx >= _sessionAnswerHistory.Count) return;

            var content = _sessionAnswerHistory[idx].Content;
            var id = await LogMessageAsync("Assistant", content).ConfigureAwait(false);
            if (!id.HasValue) return;

            await Dispatcher.InvokeAsync(() =>
            {
                if (_callSessionId != streamSessionId || !_sessionActive)
                    return;
                if (idx >= 0 && idx < _sessionAnswerHistory.Count && string.Equals(_sessionAnswerHistory[idx].Content, content, StringComparison.Ordinal))
                    _sessionAnswerHistory[idx].ServerMessageId = id;
            });
        }
        catch (Exception ex)
        {
            DesktopLogger.Warn($"AttachServerIdToAppendedAssistantAsync: {ex.Message}");
        }
    }

    private async Task<bool> ActivateCallSessionOnServerAsync()
    {
        try
        {
            if (_callSessionId == Guid.Empty) return false;
            using var res = await _apiClient.PostAsync($"callsessions/{_callSessionId}/activate", content: null);
            if (!res.IsSuccessStatusCode)
            {
                DesktopLogger.Warn($"Call session activate failed status={(int)res.StatusCode}");
                return false;
            }
            _ = RefreshCreditsFromServerAsync(force: true, showLoading: false);
            return true;
        }
        catch (Exception ex)
        {
            DesktopLogger.Warn($"Call session activate failed: {ex.Message}");
            return false;
        }
    }

    private async void SessionTimer_Tick(object? sender, EventArgs e)
    {
        if (!_sessionActive) return;

        var elapsed = DateTimeOffset.UtcNow - _sessionStartUtc;
        if (elapsed < TimeSpan.Zero) elapsed = TimeSpan.Zero;

        // MM:SS (minutes can exceed 59; keep format as requested)
        var minutes = (int)Math.Floor(elapsed.TotalMinutes);
        var seconds = elapsed.Seconds;
        SessionTimerTextBlock.Text = $"{minutes:00}:{seconds:00}";

        if (_activeSessionIsFree)
        {
            if (elapsed >= TimeSpan.FromMinutes(2))
                await EndSessionAsync("Free session ended (2:00).");
        }
        else
        {
            // Full session: extend every 30 minutes; each extension uses 0.5 credit.
            if (DateTimeOffset.UtcNow >= _nextFullExtensionUtc)
            {
                _fullExtensionsApplied++;
                _nextFullExtensionUtc = _nextFullExtensionUtc.AddMinutes(30);
                _pendingExtendSyncCount++;
                StatusTextBlock.Text = $"Session extended by 30 minutes (0.5 credit used). Extensions: {_fullExtensionsApplied}.";
            }
        }

        // Retry server sync (activate/extend/end) without blocking UI.
        var now = DateTimeOffset.UtcNow;
        if ((now - _lastServerSyncAttemptUtc) < TimeSpan.FromSeconds(5))
            return;
        _lastServerSyncAttemptUtc = now;

        if (_pendingEndSync)
        {
            if (await EndCallSessionOnServerAsync().ConfigureAwait(false))
                _pendingEndSync = false;
            return;
        }

        if (_pendingActivateSync)
        {
            if (await ActivateCallSessionOnServerAsync().ConfigureAwait(false))
                _pendingActivateSync = false;
        }

        if (_pendingExtendSyncCount > 0)
        {
            if (await ExtendCallSessionOnServerAsync().ConfigureAwait(false))
                _pendingExtendSyncCount = Math.Max(0, _pendingExtendSyncCount - 1);
        }
    }

    private async Task<bool> ExtendCallSessionOnServerAsync()
    {
        try
        {
            if (_callSessionId == Guid.Empty) return false;
            using var res = await _apiClient.PostAsync($"callsessions/{_callSessionId}/extend?minutes=30", content: null);
            if (!res.IsSuccessStatusCode)
            {
                DesktopLogger.Warn($"Call session extend failed status={(int)res.StatusCode}");
                return false;
            }
            _ = RefreshCreditsFromServerAsync(force: true, showLoading: false);
            return true;
        }
        catch (Exception ex)
        {
            DesktopLogger.Warn($"Call session extend failed: {ex.Message}");
            return false;
        }
    }

    private async void EndSessionButton_Click(object sender, RoutedEventArgs e)
    {
        await EndSessionAsync("Session ended.");
    }

    private async Task EndSessionAsync(string reason)
    {
        if (!_sessionActive) return;
        var sessionStart = _sessionStartUtc;
        var cid = _callSessionId == Guid.Empty ? (Guid?)null : _callSessionId;
        _sessionActive = false;
        BumpAnswerUiEpoch();

        DesktopAnalytics.Track(
            DesktopAnalyticsEventTypes.SessionEnded,
            JsonSerializer.Serialize(new
            {
                reason,
                minutes = Math.Round((DateTimeOffset.UtcNow - sessionStart).TotalMinutes, 2),
            }),
            cid);

        try
        {
            _sessionTimer?.Stop();
            EndSessionButton.IsEnabled = false;

            _pendingEndSync = true;
            _lastServerSyncAttemptUtc = DateTimeOffset.MinValue;
            // IMPORTANT: don't use ConfigureAwait(false) here because this method updates WPF UI.
            var endOk = await EndCallSessionOnServerAsync();
            _pendingEndSync = !endOk;
            if (endOk)
                _callSessionId = Guid.Empty;

            await StopSpeechSessionAsync();
            _mainUiStarted = false;

            _finalTranscript.Clear();
            _partialMic = string.Empty;
            _partialSystem = string.Empty;
            TranscriptTextBlock.Text = string.Empty;

            ResetSessionAnswerHistoryForInterview();
            ResetAutoAnswerTransientState();
            ResetInterviewAnswerUi();
            _saveTranscriptEnabled = true;

            StatusTextBlock.Text = reason;
            ShowSessionSetupView();
            ResetCreateSessionDraftForNewFlow();
        }
        catch (Exception ex)
        {
            DesktopLogger.Error($"EndSessionAsync error: {ex}");
            // Ensure we always update UI on the dispatcher.
            await Dispatcher.InvokeAsync(() =>
            {
                ResetSessionAnswerHistoryForInterview();
                ResetAutoAnswerTransientState();
                ResetInterviewAnswerUi();
                StatusTextBlock.Text = $"End session error: {ex.Message}";
                ShowSessionSetupView();
                ResetCreateSessionDraftForNewFlow();
            });
        }
    }

    private async Task<bool> EndCallSessionOnServerAsync()
    {
        try
        {
            if (_callSessionId == Guid.Empty) return false;
            using var res = await _apiClient.PostAsync($"callsessions/{_callSessionId}/end", content: null);
            if (!res.IsSuccessStatusCode)
            {
                DesktopLogger.Warn($"Call session end failed status={(int)res.StatusCode}");
                return false;
            }
            return true;
        }
        catch (Exception ex)
        {
            DesktopLogger.Warn($"Call session end failed: {ex.Message}");
            return false;
        }
    }

    private async Task EnsureSpeechConfiguredAsync()
    {
        if (_speechInitInFlight)
            return;

        var hasMic = _speechRecognizer != null && _isListening;
        var hasSystem = _systemSpeechRecognizer != null;
        if (hasMic && hasSystem)
            return;

        _speechInitInFlight = true;
        try
        {
            for (int attempt = 1; attempt <= 3; attempt++)
            {
                var ok = await InitializeSpeechAsync().ConfigureAwait(false);
                if (ok)
                    return;

                await Task.Delay(220).ConfigureAwait(false);
            }

            await Dispatcher.InvokeAsync(() =>
            {
                StatusTextBlock.Text = "Speech setup failed. Check Azure Speech key/region or server token endpoint.";
            });
        }
        finally
        {
            _speechInitInFlight = false;
        }
    }

    private async Task<bool> InitializeSpeechAsync()
    {
        try
        {
            if (_speechRecognizer != null || _systemSpeechRecognizer != null || _loopbackCapture != null)
                await StopSpeechSessionAsync().ConfigureAwait(false);

            var s = App.Settings.AzureSpeech;
            SpeechConfig config;
            if (!string.IsNullOrWhiteSpace(s.Key) && !string.IsNullOrWhiteSpace(s.Region))
            {
                config = SpeechConfig.FromSubscription(s.Key, s.Region);
            }
            else
            {
                var tokenInfo = await GetServerSpeechTokenAsync().ConfigureAwait(false);
                if (tokenInfo == null || string.IsNullOrWhiteSpace(tokenInfo.Token) || string.IsNullOrWhiteSpace(tokenInfo.Region))
                {
                    await Dispatcher.InvokeAsync(() =>
                    {
                        StatusTextBlock.Text = "Speech is not configured (local or server).";
                    });
                    return false;
                }
                config = SpeechConfig.FromAuthorizationToken(tokenInfo.Token, tokenInfo.Region);
            }
            config.SetProperty("SPEECH-EndpointSilenceTimeoutMs", s.EndpointSilenceTimeoutMs);
            config.SpeechRecognitionLanguage = MapSessionLanguageToAzureSpeechLocale(App.Settings.SessionLanguage);

            // 1) Microphone recognizer
            var micAudioConfig = AudioConfig.FromDefaultMicrophoneInput();
            _speechRecognizer = new SpeechRecognizer(config, micAudioConfig);

            _speechRecognizer.Recognizing += (sender, e) =>
            {
                Dispatcher.Invoke(() =>
                {
                    StatusTextBlock.Text = "Listening (mic)...";
                    _partialMic = e.Result.Text ?? string.Empty;
                    UpdateTranscriptDisplay();
                });
            };

            _speechRecognizer.Recognized += (sender, e) =>
            {
                Dispatcher.Invoke(() =>
                {
                    if (e.Result.Reason == ResultReason.RecognizedSpeech)
                    {
                        var text = (e.Result.Text ?? string.Empty).Trim();
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            AppendTranscriptPlain(text);
                            if (_saveTranscriptEnabled) _ = LogMessageAsync("User", text);
                            ScheduleMicAutoAnswer(text);
                        }
                        _partialMic = string.Empty;
                        UpdateTranscriptDisplay();
                    }
                });
            };

            _speechRecognizer.Canceled += (sender, e) =>
            {
                Dispatcher.Invoke(() =>
                {
                    StatusTextBlock.Text = $"Speech canceled: {e.Reason}";
                });
            };

            _speechRecognizer.SessionStopped += (sender, e) =>
            {
                Dispatcher.Invoke(() =>
                {
                    StatusTextBlock.Text = "Speech session stopped.";
                    _isListening = false;
                });
            };

            await _speechRecognizer.StartContinuousRecognitionAsync().ConfigureAwait(false);
            _isListening = true;

            await Dispatcher.InvokeAsync(() =>
            {
                StatusTextBlock.Text = "Listening...";
            });

            // 2) System audio (loopback) recognizer
            await InitializeSystemAudioSpeechAsync(config).ConfigureAwait(false);
            return _speechRecognizer != null && _isListening;
        }
        catch (Exception ex)
        {
            await Dispatcher.InvokeAsync(() =>
            {
                StatusTextBlock.Text = $"Speech init error: {ex.Message}";
            });
            DesktopLogger.Warn($"InitializeSpeechAsync failed: {ex.Message}");
            return false;
        }
    }

    private async Task LoadResumeContextAsync(Guid resumeId)
    {
        try
        {
            DesktopLogger.Info($"Loading resume text for {resumeId}");
            var response = await _apiClient.GetAsync($"resumes/{resumeId}/text");
            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync();
                DesktopLogger.Warn($"GET resumes/{resumeId}/text failed status={(int)response.StatusCode} {response.ReasonPhrase} body={body}");
                Dispatcher.Invoke(() =>
                {
                    StatusTextBlock.Text = $"Resume context load failed: {(int)response.StatusCode} {response.ReasonPhrase}";
                });
                return;
            }

            var text = await response.Content.ReadAsStringAsync();
            _resumeContext = string.IsNullOrWhiteSpace(text) ? null : text;
            DesktopLogger.Info($"Resume context loaded, length={_resumeContext?.Length ?? 0}");
            Dispatcher.Invoke(() =>
            {
                StatusTextBlock.Text = $"Logging → {_apiClient.BaseAddress} | Session {_callSessionId}" + (App.LaunchedViaProtocol ? " (protocol)" : "") +
                    Environment.NewLine +
                    $"Resume → {resumeId} (loaded)" +
                    Environment.NewLine +
                    $"Log → {DesktopLogger.LogFilePath}" +
                    Environment.NewLine +
                    $"Fallback Log → {DesktopLogger.FallbackLogFilePath}";
            });
        }
        catch (Exception ex)
        {
            DesktopLogger.Error($"LoadResumeContextAsync exception: {ex}");
            Dispatcher.Invoke(() =>
            {
                StatusTextBlock.Text = $"Resume context error: {ex.Message}";
            });
        }
    }

    private async Task InitializeSystemAudioSpeechAsync(SpeechConfig baseConfig)
    {
        try
        {
            // Azure Speech expects PCM input. We'll capture loopback audio via NAudio, convert to 16kHz mono 16-bit PCM,
            // and push into a PushAudioInputStream.
            var format = AudioStreamFormat.GetWaveFormatPCM(16000, 16, 1);
            _systemPushStream = AudioInputStream.CreatePushStream(format);
            var sysAudioConfig = AudioConfig.FromStreamInput(_systemPushStream);
            _systemSpeechRecognizer = new SpeechRecognizer(baseConfig, sysAudioConfig);

            _systemSpeechRecognizer.Recognizing += (sender, e) =>
            {
                Dispatcher.Invoke(() =>
                {
                    _partialSystem = e.Result.Text ?? string.Empty;
                    UpdateTranscriptDisplay();
                });
            };

            _systemSpeechRecognizer.Recognized += (sender, e) =>
            {
                Dispatcher.Invoke(() =>
                {
                    if (e.Result.Reason == ResultReason.RecognizedSpeech)
                    {
                        var text = (e.Result.Text ?? string.Empty).Trim();
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            AppendTranscriptPlain(text);
                            if (_saveTranscriptEnabled) _ = LogMessageAsync("Interviewer", text);
                            ScheduleInterviewerAutoAnswer(text);
                        }
                        _partialSystem = string.Empty;
                        UpdateTranscriptDisplay();
                    }
                });
            };

            _systemSpeechRecognizer.Canceled += (sender, e) =>
            {
                Dispatcher.Invoke(() =>
                {
                    StatusTextBlock.Text = $"System audio canceled: {e.Reason}";
                });
            };

            await _systemSpeechRecognizer.StartContinuousRecognitionAsync().ConfigureAwait(false);

            StartLoopbackCapture();
        }
        catch (Exception ex)
        {
            Dispatcher.Invoke(() =>
            {
                StatusTextBlock.Text = $"System audio init error: {ex.Message}";
            });
        }
    }

    private void StartLoopbackCapture()
    {
        _loopbackCapture = new WasapiLoopbackCapture();
        var captureFormat = _loopbackCapture.WaveFormat; // usually 32-bit float stereo, 44.1 or 48 kHz

        _loopbackCapture.DataAvailable += (_, e) =>
        {
            if (_systemPushStream == null || e.BytesRecorded == 0) return;

            // Convert capture format -> 16 kHz mono 16-bit PCM and push to Azure
            byte[]? pcm16 = ConvertTo16KHzMono16Bit(captureFormat, e.Buffer, e.BytesRecorded);
            if (pcm16 != null && pcm16.Length > 0)
            {
                lock (_systemPushLock)
                {
                    try { _systemPushStream.Write(pcm16, pcm16.Length); } catch { /* ignore */ }
                }
            }
        };

        _loopbackCapture.RecordingStopped += (_, __) => { };

        _loopbackCapture.StartRecording();
    }

    /// <summary>
    /// Converts raw capture bytes (typically 32-bit float stereo 44.1/48 kHz, or 16-bit PCM) to 16 kHz mono 16-bit PCM for Azure Speech.
    /// </summary>
    private static byte[]? ConvertTo16KHzMono16Bit(WaveFormat captureFormat, byte[] captureBytes, int captureLength)
    {
        if (captureLength <= 0) return null;

        int sampleRate = captureFormat.SampleRate;
        int channels = captureFormat.Channels;
        bool isFloat = captureFormat.Encoding == WaveFormatEncoding.IeeeFloat && captureFormat.BitsPerSample == 32;
        bool is16Bit = captureFormat.BitsPerSample == 16 && (captureFormat.Encoding == WaveFormatEncoding.Pcm || captureFormat.Encoding == WaveFormatEncoding.Extensible);

        if (!isFloat && !is16Bit) return null;
        if (channels < 1) return null;

        // Downsample to 16 kHz: take every ratio-th frame
        int ratio = Math.Max(1, sampleRate / 16000);
        int bytesPerSample = isFloat ? 4 : 2;
        int frameSize = channels * bytesPerSample;
        int frameCount = captureLength / frameSize;
        int outFrames = (frameCount + ratio - 1) / ratio;
        var outPcm = new byte[outFrames * 2]; // 16-bit = 2 bytes per sample

        for (int i = 0; i < outFrames; i++)
        {
            int srcFrame = i * ratio;
            if (srcFrame >= frameCount) break;

            float mono = 0f;
            for (int c = 0; c < channels; c++)
            {
                int idx = (srcFrame * channels + c) * bytesPerSample;
                if (idx + bytesPerSample > captureLength) continue;
                if (isFloat)
                    mono += BitConverter.ToSingle(captureBytes, idx);
                else
                    mono += BitConverter.ToInt16(captureBytes, idx) / 32768f;
            }
            mono /= channels;
            float clamped = Math.Clamp(mono, -1f, 1f);
            short sample = (short)(clamped * 32767f);
            int outIdx = i * 2;
            outPcm[outIdx] = (byte)(sample & 0xFF);
            outPcm[outIdx + 1] = (byte)((sample >> 8) & 0xFF);
        }

        return outPcm;
    }

    /// <summary>Appends recognized speech to the live transcript UI as plain flowing text (no Me:/Call: labels). Server logs still use roles via <see cref="LogMessageAsync"/>.</summary>
    private void AppendTranscriptPlain(string text)
    {
        var t = (text ?? string.Empty).Trim();
        if (t.Length == 0) return;

        if (_finalTranscript.Length > 0)
            _finalTranscript.Append(' ');
        _finalTranscript.Append(t);

        const int maxChars = 6000;
        if (_finalTranscript.Length > maxChars)
        {
            var trimmed = _finalTranscript.ToString()[^maxChars..].TrimStart();
            _finalTranscript.Clear();
            _finalTranscript.Append(trimmed);
        }
    }

    private void UpdateTranscriptDisplay()
    {
        var sb = new StringBuilder();
        sb.Append(_finalTranscript);

        void appendSpaceIfNeeded()
        {
            if (sb.Length == 0) return;
            if (char.IsWhiteSpace(sb[^1])) return;
            sb.Append(' ');
        }

        if (!string.IsNullOrWhiteSpace(_partialMic))
        {
            appendSpaceIfNeeded();
            sb.Append(_partialMic.Trim());
        }

        if (!string.IsNullOrWhiteSpace(_partialSystem))
        {
            appendSpaceIfNeeded();
            sb.Append(_partialSystem.Trim());
        }

        TranscriptTextBlock.Text = sb.ToString();
        TranscriptTextBlock.ScrollToEnd();
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        var helper = new System.Windows.Interop.WindowInteropHelper(this);
        var handle = helper.Handle;

        SetWindowDisplayAffinity(handle, WDA_EXCLUDEFROMCAPTURE);

        // Keep window within screen/work area (leave margin so bottom stays visible if window starts below top)
        var workArea = SystemParameters.WorkArea;
        const int heightMargin = 60;
        MaxHeight = Math.Max(200, workArea.Height - heightMargin);
        MaxWidth = workArea.Width;

        // Remove extended styles that can draw a border (e.g. yellow focus/activation border)
        int exStyle = GetWindowLong(handle, GWL_EXSTYLE);
        exStyle &= ~WS_EX_CLIENTEDGE;
        exStyle &= ~WS_EX_STATICEDGE;
        // Hide from Alt+Tab: mark as tool window and ensure it's not treated as an "app window".
        exStyle |= WS_EX_TOOLWINDOW;
        exStyle &= ~WS_EX_APPWINDOW;
        SetWindowLong(handle, GWL_EXSTYLE, exStyle);
        SetWindowPos(handle, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);

        // Remove WS_BORDER from style so no system-drawn border
        int style = GetWindowLong(handle, GWL_STYLE);
        style &= ~WS_BORDER;
        SetWindowLong(handle, GWL_STYLE, style);

        // DWM border: light in both startup and main UI.
        TrySetDwmBorderColor(_mainUiStarted ? 0x00ECE7E5 : 0x00FBF7F6);

        UpdateAiAnswerBodyMaxHeight();
    }

    private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ButtonState == MouseButtonState.Pressed)
        {
            DragMove();
        }
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void HideButton_Click(object sender, RoutedEventArgs e)
    {
        MinimizeToRestoreChip();
    }

    private async void MicToggleButton_Click(object sender, RoutedEventArgs e)
    {
        if (_speechRecognizer == null) return;
        _micOn = !_micOn;
        try
        {
            if (_micOn)
            {
                await _speechRecognizer.StartContinuousRecognitionAsync();
                _isListening = true;
                MicToggleButton.Content = "\uE720";
                MicToggleButton.FontFamily = new FontFamily("Segoe MDL2 Assets");
                MicToggleButton.ToolTip = "Microphone on";
                MicToggleButton.Background = new SolidColorBrush(Color.FromRgb(0x0D, 0x94, 0x88));
                MicToggleButton.Foreground = new SolidColorBrush(Color.FromRgb(0xF0, 0xFD, 0xFA));
                MicToggleButton.BorderBrush = new SolidColorBrush(Color.FromRgb(0x14, 0xB8, 0xA6));
                StatusTextBlock.Text = "Mic on.";
            }
            else
            {
                await _speechRecognizer.StopContinuousRecognitionAsync();
                _isListening = false;
                _partialMic = string.Empty;
                UpdateTranscriptDisplay();
                MicToggleButton.Content = "\uE720";
                MicToggleButton.FontFamily = new FontFamily("Segoe MDL2 Assets");
                MicToggleButton.ToolTip = "Microphone off";
                MicToggleButton.Background = new SolidColorBrush(Color.FromRgb(0x2A, 0x2A, 0x38));
                MicToggleButton.Foreground = new SolidColorBrush(Color.FromRgb(0x94, 0xA3, 0xB8));
                MicToggleButton.BorderBrush = new SolidColorBrush(Color.FromRgb(0x3F, 0x3F, 0x50));
                StatusTextBlock.Text = "Mic off.";
            }
        }
        catch (Exception ex)
        {
            _micOn = !_micOn;
            StatusTextBlock.Text = $"Mic error: {ex.Message}";
        }
    }

    private void SpeakerToggleButton_Click(object sender, RoutedEventArgs e)
    {
        _speakerOn = !_speakerOn;
        try
        {
            if (_speakerOn)
            {
                StartLoopbackCapture();
                SpeakerToggleButton.Content = "\uE767";
                SpeakerToggleButton.FontFamily = new FontFamily("Segoe MDL2 Assets");
                SpeakerToggleButton.ToolTip = "Computer audio (speaker) on";
                SpeakerToggleButton.Background = new SolidColorBrush(Color.FromRgb(0x4F, 0x46, 0xE5));
                SpeakerToggleButton.Foreground = new SolidColorBrush(Color.FromRgb(0xEE, 0xF2, 0xFF));
                SpeakerToggleButton.BorderBrush = new SolidColorBrush(Color.FromRgb(0x81, 0x8C, 0xF8));
                StatusTextBlock.Text = "Speaker (computer audio) on.";
            }
            else
            {
                _loopbackCapture?.StopRecording();
                _loopbackCapture?.Dispose();
                _loopbackCapture = null;
                _partialSystem = string.Empty;
                UpdateTranscriptDisplay();
                SpeakerToggleButton.Content = "\uE74F";
                SpeakerToggleButton.FontFamily = new FontFamily("Segoe MDL2 Assets");
                SpeakerToggleButton.ToolTip = "Computer audio (speaker) off";
                SpeakerToggleButton.Background = new SolidColorBrush(Color.FromRgb(0x2A, 0x2A, 0x38));
                SpeakerToggleButton.Foreground = new SolidColorBrush(Color.FromRgb(0x94, 0xA3, 0xB8));
                SpeakerToggleButton.BorderBrush = new SolidColorBrush(Color.FromRgb(0x3F, 0x3F, 0x50));
                StatusTextBlock.Text = "Speaker off.";
            }
        }
        catch (Exception ex)
        {
            _speakerOn = !_speakerOn;
            StatusTextBlock.Text = $"Speaker error: {ex.Message}";
        }
    }

    private async void AiAnswerButton_Click(object sender, RoutedEventArgs e)
    {
        _ = IncrementAiUsageAsync();
        var transcript = (TranscriptTextBlock.Text ?? "").Trim();
        if (string.IsNullOrWhiteSpace(transcript))
        {
            StatusTextBlock.Text = "Transcript is empty. Speak or paste text first.";
            return;
        }

        var latestQuestion = GetLatestQuestionFromTranscript(transcript);
        if (string.IsNullOrWhiteSpace(latestQuestion))
        {
            StatusTextBlock.Text = "Transcript is empty. Speak or paste text first.";
            return;
        }

        var question = QuestionTextBox.Text?.Trim();
        // Only the latest transcribed question is sent (not earlier questions in the same line).
        string userContent = string.IsNullOrWhiteSpace(question)
            ? $"Answer this question (from voice transcription): {latestQuestion}"
            : $"Answer this question. Context from user: {question}. Latest transcribed question: {latestQuestion}";

        var headingForUi = string.IsNullOrWhiteSpace(question)
            ? SummarizeForAnswerHeading(latestQuestion)
            : question.Trim();
        await GetAnswerAsync(userContent, AnswerFromTranscriptSystemPrompt, headingForUi);
    }

    private async void AskButton_Click(object sender, RoutedEventArgs e)
    {
        _ = IncrementAiUsageAsync();
        var question = QuestionTextBox.Text?.Trim();
        if (string.IsNullOrWhiteSpace(question))
        {
            StatusTextBlock.Text = "Type a question first.";
            return;
        }

        await GetAnswerAsync(question, _systemPrompt, question.Trim());
    }

    /// <summary>Single-line style label for answer header when the “question” is long transcript text.</summary>
    private static string SummarizeForAnswerHeading(string text)
    {
        var t = (text ?? string.Empty).Trim().Replace("\r\n", " ").Replace("\n", " ");
        while (t.Contains("  ", StringComparison.Ordinal))
            t = t.Replace("  ", " ");
        if (t.Length > 220)
            return t[..220].TrimEnd() + "…";
        return t;
    }

    /// <summary>Splits auto-detected transcript into separate question units (primarily on '?').</summary>
    private static List<string> SplitTranscriptIntoQuestionSegments(string text)
    {
        var t = (text ?? string.Empty).Trim();
        var list = new List<string>();
        if (t.Length == 0) return list;

        var parts = t.Split('?');
        for (var i = 0; i < parts.Length - 1; i++)
        {
            var s = parts[i].Trim();
            if (s.Length > 0)
                list.Add(s + "?");
        }

        var tail = parts[^1].Trim();
        if (tail.Length > 0)
        {
            if (IsObviousNonQuestionChatter(tail) && tail.Length < 24 && list.Count > 0)
            {
                // Drop trailing "ok / thanks" after real questions
            }
            else
            {
                list.Add(tail);
            }
        }

        if (list.Count == 0)
            list.Add(t);
        return list;
    }

    /// <summary>Latest transcribed question only (last segment after splitting on '?'), for manual AI Answer.</summary>
    private static string GetLatestQuestionFromTranscript(string transcript)
    {
        var t = (transcript ?? string.Empty).Trim();
        if (t.Length == 0) return string.Empty;
        var segments = SplitTranscriptIntoQuestionSegments(t);
        return segments.Count == 0 ? t : segments[^1].Trim();
    }

    private static string BuildAutoAnswerPromptFromSegments(string prefix, IReadOnlyList<string> segments)
    {
        if (segments.Count == 0)
            return prefix;
        if (segments.Count == 1)
            return $"{prefix}\n\n{segments[0]}";

        var sb = new StringBuilder();
        sb.Append(prefix);
        sb.Append(
            "\n\nMultiple questions were detected in the transcription. Answer each one clearly; start each answer with the matching number (1., 2., 3., …).\n\n");
        for (var i = 0; i < segments.Count; i++)
            sb.AppendLine($"{i + 1}) {segments[i]}");
        return sb.ToString();
    }

    private static string FormatSplitQuestionsForAnswerHeading(IReadOnlyList<string> segments)
    {
        if (segments.Count == 0) return string.Empty;
        if (segments.Count == 1)
            return SummarizeForAnswerHeading(segments[0]);

        var sb = new StringBuilder();
        for (var i = 0; i < segments.Count; i++)
        {
            var line = SummarizeForAnswerHeading(segments[i]);
            if (sb.Length > 0) sb.AppendLine();
            sb.Append($"{i + 1}. {line}");
        }

        return sb.ToString();
    }

    private static void AppendBoldQuestionHeading(FlowDocument doc, string? boldHeading)
    {
        if (string.IsNullOrWhiteSpace(boldHeading)) return;

        var lines = boldHeading.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Split('\n', StringSplitOptions.None);
        var nonEmpty = new List<string>();
        foreach (var raw in lines)
        {
            var s = raw.Trim();
            if (s.Length > 0)
                nonEmpty.Add(s);
        }

        for (var i = 0; i < nonEmpty.Count; i++)
        {
            var bottom = i == nonEmpty.Count - 1 ? 10 : 4;
            doc.Blocks.Add(new Paragraph(new Run(nonEmpty[i])
            {
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White
            })
            {
                Margin = new Thickness(0, 0, 0, bottom)
            });
        }
    }

    private void AutoAnswerToggle_Changed(object sender, RoutedEventArgs e)
    {
        if (AutoAnswerToggle?.IsChecked != true)
            ResetAutoAnswerDebounceBuffersOnly();
    }

    private void ResetAutoAnswerDebounceBuffersOnly()
    {
        _callAutoAnswerBuffer = string.Empty;
        _micAutoAnswerBuffer = string.Empty;
        _callAutoAnswerDebounceTimer?.Stop();
        _micAutoAnswerDebounceTimer?.Stop();
    }

    private void ResetAutoAnswerTransientState()
    {
        ResetAutoAnswerDebounceBuffersOnly();
        _lastAutoAnswerNormKey = string.Empty;
        _lastAutoAnswerUtc = DateTimeOffset.MinValue;
    }

    private void OnCallAutoAnswerDebounceTick()
    {
        _callAutoAnswerDebounceTimer?.Stop();
        var q = _callAutoAnswerBuffer.Trim();
        _callAutoAnswerBuffer = string.Empty;
        if (string.IsNullOrWhiteSpace(q)) return;
        _ = TryRunAutoAnswerFromTranscriptAsync(q, AudioQuestionSource.Interviewer);
    }

    private void OnMicAutoAnswerDebounceTick()
    {
        _micAutoAnswerDebounceTimer?.Stop();
        var q = _micAutoAnswerBuffer.Trim();
        _micAutoAnswerBuffer = string.Empty;
        if (string.IsNullOrWhiteSpace(q)) return;
        _ = TryRunAutoAnswerFromTranscriptAsync(q, AudioQuestionSource.SelfMic);
    }

    private void ScheduleInterviewerAutoAnswer(string fragment)
    {
        if (AutoAnswerToggle?.IsChecked != true) return;
        if (!_sessionActive || !_mainUiStarted) return;

        var piece = (fragment ?? string.Empty).Trim();
        if (piece.Length == 0) return;

        if (LooksLikeInterviewerQuestion(piece))
        {
            _callAutoAnswerBuffer = string.Empty;
            _callAutoAnswerDebounceTimer?.Stop();
            _ = TryRunAutoAnswerFromTranscriptAsync(piece, AudioQuestionSource.Interviewer);
            return;
        }

        _callAutoAnswerBuffer = string.IsNullOrEmpty(_callAutoAnswerBuffer)
            ? piece
            : $"{_callAutoAnswerBuffer} {piece}";
        _callAutoAnswerDebounceTimer?.Stop();
        _callAutoAnswerDebounceTimer?.Start();
    }

    private void ScheduleMicAutoAnswer(string fragment)
    {
        if (AutoAnswerToggle?.IsChecked != true) return;
        if (!_sessionActive || !_mainUiStarted) return;

        var piece = (fragment ?? string.Empty).Trim();
        if (piece.Length == 0) return;

        if (LooksLikeMicQuestion(piece))
        {
            _micAutoAnswerBuffer = string.Empty;
            _micAutoAnswerDebounceTimer?.Stop();
            _ = TryRunAutoAnswerFromTranscriptAsync(piece, AudioQuestionSource.SelfMic);
            return;
        }

        _micAutoAnswerBuffer = string.IsNullOrEmpty(_micAutoAnswerBuffer)
            ? piece
            : $"{_micAutoAnswerBuffer} {piece}";
        _micAutoAnswerDebounceTimer?.Stop();
        _micAutoAnswerDebounceTimer?.Start();
    }

    private static string NormalizeForAutoAnswerDedup(string s)
    {
        var t = (s ?? string.Empty).Trim().ToLowerInvariant();
        while (t.Contains("  ", StringComparison.Ordinal))
            t = t.Replace("  ", " ");
        return t.TrimEnd('.', '!', '?', ',', ';', ':');
    }

    private static bool IsObviousNonQuestionChatter(string t)
    {
        var lower = t.Trim().ToLowerInvariant();
        if (lower.Length <= 22 && !t.Contains('?'))
        {
            if (lower is "ok" or "okay" or "yes" or "yeah" or "yep" or "no" or "nope" or "thank you" or "thanks"
                or "got it" or "sure" or "right" or "alright" or "sounds good" or "mm-hmm" or "uh-huh")
                return true;
        }
        return false;
    }

    private static bool LooksLikeInterviewerQuestion(string text)
    {
        if (string.IsNullOrWhiteSpace(text) || text.Length < 10) return false;
        var t = text.Trim();
        if (IsObviousNonQuestionChatter(t)) return false;

        if (t.Contains('?')) return true;

        return InterviewerQuestionLeadInRegex.IsMatch(t);
    }

    private static bool LooksLikeMicQuestion(string text)
    {
        if (string.IsNullOrWhiteSpace(text) || text.Length < 10) return false;
        var t = text.Trim();
        if (IsObviousNonQuestionChatter(t)) return false;

        if (t.Contains('?')) return true;

        var words = t.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
        if (words < 5) return false;
        if (t.StartsWith("what i ", StringComparison.OrdinalIgnoreCase))
            return false;

        return InterviewerQuestionLeadInRegex.IsMatch(t);
    }

    private async Task TryRunAutoAnswerFromTranscriptAsync(string questionText, AudioQuestionSource source)
    {
        try
        {
            if (AutoAnswerToggle?.IsChecked != true) return;
            if (!_sessionActive || !_mainUiStarted) return;
            if (_answerGenerationInFlight) return;

            var trimmed = questionText.Trim();
            if (trimmed.Length < 12) return;

            var wordCount = trimmed.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
            if (wordCount < 3) return;

            var looks = source == AudioQuestionSource.Interviewer
                ? LooksLikeInterviewerQuestion(trimmed)
                : LooksLikeMicQuestion(trimmed);
            if (!looks) return;

            var norm = NormalizeForAutoAnswerDedup(trimmed);
            if (norm.Length < 8) return;

            var now = DateTimeOffset.UtcNow;
            if (string.Equals(norm, _lastAutoAnswerNormKey, StringComparison.Ordinal)
                && (now - _lastAutoAnswerUtc).TotalSeconds < 40)
                return;
            if ((now - _lastAutoAnswerUtc).TotalSeconds < 1.15)
                return;

            _lastAutoAnswerNormKey = norm;
            _lastAutoAnswerUtc = now;

            DesktopLogger.Info($"Auto answer ({source}) len={trimmed.Length}");
            StatusTextBlock.Text = source == AudioQuestionSource.Interviewer
                ? "Auto answer: interviewer question detected..."
                : "Auto answer: question detected on mic...";

            await IncrementAiUsageAsync();

            var prefix = source == AudioQuestionSource.Interviewer
                ? "Answer this question (transcribed from the interviewer's audio). Focus only on the question(s) below:"
                : "Answer this question (transcribed from your microphone). Focus only on the question(s) below:";
            var segments = SplitTranscriptIntoQuestionSegments(trimmed);
            var userContent = BuildAutoAnswerPromptFromSegments(prefix, segments);
            var headingForUi = FormatSplitQuestionsForAnswerHeading(segments);

            await GetAnswerAsync(userContent, AnswerFromTranscriptSystemPrompt, headingForUi);
        }
        catch (Exception ex)
        {
            DesktopLogger.Warn($"TryRunAutoAnswerFromTranscriptAsync: {ex.Message}");
        }
    }

    private async Task IncrementAiUsageAsync()
    {
        try
        {
            if (_callSessionId == Guid.Empty) return;
            using var res = await _apiClient.PostAsync($"callsessions/{_callSessionId}/ai-usage", content: null);
            if (!res.IsSuccessStatusCode)
                DesktopLogger.Warn($"AI usage increment failed status={(int)res.StatusCode}");
        }
        catch
        {
            // Ignore usage tracking errors; never block the user action.
        }
    }

    private void TranscriptClearButton_Click(object sender, RoutedEventArgs e)
    {
        _finalTranscript.Clear();
        _partialMic = string.Empty;
        _partialSystem = string.Empty;
        TranscriptTextBlock.Text = string.Empty;
        ResetAutoAnswerTransientState();
    }

    private void QuestionClearButton_Click(object sender, RoutedEventArgs e)
    {
        QuestionTextBox.Clear();
    }

    private void CloseAnswerButton_Click(object sender, RoutedEventArgs e)
    {
        e.Handled = true;
        ClearAiAnswer();
        if (FindName("AnswerSectionPanel") is System.Windows.UIElement panel)
        {
            panel.Visibility = Visibility.Collapsed;
            SetAnswerSectionRowHeight(collapsed: true);
        }

        UpdateAiAnswerBodyMaxHeight();
        UpdateAnswerHistoryNav();
    }

    private void BumpAnswerUiEpoch()
    {
        Interlocked.Increment(ref _answerUiEpoch);
        Volatile.Write(ref _answerStreamLeaseEpoch, 0);
    }

    /// <summary>Clears the answer rich text, collapses the answer panel, and drops the in-flight flag (e.g. after session end).</summary>
    private void ResetInterviewAnswerUi()
    {
        _answerGenerationInFlight = false;
        ClearAiAnswer();
        if (FindName("AnswerSectionPanel") is UIElement panel)
        {
            panel.Visibility = Visibility.Collapsed;
            SetAnswerSectionRowHeight(collapsed: true);
        }

        UpdateAiAnswerBodyMaxHeight();
        UpdateAnswerHistoryNav();
    }

    private void SetAnswerSectionRowHeight(bool collapsed)
    {
        if (FindName("MainContentGrid") is System.Windows.Controls.Grid grid
            && FindName("AnswerSectionPanel") is UIElement answerPanel)
        {
            var idx = System.Windows.Controls.Grid.GetRow(answerPanel);
            if (idx >= 0 && idx < grid.RowDefinitions.Count)
            {
                grid.RowDefinitions[idx].Height = collapsed
                ? new System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)
                : new System.Windows.GridLength(1, System.Windows.GridUnitType.Star);
            }
        }
    }

    private string ComposeExtendedSystemPrompt(string? systemPrompt)
    {
        var basePrompt = string.IsNullOrWhiteSpace(systemPrompt) ? _systemPrompt : systemPrompt;
        var outputLanguage = GetOutputLanguageNameForPrompt(App.Settings.SessionLanguage);
        return basePrompt +
               "\n\nYou are role-playing as the job candidate described in the resume. " +
               "Always answer in FIRST PERSON as that candidate. " +
               "Never mention that you are an AI, assistant, or language model. " +
               "Use the resume details (name, experience, skills, education) to answer questions such as 'What is your name?' or 'Introduce yourself' as the candidate." +
               $"\n\nLANGUAGE: Respond in {outputLanguage}. If you include code, keep code keywords/identifiers in their original language (usually English), but explain in {outputLanguage}." +
               "\n\nRESPONSE FORMAT RULES (VERY IMPORTANT):" +
               "\n1) Always answer in clear numbered points (1., 2., 3.) suitable for interview speaking." +
               "\n2) Keep language interview-friendly, concise, and confident." +
               "\n3) If code/query is needed, put it in fenced code blocks using triple backticks and language tag (```sql, ```csharp, ```javascript, etc.)." +
               "\n4) In code/query blocks, add short inline comments to explain key lines so the user can explain each line to the interviewer." +
               "\n5) For non-code answers, still keep point-by-point structure and include a short 'How to say this in interview' line at the end.";
    }

    private async Task GetAnswerAsync(string userContent, string? systemPrompt = null, string? displayQuestionForUi = null)
    {
        var payload = new DesktopAiAnswerRequest
        {
            UserContent = userContent,
            SystemPrompt = ComposeExtendedSystemPrompt(systemPrompt),
            ResumeContext = _resumeContext
        };

        using var request = new HttpRequestMessage(HttpMethod.Post, "desktop/ai/answer-stream")
        {
            Content = JsonContent.Create(payload)
        };

        await ExecuteAnswerStreamCoreAsync(request, displayQuestionForUi, aiChannel: "transcript").ConfigureAwait(false);
    }

    private async Task ExecuteAnswerStreamCoreAsync(
        HttpRequestMessage request,
        string? displayQuestionForUi,
        string connectingMessage = "Connecting to AI...",
        string generatingMessage = "Generating answer...",
        string? aiChannel = null)
    {
        var answerEpochSnap = Volatile.Read(ref _answerUiEpoch);
        var answerStreamSessionId = _callSessionId;

        bool AnswerStreamStillValidForSession() =>
            Volatile.Read(ref _answerUiEpoch) == answerEpochSnap
            && _callSessionId == answerStreamSessionId
            && _sessionActive;

        void ReleaseAnswerStreamLeaseAndChrome()
        {
            if (Volatile.Read(ref _answerStreamLeaseEpoch) != answerEpochSnap)
                return;

            Volatile.Write(ref _answerStreamLeaseEpoch, 0);
            if (!_mainUiStarted)
                return;

            _answerGenerationInFlight = false;
            AiAnswerButton.IsEnabled = true;
            AskButton.IsEnabled = true;
            ScreenshotAiButton.IsEnabled = true;
            UpdateAnswerHistoryNav();
        }

        try
        {
            _currentAnswerDisplayHeading = string.IsNullOrWhiteSpace(displayQuestionForUi)
                ? null
                : displayQuestionForUi.Trim();

            await Dispatcher.InvokeAsync(() =>
            {
                if (!AnswerStreamStillValidForSession())
                    return;

                AiAnswerButton.IsEnabled = false;
                AskButton.IsEnabled = false;
                ScreenshotAiButton.IsEnabled = false;
                StatusTextBlock.Text = connectingMessage;

                if (FindName("AnswerSectionPanel") is System.Windows.UIElement panelEarly)
                {
                    panelEarly.Visibility = Visibility.Visible;
                    SetAnswerSectionRowHeight(collapsed: false);
                }

                _answerGenerationInFlight = true;
                Volatile.Write(ref _answerStreamLeaseEpoch, answerEpochSnap);
                UpdateAnswerHistoryNav();
            });

            if (!AnswerStreamStillValidForSession())
            {
                await Dispatcher.InvokeAsync(ReleaseAnswerStreamLeaseAndChrome);
                return;
            }

            _ = Dispatcher.BeginInvoke(new Action(UpdateAiAnswerBodyMaxHeight), DispatcherPriority.Loaded);
            await Dispatcher.InvokeAsync(() =>
            {
                if (AnswerStreamStillValidForSession())
                    BeginStreamingAnswerDisplay(_currentAnswerDisplayHeading);
            });
            await Dispatcher.InvokeAsync(() =>
            {
                if (AnswerStreamStillValidForSession())
                    StatusTextBlock.Text = generatingMessage;
            });

            if (!AnswerStreamStillValidForSession())
            {
                await Dispatcher.InvokeAsync(ReleaseAnswerStreamLeaseAndChrome);
                return;
            }

            using var response = await _apiClient
                .SendAsync(request, HttpCompletionOption.ResponseHeadersRead)
                .ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                throw new Exception($"HTTP {(int)response.StatusCode} {response.ReasonPhrase}. {body}");
            }

            var completionSb = new StringBuilder();
            await using var respStream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
            using var reader = new StreamReader(respStream);

            while (await reader.ReadLineAsync().ConfigureAwait(false) is { } line)
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                using var jd = JsonDocument.Parse(line);
                var root = jd.RootElement;
                if (root.TryGetProperty("error", out var errProp))
                {
                    var err = errProp.GetString() ?? "Stream error";
                    throw new Exception(err);
                }

                if (!root.TryGetProperty("d", out var dProp))
                    continue;

                var delta = dProp.GetString();
                if (string.IsNullOrEmpty(delta))
                    continue;

                completionSb.Append(delta);

                await Dispatcher.InvokeAsync(() =>
                {
                    if (!AnswerStreamStillValidForSession())
                        return;
                    if (_streamingAnswerRun != null)
                        _streamingAnswerRun.Text += delta;
                }, DispatcherPriority.Background);
            }

            var completion = completionSb.ToString();
            await Dispatcher.InvokeAsync(() =>
            {
                if (!AnswerStreamStillValidForSession())
                    return;

                StatusTextBlock.Text = "Formatting answer...";
                RenderAiAnswer(completion, _currentAnswerDisplayHeading);
                RegisterSuccessfulAnswerInHistory(completion);
                StatusTextBlock.Text = "Answer ready.";
                try
                {
                    AiAnswerTextBlock.ScrollToHome();
                }
                catch
                {
                    // ignore
                }

                UpdateAiAnswerBodyMaxHeight();
                UpdateAnswerHistoryNav();
            });

            if (!string.IsNullOrWhiteSpace(completion) && AnswerStreamStillValidForSession())
            {
                await AttachServerIdToAppendedAssistantAsync(answerStreamSessionId).ConfigureAwait(false);
                if (_callSessionId != Guid.Empty && _callSessionId == answerStreamSessionId)
                {
                    DesktopAnalytics.Track(
                        DesktopAnalyticsEventTypes.AiResponseGenerated,
                        JsonSerializer.Serialize(new
                        {
                            channel = aiChannel ?? "unknown",
                            contentLength = completion.Length,
                        }),
                        _callSessionId);
                }
            }
        }
        catch (Exception ex)
        {
            await Dispatcher.InvokeAsync(() =>
            {
                if (!AnswerStreamStillValidForSession())
                    return;

                RenderAiAnswer($"Error:\n- {ex.Message}", _currentAnswerDisplayHeading);
                if (FindName("AnswerSectionPanel") is System.Windows.UIElement p)
                {
                    p.Visibility = Visibility.Visible;
                    SetAnswerSectionRowHeight(collapsed: false);
                }

                StatusTextBlock.Text = "Error.";
                UpdateAiAnswerBodyMaxHeight();
                UpdateAnswerHistoryNav();
            });
        }
        finally
        {
            await Dispatcher.InvokeAsync(ReleaseAnswerStreamLeaseAndChrome);
        }
    }

    private async void ScreenshotAiButton_Click(object sender, RoutedEventArgs e)
    {
        _ = IncrementAiUsageAsync();
        await Dispatcher.InvokeAsync(() => { StatusTextBlock.Text = "Capturing screen..."; });

        byte[] pngBytes;
        try
        {
            pngBytes = await Task.Run(() => ScreenCaptureHelper.CaptureVirtualScreenToPngBytes()).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            await Dispatcher.InvokeAsync(() =>
            {
                StatusTextBlock.Text = $"Screen capture failed: {ex.Message}";
            });
            return;
        }

        var payload = new DesktopScreenshotAnswerRequest
        {
            ImageBase64 = Convert.ToBase64String(pngBytes),
            MimeType = "image/png",
            SystemPrompt = ComposeExtendedSystemPrompt(_systemPrompt),
            ResumeContext = _resumeContext
        };

        using var request = new HttpRequestMessage(HttpMethod.Post, "desktop/ai/screenshot-answer-stream")
        {
            Content = JsonContent.Create(payload)
        };

        var cid = _callSessionId == Guid.Empty ? (Guid?)null : _callSessionId;
        DesktopAnalytics.Track(DesktopAnalyticsEventTypes.AnalyzeScreenRequested, null, cid);

        await ExecuteAnswerStreamCoreAsync(
            request,
            "Screenshot",
            connectingMessage: "Sending screenshot…",
            generatingMessage: "Reading screen & generating answer…",
            aiChannel: "screenshot").ConfigureAwait(false);
    }

    private void ClearAiAnswer()
    {
        _streamingAnswerRun = null;
        _currentAnswerDisplayHeading = null;
        AiAnswerTextBlock.Document = new FlowDocument(new Paragraph());
    }

    private void BeginStreamingAnswerDisplay(string? boldHeading)
    {
        _streamingAnswerRun = new Run(string.Empty) { Foreground = Brushes.White };
        var doc = new FlowDocument
        {
            PagePadding = new Thickness(0),
            LineHeight = 20
        };
        AppendBoldQuestionHeading(doc, boldHeading);

        var p = new Paragraph { Margin = new Thickness(0, 0, 0, 4) };
        p.Inlines.Add(_streamingAnswerRun);
        doc.Blocks.Add(p);
        AiAnswerTextBlock.Document = doc;
        Dispatcher.BeginInvoke(new Action(() =>
        {
            ScrollViewer.SetCanContentScroll(AiAnswerTextBlock, false);
            ScrollViewer.SetPanningMode(AiAnswerTextBlock, PanningMode.VerticalOnly);
            UpdateAiAnswerBodyMaxHeight();
        }), DispatcherPriority.Loaded);
    }

    private void RenderAiAnswer(string content, string? boldHeading = null)
    {
        _streamingAnswerRun = null;
        var doc = new FlowDocument
        {
            PagePadding = new Thickness(0),
            LineHeight = 20
        };

        AppendBoldQuestionHeading(doc, boldHeading);

        var normalized = (content ?? string.Empty).Replace("\r\n", "\n");
        var lines = normalized.Split('\n');
        var codeBuffer = new StringBuilder();
        var inCodeBlock = false;

        foreach (var raw in lines)
        {
            var line = raw ?? string.Empty;
            var trimmed = line.Trim();

            if (trimmed.StartsWith("```", StringComparison.Ordinal))
            {
                if (!inCodeBlock)
                {
                    inCodeBlock = true;
                    codeBuffer.Clear();
                }
                else
                {
                    AddCodeBlock(doc, codeBuffer.ToString());
                    inCodeBlock = false;
                }
                continue;
            }

            if (inCodeBlock)
            {
                codeBuffer.AppendLine(line);
                continue;
            }

            if (string.IsNullOrWhiteSpace(line))
            {
                doc.Blocks.Add(new Paragraph(new Run(" ")) { Margin = new Thickness(0, 3, 0, 3) });
                continue;
            }

            string text = line;
            if (Regex.IsMatch(trimmed, @"^\d+[\.\)]\s"))
                text = trimmed;
            else if (trimmed.StartsWith("- "))
                text = "• " + trimmed[2..];

            doc.Blocks.Add(new Paragraph(new Run(text))
            {
                Margin = new Thickness(0, 0, 0, 6),
                Foreground = Brushes.White
            });
        }

        if (inCodeBlock && codeBuffer.Length > 0)
            AddCodeBlock(doc, codeBuffer.ToString());

        AiAnswerTextBlock.Document = doc;
        // Replacing Document can rebuild the template; re-apply smooth (pixel) scroll on the internal ScrollViewer.
        Dispatcher.BeginInvoke(new Action(() =>
        {
            ScrollViewer.SetCanContentScroll(AiAnswerTextBlock, false);
            ScrollViewer.SetPanningMode(AiAnswerTextBlock, PanningMode.VerticalOnly);
            UpdateAiAnswerBodyMaxHeight();
        }), DispatcherPriority.Loaded);
    }

    private static void AddCodeBlock(FlowDocument doc, string codeText)
    {
        var codeDoc = new FlowDocument
        {
            PagePadding = new Thickness(8, 6, 8, 6),
            Background = Brushes.Transparent
        };

        var codeParagraph = new Paragraph { Margin = new Thickness(0) };
        foreach (var line in codeText.Replace("\r\n", "\n").Split('\n'))
        {
            AppendHighlightedCodeLine(codeParagraph, line);
            codeParagraph.Inlines.Add(new LineBreak());
        }
        codeDoc.Blocks.Add(codeParagraph);

        var codeView = new RichTextBox
        {
            IsReadOnly = true,
            IsDocumentEnabled = false,
            BorderThickness = new Thickness(0),
            Background = Brushes.Transparent,
            VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            FontFamily = new FontFamily("Consolas"),
            FontSize = 13,
            Foreground = new SolidColorBrush(Color.FromRgb(0xF8, 0xFA, 0xFC)),
            Document = codeDoc
        };
        ScrollViewer.SetCanContentScroll(codeView, false);
        ScrollViewer.SetPanningMode(codeView, PanningMode.None);

        var border = new Border
        {
            BorderBrush = new SolidColorBrush(Color.FromRgb(0x94, 0xA3, 0xB8)),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(8),
            Background = new SolidColorBrush(Color.FromArgb(0xE6, 0x15, 0x23, 0x42)),
            Margin = new Thickness(0, 4, 0, 10),
            Child = codeView
        };

        doc.Blocks.Add(new BlockUIContainer(border));
    }

    private static void AppendHighlightedCodeLine(Paragraph paragraph, string line)
    {
        var commentIndex = line.IndexOf("//", StringComparison.Ordinal);
        if (commentIndex < 0) commentIndex = line.IndexOf("--", StringComparison.Ordinal);
        if (commentIndex < 0 && line.TrimStart().StartsWith("#", StringComparison.Ordinal)) commentIndex = 0;

        var codePart = commentIndex >= 0 ? line[..commentIndex] : line;
        var commentPart = commentIndex >= 0 ? line[commentIndex..] : string.Empty;

        foreach (var token in Regex.Split(codePart, @"(\W+)"))
        {
            if (string.IsNullOrEmpty(token)) continue;
            var run = new Run(token);
            if (CodeKeywords.Contains(token))
            {
                run.Foreground = new SolidColorBrush(Color.FromRgb(0x7D, 0xDD, 0xFF));
                run.FontWeight = FontWeights.SemiBold;
            }
            else if (Regex.IsMatch(token, "^\".*\"$|^'.*'$"))
            {
                run.Foreground = new SolidColorBrush(Color.FromRgb(0xFC, 0xA5, 0xA5));
            }
            else if (Regex.IsMatch(token, @"^\d+$"))
            {
                run.Foreground = new SolidColorBrush(Color.FromRgb(0xFD, 0xE0, 0x68));
            }
            else
            {
                run.Foreground = new SolidColorBrush(Color.FromRgb(0xF1, 0xF5, 0xF9));
            }
            paragraph.Inlines.Add(run);
        }

        if (!string.IsNullOrWhiteSpace(commentPart))
        {
            paragraph.Inlines.Add(new Run(commentPart)
            {
                Foreground = new SolidColorBrush(Color.FromRgb(0x6E, 0xEE, 0xB3))
            });
        }
    }

    private async Task<DesktopSpeechTokenResponse?> GetServerSpeechTokenAsync()
    {
        try
        {
            using var res = await _apiClient.GetAsync("desktop/speech/token");
            if (!res.IsSuccessStatusCode)
            {
                var body = await res.Content.ReadAsStringAsync();
                DesktopLogger.Warn($"GET desktop/speech/token failed status={(int)res.StatusCode} {res.ReasonPhrase} body={body}");
                return null;
            }
            return await res.Content.ReadFromJsonAsync<DesktopSpeechTokenResponse>(
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        }
        catch (Exception ex)
        {
            DesktopLogger.Warn($"GetServerSpeechTokenAsync failed: {ex.Message}");
            return null;
        }
    }

    private async Task<Guid?> LogMessageAsync(string role, string content)
    {
        try
        {
            if (_callSessionId == Guid.Empty) return null;
            if (string.IsNullOrWhiteSpace(content)) return null;

            var payload = new { role, content };
            DesktopLogger.Info($"POST callsessions/{_callSessionId}/messages role={role} len={content.Length}");
            var response = await _apiClient.PostAsJsonAsync($"callsessions/{_callSessionId}/messages", payload);
            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync();
                DesktopLogger.Warn($"POST failed status={(int)response.StatusCode} {response.ReasonPhrase} body={body}");
                throw new Exception($"HTTP {(int)response.StatusCode} {response.ReasonPhrase}. {body}");
            }

            var dto = await response.Content.ReadFromJsonAsync<CallSessionMessageDto>(
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            DesktopLogger.Info("POST ok");
            return dto?.Id;
        }
        catch (Exception ex)
        {
            Dispatcher.Invoke(() =>
            {
                var msg = ex.Message;
                if (ex is HttpRequestException && ex.InnerException != null)
                    msg = ex.InnerException.Message;
                StatusTextBlock.Text = $"Log failed: {msg}";
            });
            DesktopLogger.Error($"LogMessageAsync exception: {ex}");
            return null;
        }
    }

    protected override async void OnClosed(EventArgs e)
    {
        lock (_loginSync)
        {
            try { _activeLoginCts?.Cancel(); } catch { /* ignore */ }
            _activeLoginCts?.Dispose();
            _activeLoginCts = null;
        }

        if (_restoreChip != null)
        {
            _restoreChip.RestoreRequested -= OnRestoreChipRestoreRequested;
            try { _restoreChip.Close(); } catch { /* ignore */ }
            _restoreChip = null;
        }

        // Match EndSessionAsync: notify the server so the session is marked ended and AI notes can run
        // (SaveTranscript on the API). Otherwise closing via ✕ leaves an "Active" session with no notes.
        var cid = _callSessionId;
        var shouldPostEnd = cid != Guid.Empty && (_sessionActive || _pendingEndSync);

        BumpAnswerUiEpoch();
        _sessionActive = false;
        try { _sessionTimer?.Stop(); } catch { /* ignore */ }

        if (shouldPostEnd)
        {
            try
            {
                if (await EndCallSessionOnServerAsync().ConfigureAwait(false))
                {
                    _pendingEndSync = false;
                    _callSessionId = Guid.Empty;
                }
            }
            catch (Exception ex)
            {
                DesktopLogger.Warn($"Call session end on shutdown: {ex.Message}");
            }
        }

        await StopSpeechSessionAsync();
        _loginFlowLock.Dispose();
        _saveTranscriptEnabled = true;

        base.OnClosed(e);
    }
}