using System;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;

namespace AiInterviewAssistant;

public partial class MainWindow : Window
{
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

    private const int WDA_NONE = 0;
    private const int WDA_EXCLUDEFROMCAPTURE = 0x11;

    [DllImport("user32.dll")]
    private static extern bool SetWindowDisplayAffinity(IntPtr hWnd, int affinity);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    private const int GWL_EXSTYLE = -20;
    private const int GWL_STYLE = -16;
    private const int WS_EX_CLIENTEDGE = 0x200;
    private const int WS_EX_STATICEDGE = 0x20000;
    private const int WS_BORDER = 0x00800000;

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

    private sealed class DesktopAiAnswerResponse
    {
        public string Answer { get; set; } = string.Empty;
    }

    private sealed class DesktopSpeechTokenResponse
    {
        public string Region { get; set; } = string.Empty;
        public string Token { get; set; } = string.Empty;
        public int ExpiresInSeconds { get; set; }
    }

    private readonly HttpClient _apiClient;
    private Guid _callSessionId;
    private readonly Guid? _resumeId;
    private string? _resumeContext;
    private bool _mainUiStarted;
    private RestoreChipWindow? _restoreChip;
    private bool _startupPositionInitialized;
    private readonly DesktopWebAuthService _desktopWebAuthService;
    private readonly SemaphoreSlim _loginFlowLock = new(1, 1);
    private readonly object _loginSync = new();
    private CancellationTokenSource? _activeLoginCts;
    private int _loginAttemptVersion;
    private bool _isFreeSessionFlow = true;

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
        if (!Guid.TryParse(settings.CallSessionId, out _callSessionId))
        {
            // If parsing fails, use the fixed session id you requested for now
            _callSessionId = Guid.Parse("AB589C99-0980-4467-8AF5-ADAB340FE1A0");
        }

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

        // Startup screen only; main interview UI appears after Login.
        ApplyStartupChrome();
        StartupLoginView.LoginRequested += async (_, _) => await HandleLoginRequestedAsync();
        StartupLoginView.DashboardRequested += (_, _) => OpenWebDashboardHome();
        StartupLoginView.CloseRequested += (_, _) => Close();
        StartupLoginView.MinimizeRequested += (_, _) => MinimizeToRestoreChip();
        StartupLoginView.WindowSlotRequested += ApplyStartupWindowSlot;
        SessionSetupView.FullSessionRequested += (_, _) =>
        {
            _isFreeSessionFlow = false;
            ShowCreateSessionDetailsView();
        };
        SessionSetupView.FreeSessionRequested += (_, _) =>
        {
            _isFreeSessionFlow = true;
            ShowCreateSessionDetailsView();
        };
        SessionSetupView.CloseRequested += (_, _) => Close();
        SessionSetupView.MinimizeRequested += (_, _) => MinimizeToRestoreChip();
        SessionSetupView.WindowSlotRequested += ApplyStartupWindowSlot;
        SessionSetupView.PastSessionsRequested += (_, _) => ShowPastSessionsView();
        PastSessionsView.CreateRequested += (_, _) => ShowSessionSetupView();
        PastSessionsView.ViewAllRequested += (_, _) => OpenPastSessionsInBrowser();
        PastSessionsView.ActivateNotActivatedRequested += (sessionId, isFree) =>
        {
            _callSessionId = sessionId;
            _isFreeSessionFlow = isFree;
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
        ActivateSessionView.ActivateRequested += (_, _) => StartInterviewSession();
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
    }

    /// <summary>
    /// Hides the main window and shows a small green circle; click the circle to restore (login / app UI).
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

    private void OnRestoreChipRestoreRequested(object? sender, EventArgs e)
    {
        Show();
        if (WindowState == WindowState.Minimized)
            WindowState = WindowState.Normal;
        Activate();
        try { Focus(); } catch { /* ignore */ }
        _restoreChip?.Hide();
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

            DesktopAuthTokens tokens = await _desktopWebAuthService.AuthenticateAsync(attemptCts.Token, authTraceId);
            _apiClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", tokens.AccessToken);
            DesktopLogger.Info($"[AUTH:{authTraceId}] Desktop login successful. API bearer updated accessLen={tokens.AccessToken.Length} refreshLen={tokens.RefreshToken.Length}");
            StartupLoginView.SetLoginStatus("Login successful.");
            ShowSessionSetupView();
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

            _finalTranscript.Clear();
            _partialMic = string.Empty;
            _partialSystem = string.Empty;
            TranscriptTextBlock.Text = string.Empty;

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
        Width = 560;
        RootChromeBorder.Background = new SolidColorBrush(Color.FromArgb(0x99, 0x1E, 0x1E, 0x1E));
        RootChromeBorder.Padding = new Thickness(12);
        RootChromeBorder.CornerRadius = new CornerRadius(12);
        TrySetDwmBorderColor(0x001E1E1E);
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

    private void StartMainUiIfNeeded()
    {
        if (_mainUiStarted) return;
        _mainUiStarted = true;
        ApplyMainChrome();
        InitializeSpeechAsync().ConfigureAwait(false);
        if (_resumeId.HasValue)
        {
            _ = LoadResumeContextAsync(_resumeId.Value);
        }
    }

    private void StartInterviewSession()
    {
        StartMainUiIfNeeded();
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
        _sessionActive = false;

        try
        {
            _sessionTimer?.Stop();
            EndSessionButton.IsEnabled = false;

            _pendingEndSync = true;
            _lastServerSyncAttemptUtc = DateTimeOffset.MinValue;
            // IMPORTANT: don't use ConfigureAwait(false) here because this method updates WPF UI.
            _pendingEndSync = !await EndCallSessionOnServerAsync();
            await StopSpeechSessionAsync();
            _mainUiStarted = false;

            _finalTranscript.Clear();
            _partialMic = string.Empty;
            _partialSystem = string.Empty;
            TranscriptTextBlock.Text = string.Empty;

            StatusTextBlock.Text = reason;
            ShowSessionSetupView();
        }
        catch (Exception ex)
        {
            DesktopLogger.Error($"EndSessionAsync error: {ex}");
            // Ensure we always update UI on the dispatcher.
            await Dispatcher.InvokeAsync(() =>
            {
                StatusTextBlock.Text = $"End session error: {ex.Message}";
                ShowSessionSetupView();
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

    private async Task InitializeSpeechAsync()
    {
        try
        {
            var s = App.Settings.AzureSpeech;
            SpeechConfig config;
            if (!string.IsNullOrWhiteSpace(s.Key) && !string.IsNullOrWhiteSpace(s.Region))
            {
                config = SpeechConfig.FromSubscription(s.Key, s.Region);
            }
            else
            {
                var tokenInfo = await GetServerSpeechTokenAsync();
                if (tokenInfo == null || string.IsNullOrWhiteSpace(tokenInfo.Token) || string.IsNullOrWhiteSpace(tokenInfo.Region))
                {
                    Dispatcher.Invoke(() =>
                    {
                        StatusTextBlock.Text = "Speech is not configured (local or server).";
                    });
                    return;
                }
                config = SpeechConfig.FromAuthorizationToken(tokenInfo.Token, tokenInfo.Region);
            }
            config.SetProperty("SPEECH-EndpointSilenceTimeoutMs", s.EndpointSilenceTimeoutMs);
            config.SpeechRecognitionLanguage = "en-US";

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
                            AppendFinalLine($"Me: {text}");
                            _ = LogMessageAsync("User", text);
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

            Dispatcher.Invoke(() =>
            {
                StatusTextBlock.Text = "Listening...";
            });

            // 2) System audio (loopback) recognizer
            await InitializeSystemAudioSpeechAsync(config).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            Dispatcher.Invoke(() =>
            {
                StatusTextBlock.Text = $"Speech init error: {ex.Message}";
            });
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
                            AppendFinalLine($"Call: {text}");
                            _ = LogMessageAsync("System", text);
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

    private void AppendFinalLine(string line)
    {
        if (_finalTranscript.Length > 0)
            _finalTranscript.AppendLine();

        _finalTranscript.Append(line);

        // Keep transcript from growing without bound in UI
        const int maxChars = 6000;
        if (_finalTranscript.Length > maxChars)
        {
            var trimmed = _finalTranscript.ToString();
            trimmed = trimmed[^maxChars..];
            _finalTranscript.Clear();
            _finalTranscript.Append(trimmed);
        }
    }

    private void UpdateTranscriptDisplay()
    {
        // Show final lines plus latest partial from mic/system (if any)
        var sb = new StringBuilder();
        sb.Append(_finalTranscript);

        var partials = new StringBuilder();
        if (!string.IsNullOrWhiteSpace(_partialMic))
            partials.Append($"Me (partial): {_partialMic}");
        if (!string.IsNullOrWhiteSpace(_partialSystem))
        {
            if (partials.Length > 0) partials.AppendLine();
            partials.Append($"Call (partial): {_partialSystem}");
        }

        if (partials.Length > 0)
        {
            if (sb.Length > 0) sb.AppendLine().AppendLine();
            sb.Append(partials);
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
        SetWindowLong(handle, GWL_EXSTYLE, exStyle);

        // Remove WS_BORDER from style so no system-drawn border
        int style = GetWindowLong(handle, GWL_STYLE);
        style &= ~WS_BORDER;
        SetWindowLong(handle, GWL_STYLE, style);

        // DWM border: light on startup screen, dark after login (also updated in ApplyMainChrome).
        TrySetDwmBorderColor(_mainUiStarted ? 0x001E1E1E : 0x00FBF7F6);
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
        WindowState = WindowState.Minimized;
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
                MicToggleButton.Content = "Mic On";
                MicToggleButton.Background = new SolidColorBrush(Color.FromArgb(0xAA, 0x53, 0x5A, 0x65));
                StatusTextBlock.Text = "Mic on.";
            }
            else
            {
                await _speechRecognizer.StopContinuousRecognitionAsync();
                _isListening = false;
                _partialMic = string.Empty;
                UpdateTranscriptDisplay();
                MicToggleButton.Content = "Mic Off";
                MicToggleButton.Background = new SolidColorBrush(Color.FromArgb(0xAA, 0x39, 0x39, 0x39));
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
                SpeakerToggleButton.Content = "Speaker On";
                SpeakerToggleButton.Background = new SolidColorBrush(Color.FromArgb(0xAA, 0x53, 0x5A, 0x65));
                StatusTextBlock.Text = "Speaker (computer audio) on.";
            }
            else
            {
                _loopbackCapture?.StopRecording();
                _loopbackCapture?.Dispose();
                _loopbackCapture = null;
                _partialSystem = string.Empty;
                UpdateTranscriptDisplay();
                SpeakerToggleButton.Content = "Speaker Off";
                SpeakerToggleButton.Background = new SolidColorBrush(Color.FromArgb(0xAA, 0x39, 0x39, 0x39));
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

        var question = QuestionTextBox.Text?.Trim();
        // Treat transcript as the user's question; optional typed question adds context
        string userContent = string.IsNullOrWhiteSpace(question)
            ? $"Answer this question (from voice transcription): {transcript}"
            : $"Answer this question. Context from user: {question}. What they said: {transcript}";

        const string answerFromTranscriptPrompt =
            "You are role-playing as the job candidate whose resume is provided in the context. " +
            "Always answer in FIRST PERSON as that candidate (for example: 'My name is ...', 'I have 5 years of experience ...'). " +
            "Never say you are an AI or language model. " +
            "The user will give you text that was transcribed from speech. " +
            "Answer the question fully and in detail, using the resume details whenever relevant. " +
            "For questions like 'what is your name' or 'introduce yourself', answer using the candidate's real name and background from the resume.";
        await GetAnswerAsync(userContent, answerFromTranscriptPrompt);
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

        await GetAnswerAsync(question, _systemPrompt);
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
    }

    private void QuestionClearButton_Click(object sender, RoutedEventArgs e)
    {
        QuestionTextBox.Clear();
    }

    private void CloseAnswerButton_Click(object sender, RoutedEventArgs e)
    {
        e.Handled = true;
        AiAnswerTextBlock.Text = string.Empty;
        if (FindName("AnswerSectionPanel") is System.Windows.UIElement panel)
        {
            panel.Visibility = Visibility.Collapsed;
            SetAnswerSectionRowHeight(collapsed: true);
        }
    }

    private void SetAnswerSectionRowHeight(bool collapsed)
    {
        if (FindName("MainContentGrid") is System.Windows.Controls.Grid grid && grid.RowDefinitions.Count > 3)
        {
            grid.RowDefinitions[3].Height = collapsed
                ? new System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)
                : new System.Windows.GridLength(1, System.Windows.GridUnitType.Star);
        }
    }

    private async Task GetAnswerAsync(string userContent, string? systemPrompt = null)
    {
        try
        {
            AiAnswerButton.IsEnabled = false;
            AskButton.IsEnabled = false;
            StatusTextBlock.Text = "Thinking (server AI)...";

            var basePrompt = string.IsNullOrWhiteSpace(systemPrompt) ? _systemPrompt : systemPrompt;
            // Ensure the model always speaks as the candidate, not as an AI assistant.
            basePrompt +=
                "\n\nYou are role-playing as the job candidate described in the resume. " +
                "Always answer in FIRST PERSON as that candidate. " +
                "Never mention that you are an AI, assistant, or language model. " +
                "Use the resume details (name, experience, skills, education) to answer questions such as 'What is your name?' or 'Introduce yourself' as the candidate.";
            var payload = new DesktopAiAnswerRequest
            {
                UserContent = userContent,
                SystemPrompt = basePrompt,
                ResumeContext = _resumeContext
            };
            var response = await _apiClient.PostAsJsonAsync("desktop/ai/answer", payload);
            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync();
                throw new Exception($"HTTP {(int)response.StatusCode} {response.ReasonPhrase}. {body}");
            }
            var result = await response.Content.ReadFromJsonAsync<DesktopAiAnswerResponse>();
            var completion = result?.Answer ?? string.Empty;

            AiAnswerTextBlock.Text = completion;
            if (FindName("AnswerSectionPanel") is System.Windows.UIElement panel)
            {
                panel.Visibility = Visibility.Visible;
                SetAnswerSectionRowHeight(collapsed: false);
            }
            StatusTextBlock.Text = "Answer ready.";

            // Log AI answer to dashboard API
            if (!string.IsNullOrWhiteSpace(completion))
            {
                _ = LogMessageAsync("Assistant", completion);
            }
        }
        catch (Exception ex)
        {
            AiAnswerTextBlock.Text = $"Error: {ex.Message}";
            if (FindName("AnswerSectionPanel") is System.Windows.UIElement p)
            {
                p.Visibility = Visibility.Visible;
                SetAnswerSectionRowHeight(collapsed: false);
            }
            StatusTextBlock.Text = "Error.";
        }
        finally
        {
            AiAnswerButton.IsEnabled = true;
            AskButton.IsEnabled = true;
        }
    }

    private async Task<DesktopSpeechTokenResponse?> GetServerSpeechTokenAsync()
    {
        try
        {
            using var res = await _apiClient.GetAsync("desktop/speech/token");
            if (!res.IsSuccessStatusCode)
            {
                DesktopLogger.Warn($"GET desktop/speech/token failed status={(int)res.StatusCode}");
                return null;
            }
            return await res.Content.ReadFromJsonAsync<DesktopSpeechTokenResponse>();
        }
        catch (Exception ex)
        {
            DesktopLogger.Warn($"GetServerSpeechTokenAsync failed: {ex.Message}");
            return null;
        }
    }

    private async Task LogMessageAsync(string role, string content)
    {
        try
        {
            if (_callSessionId == Guid.Empty) return;
            if (string.IsNullOrWhiteSpace(content)) return;

            var payload = new { role, content };
            DesktopLogger.Info($"POST callsessions/{_callSessionId}/messages role={role} len={content.Length}");
            var response = await _apiClient.PostAsJsonAsync($"callsessions/{_callSessionId}/messages", payload);
            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync();
                DesktopLogger.Warn($"POST failed status={(int)response.StatusCode} {response.ReasonPhrase} body={body}");
                throw new Exception($"HTTP {(int)response.StatusCode} {response.ReasonPhrase}. {body}");
            }
            DesktopLogger.Info("POST ok");
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

        _sessionActive = false;
        try { _sessionTimer?.Stop(); } catch { /* ignore */ }
        await StopSpeechSessionAsync();
        _loginFlowLock.Dispose();

        base.OnClosed(e);
    }
}