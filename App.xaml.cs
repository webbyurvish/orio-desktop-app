using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;

namespace AiInterviewAssistant;

public partial class App : Application
{
    public static AppSettings Settings { get; private set; } = new();
    public static bool LaunchedViaProtocol { get; private set; }
    public static string? ProtocolUrl { get; private set; }

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        try
        {
            var configPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");
            DesktopLogger.Info($"Startup baseDirectory={AppContext.BaseDirectory}");
            DesktopLogger.Info($"Startup configPath={configPath} exists={File.Exists(configPath)}");
            DesktopLogger.Info($"StartupArgs (WPF) = {(e.Args == null ? "<null>" : string.Join(" | ", e.Args))}");
            DesktopLogger.Info($"CommandLineArgs = {string.Join(" | ", Environment.GetCommandLineArgs())}");
            if (File.Exists(configPath))
            {
                var json = File.ReadAllText(configPath);
                // appsettings.json may contain // comments; default System.Text.Json rejects them and would fall back to localhost defaults.
                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true,
                    ReadCommentHandling = JsonCommentHandling.Skip,
                    AllowTrailingCommas = true
                };
                Settings = JsonSerializer.Deserialize<AppSettings>(json, options) ?? new AppSettings();
                DesktopLogger.Info($"Loaded appsettings.json. ApiBaseUrl={Settings.ApiBaseUrl} CallSessionId={Settings.CallSessionId} TokenPresent={!string.IsNullOrWhiteSpace(Settings.ApiBearerToken)}");
            }
        }
        catch
        {
            Settings = new AppSettings();
            DesktopLogger.Error("Failed to load appsettings.json; using defaults.");
        }

        // If launched via orioai://start?sessionId=...&apiBaseUrl=...&token=..., override settings so this session is used for logging
        ApplyProtocolLaunchArgs(e.Args);
        DesktopLogger.Info($"After protocol parse: LaunchedViaProtocol={LaunchedViaProtocol} ApiBaseUrl={Settings.ApiBaseUrl} CallSessionId={Settings.CallSessionId} TokenPresent={!string.IsNullOrWhiteSpace(Settings.ApiBearerToken)}");
    }

    private static void ApplyProtocolLaunchArgs(string[]? startupArgs)
    {
        // When launched via protocol from browser, the URL may be in StartupEventArgs.Args or in Environment.GetCommandLineArgs().
        // Args from WPF can be empty when launched via ShellExecute; the URL is often the second argument (index 1) in the full command line.
        var allArgs = Environment.GetCommandLineArgs();
        var urlArg = (startupArgs?.FirstOrDefault(a => a?.StartsWith("orioai://", StringComparison.OrdinalIgnoreCase) == true))
            ?? (allArgs.Length > 1 ? allArgs[1] : null)
            ?? allArgs.FirstOrDefault(a => a?.StartsWith("orioai://", StringComparison.OrdinalIgnoreCase) == true);

        if (string.IsNullOrEmpty(urlArg) || !urlArg.StartsWith("orioai://", StringComparison.OrdinalIgnoreCase))
            return;

        try
        {
            LaunchedViaProtocol = true;
            ProtocolUrl = urlArg;
            DesktopLogger.Info($"Protocol URL received: {urlArg}");
            if (!Uri.TryCreate(urlArg, UriKind.Absolute, out var uri) || uri == null) return;

            var query = uri.Query?.TrimStart('?');
            if (string.IsNullOrEmpty(query)) return;

            string? sessionId = null, apiBaseUrl = null, token = null, resumeId = null, language = null;
            foreach (var part in query.Split('&'))
            {
                var idx = part.IndexOf('=');
                if (idx <= 0) continue;
                var key = Uri.UnescapeDataString(part[..idx].Trim());
                var value = Uri.UnescapeDataString(part[(idx + 1)..].Trim());
                if (string.Equals(key, "sessionId", StringComparison.OrdinalIgnoreCase)) sessionId = value;
                else if (string.Equals(key, "apiBaseUrl", StringComparison.OrdinalIgnoreCase)) apiBaseUrl = value;
                else if (string.Equals(key, "token", StringComparison.OrdinalIgnoreCase)) token = value;
                else if (string.Equals(key, "resumeId", StringComparison.OrdinalIgnoreCase)) resumeId = value;
                else if (string.Equals(key, "language", StringComparison.OrdinalIgnoreCase)) language = value;
            }

            if (!string.IsNullOrWhiteSpace(sessionId))
                Settings.CallSessionId = sessionId.Trim();
            if (!string.IsNullOrWhiteSpace(apiBaseUrl))
                Settings.ApiBaseUrl = apiBaseUrl.Trim().TrimEnd('/') + "/";
            if (!string.IsNullOrWhiteSpace(resumeId))
                Settings.ResumeId = resumeId.Trim();
            if (!string.IsNullOrWhiteSpace(language))
                Settings.SessionLanguage = language.Trim();
            // Token is optional. If not present, keep the token from appsettings.json.
            if (!string.IsNullOrEmpty(token))
                Settings.ApiBearerToken = token;

            DesktopLogger.Info($"Protocol parsed. sessionId={sessionId} apiBaseUrl={apiBaseUrl} resumeId={resumeId} language={language} tokenPresent={!string.IsNullOrEmpty(token)}");
        }
        catch
        {
            // Ignore malformed protocol URL
            DesktopLogger.Warn("Protocol URL parse failed (malformed URL).");
        }
    }
}

public class AppSettings
{
    public AzureOpenAISettings AzureOpenAI { get; set; } = new();
    public AzureSpeechSettings AzureSpeech { get; set; } = new();
    public DesktopAuthSettings DesktopAuth { get; set; } = new();
    public string ApiBaseUrl { get; set; } = "http://localhost:5050/api/";
    public string ApiBearerToken { get; set; } = "";
    public string CallSessionId { get; set; } = "AB589C99-0980-4467-8AF5-ADAB340FE1A0";
    public string? ResumeId { get; set; }
    public string? SessionLanguage { get; set; }
}

public class DesktopAuthSettings
{
    public string AuthorizeUrl { get; set; } = "http://localhost:5173/login?callbackUrl=/dashboard";
    public string ExchangeUrl { get; set; } = "http://localhost:5050/api/auth/exchange";
    public string AuthRelayPath { get; set; } = "/auth/desktop";
    /// <summary>
    /// After a successful OAuth callback on the local HttpListener, the browser is redirected here (HTTP 302)
    /// so the user sees the web "Authentication successful" page instead of localhost plain text.
    /// If empty, the origin of <see cref="AuthorizeUrl"/> plus "/auth/desktop-complete" is used.
    /// </summary>
    public string SuccessRedirectUrl { get; set; } = "";
    public string ClientId { get; set; } = "desktop";
    public string RedirectHost { get; set; } = "127.0.0.1";
    public string RedirectPath { get; set; } = "/callback";
    public int PreferredCallbackPort { get; set; } = 5000;
    public int LoginTimeoutSeconds { get; set; } = 120;
}

public class AzureOpenAISettings
{
    public string Endpoint { get; set; } = string.Empty;
    public string Key { get; set; } = string.Empty;
    public string DeploymentName { get; set; } = "gpt-4o-mini";
    public string SystemPrompt { get; set; } = "You are a helpful AI assistant.";
}

public class AzureSpeechSettings
{
    public string Key { get; set; } = string.Empty;
    public string Region { get; set; } = "eastus";
    public string EndpointSilenceTimeoutMs { get; set; } = "500";
}

