using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Json;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace AiInterviewAssistant;

public sealed class DesktopWebAuthService
{
    private readonly DesktopAuthSettings _settings;
    private readonly HttpClient _httpClient;
    private readonly SecureTokenStore _tokenStore;

    public DesktopWebAuthService(DesktopAuthSettings settings, HttpClient? httpClient = null)
    {
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        _httpClient = httpClient ?? new HttpClient();
        _tokenStore = new SecureTokenStore();
    }

    public async Task<DesktopAuthTokens> AuthenticateAsync(CancellationToken cancellationToken, string? attemptId = null)
    {
        var authTraceId = string.IsNullOrWhiteSpace(attemptId) ? Guid.NewGuid().ToString("N")[..8] : attemptId;
        var port = FindAvailablePort(_settings.PreferredCallbackPort);
        var callbackUri = BuildCallbackUri(port);
        var state = CreateCryptographicState();
        DesktopLogger.Info($"[AUTH:{authTraceId}] Listener starting at {callbackUri.GetLeftPart(UriPartial.Path)} stateLength={state.Length} preferredPort={_settings.PreferredCallbackPort} selectedPort={port}");

        using var listener = new HttpListener();
        listener.Prefixes.Add(callbackUri.GetLeftPart(UriPartial.Path).TrimEnd('/') + "/");
        listener.Start();

        var authUri = BuildAuthorizeUri(callbackUri, state);
        DesktopLogger.Info($"[AUTH:{authTraceId}] Opening browser URL: {authUri}");
        OpenBrowser(authUri);

        CallbackResult callback = await WaitForCallbackAsync(listener, state, _settings, cancellationToken, authTraceId).ConfigureAwait(false);
        DesktopLogger.Info($"[AUTH:{authTraceId}] Callback received with valid state.");
        DesktopAuthTokens tokens = await ExchangeCodeAsync(callback.Code, cancellationToken, authTraceId).ConfigureAwait(false);
        await _tokenStore.StoreAsync(tokens, cancellationToken).ConfigureAwait(false);
        DesktopLogger.Info($"[AUTH:{authTraceId}] Tokens exchanged and securely stored.");
        return tokens;
    }

    private Uri BuildCallbackUri(int port)
    {
        var redirectHost = string.IsNullOrWhiteSpace(_settings.RedirectHost) ? "127.0.0.1" : _settings.RedirectHost.Trim();
        var redirectPath = string.IsNullOrWhiteSpace(_settings.RedirectPath) ? "/callback" : _settings.RedirectPath.Trim();
        if (!redirectPath.StartsWith("/", StringComparison.Ordinal))
            redirectPath = "/" + redirectPath;

        return new Uri($"http://{redirectHost}:{port}{redirectPath}");
    }

    private Uri BuildAuthorizeUri(Uri callbackUri, string state)
    {
        var baseUri = _settings.AuthorizeUrl?.Trim();
        if (string.IsNullOrWhiteSpace(baseUri))
            throw new InvalidOperationException("Desktop auth authorize URL is not configured.");

        var uriBuilder = new UriBuilder(baseUri);
        var query = ParseQuery(uriBuilder.Query);
        var clientId = _settings.ClientId?.Trim() ?? "desktop";
        var redirectUri = callbackUri.ToString();

        // If URL is a login page with callbackUrl (e.g. /login?callbackUrl=/dashboard),
        // route post-login back through authRelayPath so desktop params survive login.
        if (query.ContainsKey("callbackUrl"))
        {
            var relayPath = string.IsNullOrWhiteSpace(_settings.AuthRelayPath) ? "/auth" : _settings.AuthRelayPath.Trim();
            if (!relayPath.StartsWith("/", StringComparison.Ordinal))
                relayPath = "/" + relayPath;
            var relayUri = $"{relayPath}?client={Uri.EscapeDataString(clientId)}&redirect_uri={Uri.EscapeDataString(redirectUri)}&state={Uri.EscapeDataString(state)}";
            query["callbackUrl"] = relayUri;
        }
        else
        {
            query["client"] = clientId;
            query["redirect_uri"] = redirectUri;
            query["state"] = state;
        }

        uriBuilder.Query = BuildQuery(query);
        return uriBuilder.Uri;
    }

    private static Dictionary<string, string> ParseQuery(string? rawQuery)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(rawQuery))
            return result;

        var query = rawQuery.TrimStart('?');
        if (string.IsNullOrWhiteSpace(query))
            return result;

        foreach (var segment in query.Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            var index = segment.IndexOf('=');
            if (index < 0)
            {
                result[Uri.UnescapeDataString(segment)] = string.Empty;
                continue;
            }

            var key = Uri.UnescapeDataString(segment[..index]);
            var value = Uri.UnescapeDataString(segment[(index + 1)..]);
            result[key] = value;
        }

        return result;
    }

    private static string BuildQuery(Dictionary<string, string> values)
    {
        var builder = new StringBuilder();
        foreach (var pair in values)
        {
            if (builder.Length > 0)
                builder.Append('&');
            builder.Append(Uri.EscapeDataString(pair.Key));
            builder.Append('=');
            builder.Append(Uri.EscapeDataString(pair.Value ?? string.Empty));
        }

        return builder.ToString();
    }

    private static void OpenBrowser(Uri authUri)
    {
        var psi = new ProcessStartInfo
        {
            FileName = authUri.ToString(),
            UseShellExecute = true
        };
        Process.Start(psi);
    }

    private static async Task<CallbackResult> WaitForCallbackAsync(
        HttpListener listener,
        string expectedState,
        DesktopAuthSettings settings,
        CancellationToken cancellationToken,
        string authTraceId)
    {
        while (true)
        {
            Task<HttpListenerContext> contextTask = listener.GetContextAsync();
            Task completed = await Task.WhenAny(contextTask, Task.Delay(Timeout.Infinite, cancellationToken)).ConfigureAwait(false);
            if (completed != contextTask)
                throw new OperationCanceledException("Login was cancelled or timed out.", cancellationToken);

            HttpListenerContext context = await contextTask.ConfigureAwait(false);
            DesktopLogger.Info($"[AUTH:{authTraceId}] Incoming callback request method={context.Request.HttpMethod} rawUrl={context.Request.RawUrl}");
            if (!string.Equals(context.Request.HttpMethod, "GET", StringComparison.OrdinalIgnoreCase))
            {
                await RespondAsync(context.Response, HttpStatusCode.MethodNotAllowed, "Method not allowed.").ConfigureAwait(false);
                continue;
            }

            var query = context.Request.QueryString;
            string? error = query["error"];
            string? code = query["code"];
            string? state = query["state"];
            var codeLength = string.IsNullOrWhiteSpace(code) ? 0 : code.Length;
            DesktopLogger.Info($"[AUTH:{authTraceId}] Callback query parsed codeLength={codeLength} hasState={!string.IsNullOrWhiteSpace(state)} hasError={!string.IsNullOrWhiteSpace(error)}");

            if (!string.IsNullOrWhiteSpace(error))
            {
                await RespondAsync(context.Response, HttpStatusCode.BadRequest, "Login failed. You can close this tab and try again.").ConfigureAwait(false);
                throw new InvalidOperationException($"Web login failed: {error}");
            }

            if (string.IsNullOrWhiteSpace(code) || string.IsNullOrWhiteSpace(state))
            {
                await RespondAsync(context.Response, HttpStatusCode.BadRequest, "Invalid callback payload. You can close this tab and try again.").ConfigureAwait(false);
                throw new InvalidOperationException("Missing auth code/state in callback.");
            }

            if (!FixedTimeEquals(state, expectedState))
            {
                await RespondAsync(context.Response, HttpStatusCode.BadRequest, "Invalid login state. You can close this tab and try again.").ConfigureAwait(false);
                throw new InvalidOperationException("State validation failed.");
            }

            var successPage = GetSuccessRedirectUrl(settings);
            if (!string.IsNullOrWhiteSpace(successPage))
            {
                await RespondRedirectAsync(context.Response, successPage.Trim()).ConfigureAwait(false);
            }
            else
            {
                await RespondAsync(context.Response, HttpStatusCode.OK, "Login completed. You can close this tab and return to the desktop app.").ConfigureAwait(false);
            }

            return new CallbackResult(code);
        }
    }

    private static string? GetSuccessRedirectUrl(DesktopAuthSettings settings)
    {
        var configured = settings.SuccessRedirectUrl?.Trim();
        if (!string.IsNullOrWhiteSpace(configured))
            return configured;

        var authorize = settings.AuthorizeUrl?.Trim();
        if (string.IsNullOrWhiteSpace(authorize) || !Uri.TryCreate(authorize, UriKind.Absolute, out var authUri))
            return null;

        return new Uri(authUri, "/auth/desktop-complete").ToString();
    }

    private async Task<DesktopAuthTokens> ExchangeCodeAsync(string code, CancellationToken cancellationToken, string authTraceId)
    {
        var exchangeUrl = _settings.ExchangeUrl?.Trim();
        if (string.IsNullOrWhiteSpace(exchangeUrl))
            throw new InvalidOperationException("Desktop auth exchange URL is not configured.");
        DesktopLogger.Info($"[AUTH:{authTraceId}] Exchanging code at {exchangeUrl} codeLength={code.Length}");

        using var request = new HttpRequestMessage(HttpMethod.Post, exchangeUrl)
        {
            Content = JsonContent.Create(new { code })
        };

        using HttpResponseMessage response = await _httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
        string body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
            throw new InvalidOperationException($"Code exchange failed ({(int)response.StatusCode}): {body}");
        DesktopLogger.Info($"[AUTH:{authTraceId}] Exchange response status={(int)response.StatusCode} bodyLength={body.Length}");

        var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        var payload = JsonSerializer.Deserialize<DesktopAuthTokens>(body, options);
        if (payload == null || string.IsNullOrWhiteSpace(payload.AccessToken) || string.IsNullOrWhiteSpace(payload.RefreshToken))
            throw new InvalidOperationException("Code exchange response is missing tokens.");

        return payload with { IssuedUtc = DateTimeOffset.UtcNow };
    }

    private static async Task RespondAsync(HttpListenerResponse response, HttpStatusCode statusCode, string message)
    {
        response.StatusCode = (int)statusCode;
        response.ContentType = "text/html; charset=utf-8";
        byte[] bytes = Encoding.UTF8.GetBytes($"<html><body><h3>{WebUtility.HtmlEncode(message)}</h3></body></html>");
        response.ContentLength64 = bytes.Length;
        await response.OutputStream.WriteAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
        response.OutputStream.Close();
    }

    private static async Task RespondRedirectAsync(HttpListenerResponse response, string absoluteUrl)
    {
        response.StatusCode = (int)HttpStatusCode.Redirect;
        response.AddHeader("Location", absoluteUrl);
        response.ContentType = "text/html; charset=utf-8";
        const string body = "<html><body>Redirecting…</body></html>";
        byte[] bytes = Encoding.UTF8.GetBytes(body);
        response.ContentLength64 = bytes.Length;
        await response.OutputStream.WriteAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
        response.OutputStream.Close();
    }

    private static string CreateCryptographicState()
    {
        Span<byte> random = stackalloc byte[32];
        RandomNumberGenerator.Fill(random);
        return Convert.ToBase64String(random).TrimEnd('=').Replace('+', '-').Replace('/', '_');
    }

    private static int FindAvailablePort(int preferredPort)
    {
        if (preferredPort > 0 && IsPortAvailable(preferredPort))
            return preferredPort;

        using var probe = new TcpListener(IPAddress.Loopback, 0);
        probe.Start();
        return ((IPEndPoint)probe.LocalEndpoint).Port;
    }

    private static bool IsPortAvailable(int port)
    {
        try
        {
            using var probe = new TcpListener(IPAddress.Loopback, port);
            probe.Start();
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool FixedTimeEquals(string left, string right)
    {
        byte[] leftBytes = Encoding.UTF8.GetBytes(left);
        byte[] rightBytes = Encoding.UTF8.GetBytes(right);
        return CryptographicOperations.FixedTimeEquals(leftBytes, rightBytes);
    }

    private readonly record struct CallbackResult(string Code);
}

public sealed record DesktopAuthTokens
{
    public string AccessToken { get; init; } = string.Empty;
    public string RefreshToken { get; init; } = string.Empty;
    public DateTimeOffset IssuedUtc { get; init; } = DateTimeOffset.UtcNow;
}

public sealed class SecureTokenStore
{
    private static readonly string TokenPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SmeedAI",
        "AiInterviewAssistant",
        "tokens.dat");

    public async Task StoreAsync(DesktopAuthTokens tokens, CancellationToken cancellationToken)
    {
        string json = JsonSerializer.Serialize(tokens);
        byte[] plain = Encoding.UTF8.GetBytes(json);
        byte[] encrypted = ProtectedData.Protect(plain, null, DataProtectionScope.CurrentUser);
        var directory = Path.GetDirectoryName(TokenPath);
        if (!string.IsNullOrWhiteSpace(directory))
            Directory.CreateDirectory(directory);
        await File.WriteAllBytesAsync(TokenPath, encrypted, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Removes persisted OAuth tokens so the next launch requires sign-in again.</summary>
    public static void ClearPersistedTokens()
    {
        try
        {
            if (File.Exists(TokenPath))
                File.Delete(TokenPath);
        }
        catch
        {
            // Ignore IO errors; logout still clears in-memory auth.
        }
    }
}
