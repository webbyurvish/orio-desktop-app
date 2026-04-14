using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;

namespace AiInterviewAssistant;

/// <summary>Batches product analytics to POST /api/analytics/events (same contract as the web app).</summary>
internal static class DesktopAnalytics
{
    private static HttpClient? _client;
    private static readonly object Gate = new();
    private static readonly List<AnalyticsEventPayload> Queue = new();
    private static Timer? _debounce;

    public static void Configure(HttpClient client) => _client = client;

    public static void Track(string eventType, string? metadataJson = null, Guid? callSessionId = null)
    {
        var c = _client;
        if (c?.DefaultRequestHeaders.Authorization == null)
            return;

        lock (Gate)
        {
            Queue.Add(new AnalyticsEventPayload
            {
                EventType = eventType,
                MetadataJson = metadataJson,
                CallSessionId = callSessionId,
                Source = "desktop",
            });

            _debounce ??= new Timer(
                _ => _ = FlushAsync(),
                null,
                1200,
                Timeout.Infinite);
        }
    }

    private static async Task FlushAsync()
    {
        List<AnalyticsEventPayload>? batch;
        lock (Gate)
        {
            if (_debounce != null)
            {
                _debounce.Dispose();
                _debounce = null;
            }

            if (Queue.Count == 0)
                return;

            batch = new List<AnalyticsEventPayload>(Queue);
            Queue.Clear();
        }

        var c = _client;
        if (c == null || batch.Count == 0)
            return;

        try
        {
            using var resp = await c.PostAsJsonAsync(
                    "analytics/events",
                    new AnalyticsBatchPayload { Events = batch })
                .ConfigureAwait(false);

            if (!resp.IsSuccessStatusCode)
                DesktopLogger.Warn($"Analytics flush failed: {(int)resp.StatusCode}");
        }
        catch (System.Exception ex)
        {
            DesktopLogger.Warn($"Analytics flush error: {ex.Message}");
        }
    }
}

/// <summary>Aligned with <c>PKeetDashboard.API.Analytics.AnalyticsEventTypes</c>.</summary>
internal static class DesktopAnalyticsEventTypes
{
    public const string SessionActivated = "SESSION_ACTIVATED";
    public const string SessionEnded = "SESSION_ENDED";
    public const string AiResponseGenerated = "AI_RESPONSE_GENERATED";
    public const string AnalyzeScreenRequested = "ANALYZE_SCREEN_REQUESTED";
}

internal sealed class AnalyticsBatchPayload
{
    [JsonPropertyName("events")]
    public List<AnalyticsEventPayload> Events { get; set; } = new();
}

internal sealed class AnalyticsEventPayload
{
    [JsonPropertyName("eventType")]
    public string EventType { get; set; } = "";

    [JsonPropertyName("metadataJson")]
    public string? MetadataJson { get; set; }

    [JsonPropertyName("callSessionId")]
    public Guid? CallSessionId { get; set; }

    [JsonPropertyName("source")]
    public string Source { get; set; } = "desktop";
}
