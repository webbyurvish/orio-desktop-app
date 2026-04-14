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
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
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
        // SQL
        "select", "from", "where", "join", "left", "right", "inner", "outer", "cross", "full", "on", "as", "and", "or", "not",
        "group", "by", "order", "limit", "offset", "having", "union", "all", "distinct", "insert", "into", "values", "update",
        "set", "delete", "create", "table", "alter", "drop", "truncate", "index", "view", "primary", "key", "foreign",
        "references", "constraint", "default", "check", "unique", "cascade", "between", "in", "like", "ilike", "exists",
        "case", "when", "then", "else", "end", "with", "recursive", "over", "partition", "window", "asc", "desc", "nulls",
        "first", "last", "count", "sum", "avg", "min", "max", "cast", "coalesce", "nullif", "is", "null", "true", "false",
        // C# / general
        "if", "else", "for", "foreach", "while", "return", "class", "public", "private", "protected", "internal", "static",
        "void", "string", "int", "bool", "var", "new", "using", "async", "await", "try", "catch", "finally", "switch",
        "case", "break", "continue", "namespace", "interface", "enum", "struct", "readonly", "const", "delegate", "event",
        "typeof", "nameof", "lock", "yield", "throw", "this", "base", "sizeof", "stackalloc", "record", "init",
        "get", "set", "add", "remove", "operator", "implicit", "explicit", "params", "ref", "out", "in", "let", "into",
        "orderby", "ascending", "descending", "equals"
    };

    private static SolidColorBrush FreezeBrush(byte r, byte g, byte b)
    {
        var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
        if (brush.CanFreeze)
            brush.Freeze();
        return brush;
    }

    private static SolidColorBrush FreezeBrushAlpha(byte a, byte r, byte g, byte b)
    {
        var brush = new SolidColorBrush(Color.FromArgb(a, r, g, b));
        if (brush.CanFreeze)
            brush.Freeze();
        return brush;
    }

    /// <summary>Strips default WPF button chrome so only <see cref="Button.Content"/> is visible (code-block copy control).</summary>
    private static ControlTemplate CreateChromeOnlyButtonTemplate()
    {
        var presenter = new FrameworkElementFactory(typeof(ContentPresenter));
        presenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Stretch);
        presenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Stretch);
        var template = new ControlTemplate(typeof(Button));
        template.VisualTree = presenter;
        return template;
    }

    private static readonly SolidColorBrush CodeHostBackgroundBrush = FreezeBrush(0x26, 0x28, 0x34);
    private static readonly SolidColorBrush CodeHostBorderBrush = FreezeBrush(0x58, 0x5C, 0x70);
    private static readonly SolidColorBrush CodeFgDefaultBrush = FreezeBrush(0xEE, 0xF0, 0xF4);
    private static readonly SolidColorBrush CodeFgKeywordBrush = FreezeBrush(0x93, 0xC5, 0xFD);
    private static readonly SolidColorBrush CodeFgStringBrush = FreezeBrush(0xFC, 0xD3, 0x4D);
    private static readonly SolidColorBrush CodeFgNumberBrush = FreezeBrush(0xF0, 0xD9, 0x90);
    private static readonly SolidColorBrush CodeFgCommentBrush = FreezeBrush(0xA3, 0xD9, 0xA5);

    // Code block “card” styling (matches desktop reference screenshot).
    private static readonly SolidColorBrush CodeCardOuterBgBrush = FreezeBrush(0x0B, 0x12, 0x20);
    private static readonly SolidColorBrush CodeCardOuterBorderBrush = FreezeBrush(0x1E, 0x29, 0x3B);
    private static readonly SolidColorBrush CodeCardHeaderBgBrush = FreezeBrush(0x0F, 0x17, 0x2A);
    private static readonly SolidColorBrush CodeCardHeaderBorderBrush = FreezeBrush(0x1E, 0x29, 0x3B);
    private static readonly SolidColorBrush CodeCardLangBrush = FreezeBrush(0xC0, 0x84, 0xFC);
    private static readonly SolidColorBrush CodeCardCopyFgBrush = FreezeBrush(0xC7, 0xD2, 0xFE);
    private static readonly SolidColorBrush CodeCardCopyHoverBgBrush = FreezeBrushAlpha(0x2A, 0xFF, 0xFF, 0xFF);
    private static readonly SolidColorBrush CodeCardCopyHoverBorderBrush = FreezeBrushAlpha(0x30, 0xFF, 0xFF, 0xFF);
    private static readonly SolidColorBrush CodeCardCopyPressedBgBrush = FreezeBrushAlpha(0x1F, 0xFF, 0xFF, 0xFF);
    private static readonly SolidColorBrush CodeCardCopyBorderBrush = FreezeBrushAlpha(0x22, 0xFF, 0xFF, 0xFF);

    /// <summary>Muted coral / amber — question title (matches reference UI).</summary>
    private static readonly SolidColorBrush AnswerQuestionHeadingBrush = FreezeBrush(0xfb, 0x92, 0x3c);
    private static readonly SolidColorBrush AnswerKeywordHighlightBrush = FreezeBrush(0x7d, 0xd3, 0xfc);
    private static readonly SolidColorBrush AnswerBodyBrush = FreezeBrush(0xe2, 0xe8, 0xf0);
    /// <summary>Rotating styles for each <c>**bold**</c> span in a line (pink, cyan, violet, mint, coral, fuchsia).</summary>
    private static readonly (SolidColorBrush Fg, SolidColorBrush Bg)[] AnswerBoldEmphasisStyles =
    {
        (FreezeBrush(0xf4, 0x72, 0xb6), FreezeBrushAlpha(0x55, 0x5c, 0x1a, 0x3d)),
        (FreezeBrush(0x67, 0xe8, 0xf9), FreezeBrushAlpha(0x58, 0x14, 0x5c, 0x72)),
        (FreezeBrush(0xc4, 0xb5, 0xfd), FreezeBrushAlpha(0x52, 0x3a, 0x2e, 0x68)),
        (FreezeBrush(0x86, 0xef, 0xac), FreezeBrushAlpha(0x50, 0x16, 0x54, 0x3a)),
        (FreezeBrush(0xfc, 0xa5, 0xa5), FreezeBrushAlpha(0x55, 0x5c, 0x22, 0x22)),
        (FreezeBrush(0xf0, 0xab, 0xfc), FreezeBrushAlpha(0x52, 0x52, 0x1a, 0x5c)),
    };
    /// <summary>==mark== — amber text on warm panel (distinct palette from **bold**).</summary>
    private static readonly SolidColorBrush AnswerMarkFgBrush = FreezeBrush(0xfb, 0xbf, 0x24);
    private static readonly SolidColorBrush AnswerMarkBgBrush = FreezeBrushAlpha(0x55, 0x5a, 0x3e, 0x0f);
    private static readonly SolidColorBrush KeywordPillBgBrush = FreezeBrush(0x1e, 0x3a, 0x52);
    private static readonly SolidColorBrush KeywordPillBorderBrush = FreezeBrush(0x3d, 0x5a, 0x80);
    /// <summary>Glossary auto-highlight in answers — brighter sky on indigo tint.</summary>
    private static readonly SolidColorBrush KeywordPillFgBrush = FreezeBrush(0xc4, 0xe5, 0xff);

    private static readonly string[] InterviewGlossaryTerms =
    {
        "ROW_NUMBER()", "DENSE_RANK()", "RANK()", "NTILE()", "PARTITION BY",
        "Transaction Isolation", "Stored Procedures", "Stored Procedure", "Entity Framework",
        "Strong Consistency", "Eventual Consistency", "Consistent Hashing", "Message Queue",
        "T-SQL", "Microsoft SQL Server", "SQL Server", "CAP Theorem", "Kubernetes", "PostgreSQL", "MongoDB", "Microservices",
        "JavaScript", "TypeScript", "Async/Await", "OAuth 2.0", "WebSocket", "Load Balancer",
        "Rate Limiting", "Circuit Breaker", "Idempotency", "Normalization", "Denormalization",
        "Deadlock", "Garbage Collection", "Memory Leak", "Thread Safety", "Race Condition",
        "Binary Search", "Dynamic Programming", "Hash Table", "Linked List", "Binary Tree",
        "REST API", "GraphQL", "gRPC", "HTTPS", "TLS", "CI/CD",
        ".NET", "C#", "LINQ", "ACID", "Redis", "Kafka", "Docker", "Azure", "AWS", "GCP",
        "Mutex", "Semaphore", "Big O", "JWT", "OAuth", "ORM", "TCP/IP", "JSON", "XML", "HTML", "CSS",
        "Scalability", "Throughput", "Latency", "Concurrency", "Sharding", "Replication", "Indexing",
    };

    private static readonly Lazy<Regex> InterviewGlossaryRegex = new(BuildInterviewGlossaryRegex);

    /// <summary>**primary** and ==secondary== emphasis in one pass (non-overlapping).</summary>
    private static readonly Regex MarkdownEmphasisRegex = new(
        @"\*\*(?<bold>.+?)\*\*|==(?<mark>.+?)==",
        RegexOptions.Singleline | RegexOptions.Compiled);

    private static readonly Regex AltQuestionsBlockRegex = new(
        @"<<<ALT_Q>>>\s*(?<body>[\s\S]*?)\s*<<<END_ALT_Q>>>",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly SolidColorBrush GlossaryHighlightBgBrush = FreezeBrushAlpha(0x46, 0x2e, 0x32, 0x72);

    /// <summary>Segoe UI Emoji enables colored Unicode emoji in the answer RichTextBox (Windows).</summary>
    private static readonly FontFamily AnswerEmojiFontFamily = new("Segoe UI, Segoe UI Emoji");

    /// <summary>Desktop WPF does not paint COLR emoji; Emoji.Wpf replaces sequences with colored vector glyphs.</summary>
    private static void ApplyAnswerEmojiGlyphs(FlowDocument doc)
    {
        if (doc == null)
            return;
        Emoji.Wpf.FlowDocumentExtensions.SubstituteGlyphs(doc);
    }

    /// <summary>Phrase tokens (spaces/parens) use leading/trailing non-identifier guards; single tokens cannot match inside words (e.g. ORM in "platform").</summary>
    private static string GlossaryAlternationPatternForTerm(string term)
    {
        var e = Regex.Escape(term);
        if (term.Any(c => c is ' ' or '\t' or '(' or ')'))
            return $"(?<![\\p{{L}}\\p{{N}}_]){e}(?![\\p{{L}}\\p{{N}}_])";
        if (term.StartsWith(".", StringComparison.Ordinal))
            return $"(?<![\\p{{L}}\\p{{N}}]){e}(?![\\p{{L}}\\p{{N}}_#])";
        return $"(?<![\\p{{L}}\\p{{N}}_]){e}(?![\\p{{L}}\\p{{N}}_#])";
    }

    private static Regex BuildInterviewGlossaryRegex()
    {
        var parts = InterviewGlossaryTerms
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(t => t.Length)
            .Select(GlossaryAlternationPatternForTerm);
        return new Regex(string.Join("|", parts), RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
    }

    private static string StripAltQuestionsForDisplay(string raw)
    {
        if (string.IsNullOrEmpty(raw)) return raw;
        var idx = raw.IndexOf("<<<ALT_Q>>>", StringComparison.OrdinalIgnoreCase);
        if (idx < 0) return raw;
        return raw[..idx].TrimEnd();
    }

    private static List<string> ExtractAltQuestions(string raw)
    {
        var list = new List<string>();
        if (string.IsNullOrWhiteSpace(raw)) return list;
        var m = AltQuestionsBlockRegex.Match(raw);
        if (!m.Success) return list;
        var body = m.Groups["body"].Value;
        foreach (var rawLine in body.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            var s = Regex.Replace(rawLine.Trim(), @"^\d+\.\s*", "").Trim();
            if (s.Length >= 8)
                list.Add(s);
        }

        return list.Count > 3 ? list.Take(3).ToList() : list;
    }

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
        "Answer the question fully, using resume details whenever relevant. " +
        "For questions like 'what is your name' or 'introduce yourself', answer using the candidate's real name and background from the resume. " +
        "Write so each line or bullet is easy to read quickly and sounds natural when spoken aloud.";

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
        public bool NaturalSpeakingMode { get; set; }
        public string? ExtraContext { get; set; }
        public string? AiModel { get; set; }
        public bool SaveTranscript { get; set; }
        public bool IsFreeSession { get; set; }
    }

    private sealed class CallSessionDto
    {
        public Guid Id { get; set; }
        public Guid? ResumeId { get; set; }
        public string? ExtraContext { get; set; }
        public bool NaturalSpeakingMode { get; set; }
    }

    private sealed class CurrentUserDto
    {
        public decimal CallCredits { get; set; }
        public string? Email { get; set; }
    }

    private const int WDA_NONE = 0;
    private const int WDA_EXCLUDEFROMCAPTURE = 0x11;

    /// <summary>
    /// TEMPORARY (screenshots / demos): <c>false</c> = window is visible to screen capture and sharing.
    /// Restore to <c>true</c> before shipping so the interview overlay stays hidden from captures again.
    /// </summary>
    private const bool ExcludeMainWindowFromScreenCapture = false;

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
    private readonly StringBuilder _finalTranscriptDetail = new();
    private string _partialMic = string.Empty;
    private string _partialSystem = string.Empty;
    // Transcript display is full + continuous, but AI input must be ONLY the latest question.
    private string _latestQuestionText = string.Empty;
    private DateTimeOffset _latestQuestionUtc = DateTimeOffset.MinValue;
    private string _previousQuestionText = string.Empty;
    private string _previousAnswerShortContext = string.Empty;
    private DateTimeOffset _lastTranscriptAppendUtc = DateTimeOffset.MinValue;
    private bool _transcriptDetailAutoScroll = true;
    private bool _transcriptDetailShown;
    private double _baseWidthBeforeTranscript;
    private double _baseLeftBeforeTranscript;
    private double _baseMainCardsColumnWidth;
    private bool _transcriptDockedRight;
    private bool _mainUiPositionInitialized;
    private bool _isListening;
    private bool _micOn = true;
    private bool _speakerOn = true;

    private readonly string _deploymentName;
    private readonly string _systemPrompt;

    private static bool IsFollowUpQuestion(string q)
    {
        var t = (q ?? string.Empty).Trim();
        if (t.Length == 0) return false;

        var lower = t.ToLowerInvariant();
        // Short prompts that start with a conjunction/pronoun usually need prior context.
        if (lower.StartsWith("and ", StringComparison.Ordinal)
            || lower.StartsWith("also ", StringComparison.Ordinal)
            || lower.StartsWith("then ", StringComparison.Ordinal)
            || lower.StartsWith("so ", StringComparison.Ordinal)
            || lower.StartsWith("what about", StringComparison.Ordinal)
            || lower.StartsWith("how about", StringComparison.Ordinal)
            || lower.StartsWith("in that case", StringComparison.Ordinal)
            || lower.StartsWith("in this case", StringComparison.Ordinal)
            || lower.StartsWith("for that", StringComparison.Ordinal)
            || lower.StartsWith("for this", StringComparison.Ordinal)
            || lower.StartsWith("can you also", StringComparison.Ordinal)
            || lower.StartsWith("could you also", StringComparison.Ordinal))
            return true;

        var words = t.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
        if (words <= 8)
        {
            if (lower.Contains("it ", StringComparison.Ordinal) || lower.Contains("that ", StringComparison.Ordinal) || lower.Contains("this ", StringComparison.Ordinal)
                || lower.Contains("they ", StringComparison.Ordinal) || lower.Contains("those ", StringComparison.Ordinal))
                return true;
        }

        return false;
    }

    private static string BuildLatestQuestionOnlyUserContent(
        string currentQuestion,
        string? typedHint,
        string? previousQuestion,
        string? previousAnswerSummary)
    {
        var q = (currentQuestion ?? string.Empty).Trim();
        var hint = (typedHint ?? string.Empty).Trim();
        var prevQ = (previousQuestion ?? string.Empty).Trim();
        var prevA = (previousAnswerSummary ?? string.Empty).Trim();

        // Hard separation: the model must treat CURRENT QUESTION as the only question to answer.
        var sb = new StringBuilder();
        sb.AppendLine("You are helping in a live interview. Answer ONLY the CURRENT QUESTION.");
        sb.AppendLine("Do NOT answer earlier questions, and do NOT merge multiple questions together.");
        sb.AppendLine();
        sb.AppendLine("CURRENT QUESTION:");
        sb.AppendLine(q);

        if (!string.IsNullOrWhiteSpace(hint))
        {
            sb.AppendLine();
            sb.AppendLine("TYPED HINT (optional; may contain extra words):");
            sb.AppendLine(hint);
        }

        // Minimal follow-up context: include only when it looks like a follow-up AND we have prior state.
        if (IsFollowUpQuestion(q) && !string.IsNullOrWhiteSpace(prevQ))
        {
            sb.AppendLine();
            sb.AppendLine("MINIMAL CONTEXT (only use if needed for the follow-up):");
            sb.AppendLine($"Previous question: {prevQ}");
            if (!string.IsNullOrWhiteSpace(prevA))
                sb.AppendLine($"Previous answer summary: {prevA}");
        }

        return sb.ToString().TrimEnd();
    }
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

    private sealed class DesktopClarifyTranscriptQuestionRequest
    {
        public string? Transcript { get; set; }
        public string? UserContext { get; set; }
    }

    private sealed class DesktopClarifyTranscriptQuestionResponse
    {
        public string? Heading { get; set; }
        public string? Body { get; set; }
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

    /// <summary>Paragraph that hosts <see cref="_streamingAnswerRun"/> during streaming; removed before final render when appending.</summary>
    private Paragraph? _streamingAnswerParagraph;

    /// <summary>Coalesces token deltas onto the UI thread (~60fps) to avoid dispatcher backlog + Emoji.Wpf work per network line.</summary>
    private readonly object _answerStreamLiveTextLock = new();
    // Pending tokens to append to the UI Run on the next flush tick (keeps UI updates O(delta), not O(totalText)).
    private readonly Queue<string> _answerStreamLivePendingTokens = new();
    private volatile bool _answerStreamLiveRendering;
    private int _answerStreamRenderEpochSnap;
    private Guid _answerStreamRenderCallSessionId;
    private DispatcherTimer? _answerStreamFlushTimer;
    // When the network stream ends, keep "typing" until backlog drains, then do final rich render.
    private bool _answerStreamCompletionPending;
    private string _answerStreamCompletionContent = string.Empty;
    private string? _answerStreamCompletionHeading;
    private bool _answerStreamCompletionAppendContinuation;
    private long _answerStreamLastDebugUiLogTicks;
    private long _answerStreamDebugUiFlushCount;
    // UI pump pacing:
    // - Keep a "typed" feel by default (small deltas per tick)
    // - Automatically speed up if we start falling behind (prevents end-of-stream jump)
    private const int AnswerStreamUiTickMs = 16;
    private const int AnswerStreamUiSoftBacklogTokens = 120;
    private const int AnswerStreamUiHardBacklogTokens = 260;

    private const int AnswerStreamUiTokensPerTick_Steady = 10;
    private const int AnswerStreamUiCharsPerTick_Steady = 46;

    private const int AnswerStreamUiTokensPerTick_CatchUp = 28;
    private const int AnswerStreamUiCharsPerTick_CatchUp = 110;

    private const int AnswerStreamUiTokensPerTick_HardCatchUp = 72;
    private const int AnswerStreamUiCharsPerTick_HardCatchUp = 230;

    /// <summary>Set when sending from a follow-up chip so the next stream appends in the same answer panel.</summary>
    private bool _continueAnswerFromFollowUpChip;

    /// <summary>Snapshot for the in-flight stream: whether this response appends after the previous answer.</summary>
    private bool _activeStreamAppendsInPlace;

    private DispatcherTimer? _answerLayoutCoalesceTimer;

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

    /// <summary>Session-scoped extra instructions (from create/edit on server); applied to every AI answer and clarify call.</summary>
    private string? _sessionExtraContext;

    /// <summary>From call session: when true, steer answers toward conversational spoken-human tone.</summary>
    private bool _naturalSpeakingMode;

    /// <summary>Bumped when a session ends or a new interview starts so in-flight AI answer streams cannot repaint the UI for the wrong session.</summary>
    private int _answerUiEpoch;

    /// <summary>Matches <see cref="_answerUiEpoch"/> snapshot for the answer stream that disabled the AI buttons; cleared on bump or release.</summary>
    private int _answerStreamLeaseEpoch;

    private readonly SemaphoreSlim _creditsRefreshLock = new(1, 1);
    private DateTimeOffset _lastCreditsRefreshUtc = DateTimeOffset.MinValue;
    private bool _pendingCreditsRefreshOnActivate;

    private readonly List<AnswerHistoryItem> _sessionAnswerHistory = new();
    private int _answerHistoryViewIndex = -1;
    private int _lastAppendedAnswerHistoryIndex = -1;
    private bool _answerGenerationInFlight;
    private bool _showAskAiAnythingRow;
    private bool _speechInitInFlight;
    private bool _answerAutoScrollEnabled = true;

    private DispatcherTimer? _callAutoAnswerDebounceTimer;
    private DispatcherTimer? _micAutoAnswerDebounceTimer;
    private string _callAutoAnswerBuffer = string.Empty;
    private string _micAutoAnswerBuffer = string.Empty;
    private string _lastAutoAnswerNormKey = string.Empty;
    private DateTimeOffset _lastAutoAnswerUtc = DateTimeOffset.MinValue;

    /// <summary>Former status strip (logging / live state) was removed from the toolbar; keep messages in the log file only.</summary>
    private static void LogInterviewUiStatus(string message)
    {
        if (string.IsNullOrWhiteSpace(message)) return;
        DesktopLogger.Info($"[Interview] {message}");
    }

    public MainWindow()
    {
        InitializeComponent();
        SourceInitialized += (_, _) => WindowPrivacy.Apply(this);

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
        SessionSetupView.BuyCreditsRequested += (_, _) => OpenWebAppPricing();
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
            if (!isFree && !CreditsState.Current.HasSufficientCreditsForPaidActivation())
            {
                MessageBox.Show(
                    "You need at least 0.5 interview credit to activate a full session. Buy credits on the website.",
                    "Activate Session",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                OpenWebAppPricing();
                return;
            }

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
        ActivateSessionView.BuyCreditsRequested += (_, _) => OpenWebAppPricing();
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

        if (InterviewMiddleScrollViewer != null)
            InterviewMiddleScrollViewer.ScrollChanged += InterviewMiddleScrollViewer_ScrollChanged;

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

        _answerLayoutCoalesceTimer = new DispatcherTimer(DispatcherPriority.Background)
        {
            Interval = TimeSpan.FromMilliseconds(120)
        };
        _answerLayoutCoalesceTimer.Tick += (_, _) =>
        {
            _answerLayoutCoalesceTimer?.Stop();
            CoerceMainWindowHeightToContent();
            ScrollInterviewMiddleToEnd();
        };

        _answerStreamFlushTimer = new DispatcherTimer(DispatcherPriority.Normal)
        {
            Interval = TimeSpan.FromMilliseconds(AnswerStreamUiTickMs)
        };
        _answerStreamFlushTimer.Tick += OnAnswerStreamFlushTimerTick;

        Activated += (_, _) =>
        {
            if (!_pendingCreditsRefreshOnActivate) return;
            _pendingCreditsRefreshOnActivate = false;
            _ = RefreshCreditsFromServerAsync(force: true, showLoading: false);
        };
    }

    /// <summary>
    /// Caps the answer host height to the visible desktop below it so the main RichTextBox gets a finite height
    /// and shows its vertical scrollbar instead of clipping content.
    /// </summary>
    private void UpdateAiAnswerBodyMaxHeight()
    {
        // Middle region scrolls in InterviewMiddleScrollViewer; do not cap RichTextBox height (avoids double scrollbars).
        AiAnswerBodyBorder?.ClearValue(FrameworkElement.MaxHeightProperty);
        if (_mainUiStarted && AnswerSectionPanel.Visibility == Visibility.Visible)
        {
            InterviewMiddleScrollViewer?.InvalidateMeasure();
            ScheduleAnswerLayoutRefresh();
        }
    }

    private void ScrollInterviewMiddleToEnd()
    {
        try
        {
            if (!_answerAutoScrollEnabled)
                return;
            if (InterviewMiddleScrollViewer?.Visibility == Visibility.Visible)
                InterviewMiddleScrollViewer.ScrollToEnd();
        }
        catch
        {
            /* ignore */
        }
    }

    private void ScrollInterviewMiddleToHome()
    {
        try
        {
            if (InterviewMiddleScrollViewer?.Visibility == Visibility.Visible)
                InterviewMiddleScrollViewer.ScrollToHome();
        }
        catch
        {
            /* ignore */
        }
    }

    private void InterviewMiddleScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
    {
        try
        {
            if (sender is not ScrollViewer sv)
                return;

            // Consider “at bottom” when within a small threshold so minor layout changes don't disable auto-scroll.
            var remaining = sv.ExtentHeight - sv.VerticalOffset - sv.ViewportHeight;
            _answerAutoScrollEnabled = remaining < 24;
        }
        catch
        {
            // best-effort only
        }
    }

    /// <summary>Re-run height measurement so <see cref="Window.SizeToContent"/> grows with the answer.</summary>
    private void CoerceMainWindowHeightToContent()
    {
        if (!_mainUiStarted)
            return;

        try
        {
            InvalidateMeasure();
            MainContentGrid?.InvalidateMeasure();
            UpdateLayout();
            var mode = SizeToContent;
            if (mode is SizeToContent.Height or SizeToContent.WidthAndHeight)
            {
                SizeToContent = SizeToContent.Manual;
                Height = double.NaN;
                SizeToContent = mode;
            }
        }
        catch
        {
            /* ignore */
        }
    }

    private void ScheduleAnswerLayoutRefresh()
    {
        if (_answerLayoutCoalesceTimer == null)
            return;
        _answerLayoutCoalesceTimer.Stop();
        _answerLayoutCoalesceTimer.Start();
    }

    private bool AnswerStreamLiveRenderingStillValid()
    {
        if (!_answerStreamLiveRendering)
            return false;
        return Volatile.Read(ref _answerUiEpoch) == _answerStreamRenderEpochSnap
               && _callSessionId == _answerStreamRenderCallSessionId
               && _sessionActive;
    }

    private void OnAnswerStreamFlushTimerTick(object? sender, EventArgs e)
    {
        if (!AnswerStreamLiveRenderingStillValid())
        {
            _answerStreamFlushTimer?.Stop();
            return;
        }

        if (_streamingAnswerRun == null)
            return;

        string append = string.Empty;
        var shouldFinalize = false;
        lock (_answerStreamLiveTextLock)
        {
            if (_answerStreamLivePendingTokens.Count == 0)
            {
                if (_answerStreamCompletionPending)
                {
                    // Network finished and backlog drained; finalize with rich render.
                    shouldFinalize = true;
                    _answerStreamCompletionPending = false;
                }
                else
                {
                    _answerStreamFlushTimer?.Stop();
                    return;
                }
            }

            if (!shouldFinalize)
            {
                var pendingBefore = _answerStreamLivePendingTokens.Count;
                int maxTokensPerTick;
                int maxCharsPerTick;
                if (pendingBefore >= AnswerStreamUiHardBacklogTokens)
                {
                    maxTokensPerTick = AnswerStreamUiTokensPerTick_HardCatchUp;
                    maxCharsPerTick = AnswerStreamUiCharsPerTick_HardCatchUp;
                }
                else if (pendingBefore >= AnswerStreamUiSoftBacklogTokens)
                {
                    maxTokensPerTick = AnswerStreamUiTokensPerTick_CatchUp;
                    maxCharsPerTick = AnswerStreamUiCharsPerTick_CatchUp;
                }
                else
                {
                    maxTokensPerTick = AnswerStreamUiTokensPerTick_Steady;
                    maxCharsPerTick = AnswerStreamUiCharsPerTick_Steady;
                }

                var sb = new StringBuilder(capacity: 256);
                var tokens = 0;
                while (_answerStreamLivePendingTokens.Count > 0)
                {
                    var next = _answerStreamLivePendingTokens.Peek();
                    if (tokens >= maxTokensPerTick)
                        break;
                    if ((sb.Length + next.Length) > maxCharsPerTick && sb.Length > 0)
                        break;
                    _answerStreamLivePendingTokens.Dequeue();
                    sb.Append(next);
                    tokens++;
                }
                append = sb.ToString();
            }
        }

        if (!string.IsNullOrEmpty(append))
        {
            _streamingAnswerRun.Text += append;
            if (_answerAutoScrollEnabled)
                ScrollInterviewMiddleToEnd();
        }

        if (shouldFinalize)
        {
            _answerStreamFlushTimer?.Stop();
            _answerStreamLiveRendering = false;
            LogInterviewUiStatus("Answer ready.");
            RenderAiAnswer(
                _answerStreamCompletionContent,
                _answerStreamCompletionHeading,
                appendContinuation: _answerStreamCompletionAppendContinuation);
            ScrollInterviewMiddleToEnd();
            return;
        }

        if (App.Settings.StreamDebugLogs)
        {
            _answerStreamDebugUiFlushCount++;
            var nowTicks = Stopwatch.GetTimestamp();
            var last = Interlocked.Read(ref _answerStreamLastDebugUiLogTicks);
            // Log at most ~2x/second to avoid log spam.
            if (last == 0 || (nowTicks - last) > Stopwatch.Frequency / 2)
            {
                Interlocked.Exchange(ref _answerStreamLastDebugUiLogTicks, nowTicks);
                int pending;
                lock (_answerStreamLiveTextLock)
                    pending = _answerStreamLivePendingTokens.Count;
                DesktopLogger.Info($"[STREAM:UI] flushCount={_answerStreamDebugUiFlushCount} renderedChars={(_streamingAnswerRun.Text?.Length ?? 0)} appendedChars={append.Length} pendingTokens={pending} autoScroll={_answerAutoScrollEnabled}");
            }
        }
    }

    private void TryFlushAnswerStreamLiveTextToRun()
    {
        if (_streamingAnswerRun == null)
            return;
        string append;
        lock (_answerStreamLiveTextLock)
        {
            if (_answerStreamLivePendingTokens.Count == 0)
                return;
            var sb = new StringBuilder(capacity: 256);
            while (_answerStreamLivePendingTokens.Count > 0)
                sb.Append(_answerStreamLivePendingTokens.Dequeue());
            append = sb.ToString();
        }

        if (!string.IsNullOrEmpty(append))
            _streamingAnswerRun.Text += append;
    }


    private sealed class AltQuestionStreamFilter
    {
        // Stream includes a trailing block starting with this marker; we should not render it live.
        private const string Marker = "<<<ALT_Q>>>";
        private const int TailKeep = 16; // >= Marker.Length - 1 (9); keep a bit extra to be safe.

        private readonly StringBuilder _tail = new();
        private bool _markerSeen;

        public void Reset()
        {
            _tail.Clear();
            _markerSeen = false;
        }

        public bool MarkerSeen => _markerSeen;

        public string Feed(string delta)
        {
            if (_markerSeen || string.IsNullOrEmpty(delta))
                return string.Empty;

            // We intentionally delay emitting the last few characters so we can detect Marker across chunk boundaries.
            var combined = _tail.Length == 0 ? delta : _tail.ToString() + delta;
            var idx = combined.IndexOf(Marker, StringComparison.OrdinalIgnoreCase);
            if (idx >= 0)
            {
                _markerSeen = true;
                _tail.Clear();
                return idx == 0 ? string.Empty : combined[..idx];
            }

            if (combined.Length <= TailKeep)
            {
                _tail.Clear();
                _tail.Append(combined);
                return string.Empty;
            }

            var emitLen = combined.Length - TailKeep;
            var emit = combined[..emitLen];
            _tail.Clear();
            _tail.Append(combined.AsSpan(emitLen));
            return emit;
        }

        public string FlushRemainder()
        {
            if (_markerSeen || _tail.Length == 0)
                return string.Empty;
            var s = _tail.ToString();
            _tail.Clear();
            return s;
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

            // If user is returning from an external web flow (pricing/dashboard),
            // refresh credits immediately so "Buy credits" → "Full session" updates without a restart.
            if (_pendingCreditsRefreshOnActivate)
            {
                _pendingCreditsRefreshOnActivate = false;
                _ = RefreshCreditsFromServerAsync(force: true, showLoading: false);
            }
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
        RootChromeBorder.Opacity = 1.0;
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
            Dispatcher.BeginInvoke(new Action(PositionWindowTopMiddle), DispatcherPriority.Loaded);
        }
    }

    private void PositionWindowTopMiddle()
    {
        try
        {
            var wa = SystemParameters.WorkArea;
            UpdateLayout();

            var w = ActualWidth > 0 ? ActualWidth : Width;
            var h = ActualHeight > 0 ? ActualHeight : Height;
            if (double.IsNaN(w) || w <= 0) w = 820;
            if (double.IsNaN(h) || h <= 0) h = 260;

            const double topMargin = 10;
            var left = wa.Left + (wa.Width - w) / 2.0;
            var top = wa.Top + topMargin;

            Left = Math.Max(wa.Left, Math.Min(left, wa.Right - w));
            Top = Math.Max(wa.Top, Math.Min(top, wa.Bottom - h));
        }
        catch
        {
            /* ignore */
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
        RootChromeBorder.Opacity = 1.0;
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
        RootChromeBorder.Opacity = 1.0;
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
            _pendingCreditsRefreshOnActivate = true;
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

    private void OpenWebAppPricing()
    {
        var targetUrl = $"{GetWebAppOrigin().TrimEnd('/')}/#pricing";
        try
        {
            _pendingCreditsRefreshOnActivate = true;
            MinimizeToRestoreChip();
            Process.Start(new ProcessStartInfo { FileName = targetUrl, UseShellExecute = true });
            DesktopLogger.Info($"Opened web pricing: {targetUrl}");
        }
        catch (Exception ex)
        {
            DesktopLogger.Error($"Failed to open web pricing: {ex}");
            MessageBox.Show("Unable to open the website.", "Pricing", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OpenPastSessionsInBrowser()
    {
        var targetUrl = $"{GetWebAppOrigin().TrimEnd('/')}/dashboard/call-sessions";

        try
        {
            _pendingCreditsRefreshOnActivate = true;
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
            _naturalSpeakingMode = false;

            if (Guid.TryParse(App.Settings.CallSessionId, out var sid))
                _callSessionId = sid;
            else
                _callSessionId = Guid.Parse("AB589C99-0980-4467-8AF5-ADAB340FE1A0");

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
        RootChromeBorder.Opacity = 1.0;
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
        RootChromeBorder.Opacity = 1.0;
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

        if (!_isFreeSessionFlow && !CreditsState.Current.HasSufficientCreditsForPaidActivation())
        {
            MessageBox.Show(
                "You need at least 0.5 interview credit to create a full session. Buy credits on the website.",
                "Create Session",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            OpenWebAppPricing();
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
                NaturalSpeakingMode = CreateSessionStep2View.NaturalSpeakingMode,
                ExtraContext = CreateSessionStep2View.ExtraContext,
                AiModel = null,
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
            SetSessionExtraContext(payload.ExtraContext);
            _naturalSpeakingMode = created.NaturalSpeakingMode || payload.NaturalSpeakingMode;
            DesktopLogger.Info($"Create session success sessionId={_callSessionId} isFreeSession={payload.IsFreeSession} naturalSpeaking={_naturalSpeakingMode}");
            if (created.ResumeId.HasValue)
                _ = LoadResumeContextAsync(created.ResumeId.Value);

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
        ActivateSessionView.PrepareForDisplay(!_isFreeSessionFlow);

        ResizeMode = ResizeMode.NoResize;
        StartupLoginView.Visibility = Visibility.Collapsed;
        SessionSetupView.Visibility = Visibility.Collapsed;
        PastSessionsView.Visibility = Visibility.Collapsed;
        CreateSessionDetailsView.Visibility = Visibility.Collapsed;
        CreateSessionStep2View.Visibility = Visibility.Collapsed;
        ActivateSessionView.Visibility = Visibility.Visible;
        MainContentGrid.Visibility = Visibility.Collapsed;
        RootChromeBorder.Background = Brushes.Transparent;
        RootChromeBorder.Opacity = 1.0;
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
        // Start compact; then auto-fit to toolbar content after measure.
        Width = Math.Min(820, SystemParameters.WorkArea.Width - 24);
        // Outer window chrome is transparent; inner cards provide the UI surface.
        RootChromeBorder.Padding = new Thickness(0);
        RootChromeBorder.CornerRadius = new CornerRadius(0);
        RootChromeBorder.BorderThickness = new Thickness(0);
        // Window stays fully opaque; interview panel fade is RootChromeBorder.Opacity (see ApplyChromeTranslucencyFromSlider).
        Opacity = 1.0;
        ApplyChromeTranslucencyFromSlider();
        TrySetDwmBorderColor(0x00ECE7E5);

        Dispatcher.BeginInvoke(new Action(CoerceMainWindowWidthToToolbar), DispatcherPriority.Loaded);
    }

    private void CoerceMainWindowWidthToToolbar()
    {
        try
        {
            if (RootChromeBorder == null)
                return;

            // Measure the full toolbar row (scroll strip + fixed timer + close), not only the scrollable StackPanel,
            // so auto-width never hides the session timer.
            var row = TopToolbarRowGrid;
            if (row != null)
            {
                row.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                var desiredRow = row.DesiredSize.Width;
                if (desiredRow > 0)
                {
                    var pad = RootChromeBorder.Padding.Left + RootChromeBorder.Padding.Right;
                    var cardPad = 0.0;
                    if (TopToolbarCardBorder != null)
                        cardPad = TopToolbarCardBorder.Padding.Left + TopToolbarCardBorder.Padding.Right;
                    var chrome = 28; // border + small margin
                    var target = Math.Ceiling(desiredRow + pad + cardPad + chrome);

                    var max = SystemParameters.WorkArea.Width - 24;
                    target = Math.Min(target, max);
                    target = Math.Max(target, MinWidth);

                    Width = target;
                    if (_mainUiPositionInitialized)
                        PositionWindowTopMiddle();
                    return;
                }
            }

            if (TopToolbarContent == null)
                return;

            TopToolbarContent.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            var desired = TopToolbarContent.DesiredSize.Width;
            if (desired <= 0)
                return;

            var pad2 = RootChromeBorder.Padding.Left + RootChromeBorder.Padding.Right;
            var cardPad2 = 0.0;
            if (TopToolbarCardBorder != null)
                cardPad2 = TopToolbarCardBorder.Padding.Left + TopToolbarCardBorder.Padding.Right;
            var chrome2 = 28;
            var target2 = Math.Ceiling(desired + pad2 + cardPad2 + chrome2);

            var max2 = SystemParameters.WorkArea.Width - 24;
            target2 = Math.Min(target2, max2);
            target2 = Math.Max(target2, MinWidth);

            Width = target2;
            if (_mainUiPositionInitialized)
                PositionWindowTopMiddle();
        }
        catch
        {
            /* ignore */
        }
    }

    /// <summary>
    /// Maps the opacity slider (55–100%) to <see cref="Border.Opacity"/> on <see cref="RootChromeBorder"/> so the
    /// panel fill, border, buttons, text, and nested controls all fade uniformly. Brushes use full alpha; element
    /// opacity handles translucency (no separate “background only” path).
    /// </summary>
    private void ApplyChromeTranslucencyFromSlider()
    {
        if (!_mainUiStarted || RootChromeBorder == null) return;

        var pct = WindowOpacitySlider?.Value ?? 90;
        // Keep the outer host transparent; apply opacity to the two visible cards only.
        RootChromeBorder.Opacity = 1.0;
        RootChromeBorder.Background = Brushes.Transparent;
        RootChromeBorder.BorderBrush = Brushes.Transparent;

        var cardOpacity = Math.Clamp(pct / 100.0, 0, 1);
        var fullyOpaque = pct >= 99.5;

        if (TopToolbarCardBorder != null)
        {
            TopToolbarCardBorder.Opacity = cardOpacity;
            // At 100% opacity, also remove semi-transparent fill so the card looks truly solid.
            TopToolbarCardBorder.Background = fullyOpaque
                ? new SolidColorBrush(Color.FromRgb(0x12, 0x12, 0x1C))
                : new SolidColorBrush(Color.FromArgb(0xB0, 0x12, 0x12, 0x1C));
        }

        if (AnswerCardBorder != null)
        {
            AnswerCardBorder.Opacity = cardOpacity;
            AnswerCardBorder.Background = fullyOpaque
                ? new SolidColorBrush(Color.FromRgb(0x12, 0x12, 0x1C))
                : new SolidColorBrush(Color.FromArgb(0xB0, 0x12, 0x12, 0x1C));
        }
    }

    private void WindowOpacitySlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (!_mainUiStarted) return;
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
        if (!_mainUiPositionInitialized)
        {
            _mainUiPositionInitialized = true;
            // Do a top-middle placement after initial layout, and again after the toolbar measures and adjusts Width.
            Dispatcher.BeginInvoke(new Action(PositionWindowTopMiddle), DispatcherPriority.Loaded);
            Dispatcher.BeginInvoke(new Action(PositionWindowTopMiddle), DispatcherPriority.Background);
        }
        await EnsureSpeechConfiguredAsync().ConfigureAwait(true);
        if (_resumeId.HasValue)
        {
            _ = LoadResumeContextAsync(_resumeId.Value);
        }
    }

    private async void StartInterviewSessionAsync()
    {
        await StartMainUiIfNeededAsync().ConfigureAwait(true);

        if (!_isFreeSessionFlow && !CreditsState.Current.HasSufficientCreditsForPaidActivation())
        {
            MessageBox.Show(
                "You need at least 0.5 interview credit to activate a full session. Buy credits on the website.",
                "Activate Session",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            OpenWebAppPricing();
            return;
        }

        await RefreshSessionExtraContextFromServerAsync().ConfigureAwait(true);

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
        LogInterviewUiStatus($"Session started ({mode}).");

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
                // Do not auto-expand the answer region; user opens it via AI Answer (or another answer action).
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

    private void RegisterSuccessfulAnswerInHistory(string? completion, bool mergeIntoLatest = false)
    {
        if (string.IsNullOrWhiteSpace(completion))
            return;

        if (mergeIntoLatest && _sessionAnswerHistory.Count > 0)
        {
            var last = _sessionAnswerHistory[^1];
            last.Content = last.Content.TrimEnd() + "\n\n──────────\n\n" + completion.TrimEnd();
            if (!string.IsNullOrWhiteSpace(_currentAnswerDisplayHeading))
            {
                last.Heading = string.IsNullOrWhiteSpace(last.Heading)
                    ? _currentAnswerDisplayHeading
                    : last.Heading + "\n\n" + _currentAnswerDisplayHeading;
            }

            _lastAppendedAnswerHistoryIndex = _sessionAnswerHistory.Count - 1;
            _answerHistoryViewIndex = _lastAppendedAnswerHistoryIndex;
            return;
        }

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
        var raw = item.Content ?? string.Empty;
        var altQs = ExtractAltQuestions(raw);
        var clean = StripAltQuestionsForDisplay(raw).TrimEnd();
        RenderAiAnswer(raw, item.Heading);
        PopulateFollowUpSuggestionsFromAltQuestions(altQs, item.Heading, clean, TranscriptTextBlock.Text);
        _answerAutoScrollEnabled = false;
        ScrollInterviewMiddleToHome();

        _showAskAiAnythingRow = clean.Length > 12
                                && !clean.TrimStart().StartsWith("Error:", StringComparison.OrdinalIgnoreCase);
        UpdateAskAiAnythingRowVisibility();
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
                LogInterviewUiStatus($"Session extended by 30 minutes (0.5 credit used). Extensions: {_fullExtensionsApplied}.");
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
            {
                _callSessionId = Guid.Empty;
                SetSessionExtraContext(null);
                _naturalSpeakingMode = false;
            }

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

            LogInterviewUiStatus(reason);
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
                LogInterviewUiStatus($"End session error: {ex.Message}");
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
                LogInterviewUiStatus("Speech setup failed. Check Azure Speech key/region or server token endpoint.");
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
                        LogInterviewUiStatus("Speech is not configured (local or server).");
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
                    LogInterviewUiStatus("Listening…");
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
                    LogInterviewUiStatus($"Speech canceled: {e.Reason}");
                });
            };

            _speechRecognizer.SessionStopped += (sender, e) =>
            {
                Dispatcher.Invoke(() =>
                {
                    LogInterviewUiStatus("Speech session stopped.");
                    _isListening = false;
                });
            };

            await _speechRecognizer.StartContinuousRecognitionAsync().ConfigureAwait(false);
            _isListening = true;

            await Dispatcher.InvokeAsync(() =>
            {
                LogInterviewUiStatus("Listening…");
            });

            // 2) System audio (loopback) recognizer
            await InitializeSystemAudioSpeechAsync(config).ConfigureAwait(false);
            return _speechRecognizer != null && _isListening;
        }
        catch (Exception ex)
        {
            await Dispatcher.InvokeAsync(() =>
            {
                LogInterviewUiStatus($"Speech init error: {ex.Message}");
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
                    LogInterviewUiStatus($"Resume context load failed: {(int)response.StatusCode} {response.ReasonPhrase}");
                });
                return;
            }

            var text = await response.Content.ReadAsStringAsync();
            _resumeContext = string.IsNullOrWhiteSpace(text) ? null : text;
            DesktopLogger.Info($"Resume context loaded, length={_resumeContext?.Length ?? 0}");
            Dispatcher.Invoke(() =>
            {
                LogInterviewUiStatus($"Resume context loaded (resume {resumeId}).");
            });
        }
        catch (Exception ex)
        {
            DesktopLogger.Error($"LoadResumeContextAsync exception: {ex}");
            Dispatcher.Invoke(() =>
            {
                LogInterviewUiStatus($"Resume context error: {ex.Message}");
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
                    LogInterviewUiStatus($"System audio canceled: {e.Reason}");
                });
            };

            await _systemSpeechRecognizer.StartContinuousRecognitionAsync().ConfigureAwait(false);

            StartLoopbackCapture();
        }
        catch (Exception ex)
        {
            Dispatcher.Invoke(() =>
            {
                LogInterviewUiStatus($"System audio init error: {ex.Message}");
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

        // Detailed transcript: insert newlines when there was a pause between recognized chunks,
        // or after strong sentence-ending punctuation.
        var now = DateTimeOffset.UtcNow;
        var gap = _lastTranscriptAppendUtc == DateTimeOffset.MinValue ? TimeSpan.Zero : (now - _lastTranscriptAppendUtc);
        _lastTranscriptAppendUtc = now;

        bool shouldNewLine = gap > TimeSpan.FromSeconds(2);
        if (!shouldNewLine && _finalTranscriptDetail.Length > 0)
        {
            var last = _finalTranscriptDetail[^1];
            if (last is '.' or '!' or '?')
                shouldNewLine = true;
        }

        if (_finalTranscriptDetail.Length > 0)
            _finalTranscriptDetail.Append(shouldNewLine ? "\n" : " ");
        _finalTranscriptDetail.Append(t);

        const int maxChars = 6000;
        if (_finalTranscript.Length > maxChars)
        {
            var trimmed = _finalTranscript.ToString()[^maxChars..].TrimStart();
            _finalTranscript.Clear();
            _finalTranscript.Append(trimmed);
        }

        const int maxDetailChars = 20000;
        if (_finalTranscriptDetail.Length > maxDetailChars)
        {
            var trimmed = _finalTranscriptDetail.ToString()[^maxDetailChars..];
            var cut = trimmed.IndexOf('\n');
            if (cut > 0)
                trimmed = trimmed[(cut + 1)..];
            _finalTranscriptDetail.Clear();
            _finalTranscriptDetail.Append(trimmed.TrimStart());
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

        var full = sb.ToString();
        TranscriptTextBlock.Text = FitTranscriptTailToWidth(full);

        UpdateTranscriptDetailDisplay();
    }

    private void UpdateTranscriptDetailDisplay()
    {
        try
        {
            if (TranscriptDetailTextBox == null)
                return;

            var sb = new StringBuilder();
            sb.Append(_finalTranscriptDetail);

            void appendSpaceIfNeeded()
            {
                if (sb.Length == 0) return;
                if (sb[^1] is '\n' or '\r' || char.IsWhiteSpace(sb[^1])) return;
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

            TranscriptDetailTextBox.Text = sb.ToString().Trim();

            if (_transcriptDetailAutoScroll && TranscriptDetailCard?.Visibility == Visibility.Visible && TranscriptDetailScrollViewer != null)
                TranscriptDetailScrollViewer.ScrollToEnd();
        }
        catch
        {
            /* ignore */
        }
    }

    private void TranscriptExpandButton_Click(object sender, RoutedEventArgs e)
    {
        e.Handled = true;
        if (TranscriptDetailCard == null || TranscriptLeftColumn == null || TranscriptRightColumn == null)
            return;

        if (TranscriptDetailCard.Visibility == Visibility.Visible)
        {
            HideTranscriptDetailCard();
            return;
        }

        // Decide left/right based on where the app window is on the screen.
        try
        {
            var wa = SystemParameters.WorkArea;
            var winCenter = Left + (ActualWidth > 0 ? ActualWidth : Width) / 2.0;
            var workCenter = wa.Left + wa.Width / 2.0;
            // If window is on left half → dock transcript on right. Otherwise dock on left.
            DockTranscriptDetailCardRight(winCenter < workCenter);
        }
        catch
        {
            DockTranscriptDetailCardRight(false);
        }

        UpdateTranscriptDetailDisplay();
        ShowTranscriptDetailCard();
    }

    private void TranscriptAutoScrollButton_Click(object sender, RoutedEventArgs e)
    {
        e.Handled = true;
        _transcriptDetailAutoScroll = !_transcriptDetailAutoScroll;
        if (TranscriptAutoScrollButton != null)
            TranscriptAutoScrollButton.Content = _transcriptDetailAutoScroll ? "Auto scroll: On" : "Auto scroll: Off";
        if (_transcriptDetailAutoScroll && TranscriptDetailScrollViewer != null)
            TranscriptDetailScrollViewer.ScrollToEnd();
    }

    private void TranscriptCollapseButton_Click(object sender, RoutedEventArgs e)
    {
        e.Handled = true;
        HideTranscriptDetailCard();
    }

    private void ShowTranscriptDetailCard()
    {
        if (TranscriptDetailCard == null || TranscriptLeftColumn == null || TranscriptRightColumn == null)
            return;

        if (!_transcriptDetailShown)
        {
            _transcriptDetailShown = true;
            _baseWidthBeforeTranscript = Width;
            _baseLeftBeforeTranscript = Left;
            _baseMainCardsColumnWidth = (TopToolbarCardBorder?.ActualWidth > 0 ? TopToolbarCardBorder.ActualWidth : (AnswerCardBorder?.ActualWidth ?? 0));
            if (_baseMainCardsColumnWidth <= 0 && MainCardsColumn != null && MainCardsColumn.ActualWidth > 0)
                _baseMainCardsColumnWidth = MainCardsColumn.ActualWidth;
        }

        const double cardWidth = 320;
        const double gap = 10;
        var added = cardWidth + gap;

        // Expand the window so the answer card keeps its width, and keep the main cards anchored
        // (do not shift the existing two cards on screen when showing the left transcript card).
        try
        {
            var wa = SystemParameters.WorkArea;
            var w = Math.Max(_baseWidthBeforeTranscript, ActualWidth > 0 ? ActualWidth : Width);
            var newWidth = Math.Min(wa.Width - 12, w + added);
            Width = newWidth;

            if (!_transcriptDockedRight)
            {
                // Transcript on LEFT: move window left by the added width (clamped) so main cards stay in place.
                var targetLeft = _baseLeftBeforeTranscript - added;
                var minLeft = wa.Left + 6;
                var maxLeft = wa.Right - newWidth - 6;
                Left = Math.Max(minLeft, Math.Min(targetLeft, maxLeft));
            }
        }
        catch
        {
            Width = _baseWidthBeforeTranscript + added;
        }

        // Freeze the main cards column width while transcript is visible.
        if (MainCardsColumn != null && _baseMainCardsColumnWidth > 0)
            MainCardsColumn.Width = new GridLength(_baseMainCardsColumnWidth);

        TranscriptDetailCard.Visibility = Visibility.Visible;
        UpdateTranscriptDetailDisplay();
    }

    private void HideTranscriptDetailCard()
    {
        if (TranscriptDetailCard == null || TranscriptLeftColumn == null || TranscriptRightColumn == null)
            return;
        TranscriptDetailCard.Visibility = Visibility.Collapsed;
        TranscriptLeftColumn.Width = new GridLength(0);
        TranscriptRightColumn.Width = new GridLength(0);

        if (_transcriptDetailShown)
        {
            _transcriptDetailShown = false;
            try
            {
                Width = _baseWidthBeforeTranscript;
                Left = _baseLeftBeforeTranscript;
            }
            catch
            {
                /* ignore */
            }
        }

        // Restore main cards column sizing.
        if (MainCardsColumn != null)
            MainCardsColumn.Width = new GridLength(1, GridUnitType.Star);
    }

    private void DockTranscriptDetailCardRight(bool dockRight)
    {
        if (TranscriptDetailCard == null || TranscriptLeftColumn == null || TranscriptRightColumn == null)
            return;
        _transcriptDockedRight = dockRight;
        const double cardWidth = 320;
        if (dockRight)
        {
            Grid.SetColumn(TranscriptDetailCard, 4);
            TranscriptRightColumn.Width = new GridLength(cardWidth);
            TranscriptLeftColumn.Width = new GridLength(0);
        }
        else
        {
            Grid.SetColumn(TranscriptDetailCard, 0);
            TranscriptLeftColumn.Width = new GridLength(cardWidth);
            TranscriptRightColumn.Width = new GridLength(0);
        }
    }

    private string FitTranscriptTailToWidth(string full)
    {
        var t = (full ?? string.Empty).Replace("\r\n", " ").Replace("\n", " ");
        while (t.Contains("  ", StringComparison.Ordinal))
            t = t.Replace("  ", " ", StringComparison.Ordinal);
        t = t.Trim();
        if (t.Length == 0)
            return string.Empty;

        // Fallback until we have layout.
        if (TranscriptTextBlock == null || TranscriptTextBlock.ActualWidth <= 20)
        {
            const int uiTailChars = 180;
            return t.Length > uiTailChars ? t[^uiTailChars..].TrimStart() : t;
        }

        // Available width minus padding.
        var pad = TranscriptTextBlock.Padding.Left + TranscriptTextBlock.Padding.Right;
        var available = Math.Max(40, TranscriptTextBlock.ActualWidth - pad - 12);

        var typeface = new Typeface(
            TranscriptTextBlock.FontFamily,
            TranscriptTextBlock.FontStyle,
            TranscriptTextBlock.FontWeight,
            TranscriptTextBlock.FontStretch);

        double measure(string s)
        {
            var ft = new FormattedText(
                s,
                CultureInfo.CurrentUICulture,
                FlowDirection.LeftToRight,
                typeface,
                TranscriptTextBlock.FontSize,
                Brushes.Black,
                VisualTreeHelper.GetDpi(this).PixelsPerDip);
            return ft.WidthIncludingTrailingWhitespace;
        }

        // If whole string fits, keep it.
        if (measure(t) <= available)
            return t;

        // Binary search for the smallest start index that fits (show the newest tail).
        int lo = 0, hi = t.Length - 1;
        while (lo < hi)
        {
            var mid = (lo + hi) / 2;
            var candidate = t[mid..].TrimStart();
            if (candidate.Length == 0)
            {
                lo = mid + 1;
                continue;
            }

            if (measure(candidate) <= available)
                hi = mid;
            else
                lo = mid + 1;
        }

        var result = t[lo..].TrimStart();
        // Add a subtle leading ellipsis to indicate clipping.
        if (result.Length > 0 && lo > 0)
            result = "… " + result;
        return result;
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        var helper = new System.Windows.Interop.WindowInteropHelper(this);
        var handle = helper.Handle;

        SetWindowDisplayAffinity(
            handle,
            ExcludeMainWindowFromScreenCapture ? WDA_EXCLUDEFROMCAPTURE : WDA_NONE);

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
        TryLoadToolbarLogo();
    }

    private void TryLoadToolbarLogo()
    {
        if (ToolbarLogoImage == null) return;

        foreach (var uriString in new[]
                 {
                     "pack://siteoforigin:,,,/assets/smeed-logo.png",
                     "pack://application:,,,/assets/smeed-logo.png"
                 })
        {
            try
            {
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.UriSource = new Uri(uriString, UriKind.Absolute);
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.EndInit();
                if (bmp.CanFreeze)
                    bmp.Freeze();
                ToolbarLogoImage.Source = bmp;
                return;
            }
            catch
            {
                /* try next uri */
            }
        }
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

    private void HideChip_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
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
                MicToggleButton.Background = new SolidColorBrush(Color.FromRgb(0x4F, 0x46, 0xE5));
                MicToggleButton.Foreground = new SolidColorBrush(Color.FromRgb(0xEE, 0xF2, 0xFF));
                MicToggleButton.BorderBrush = new SolidColorBrush(Color.FromRgb(0x81, 0x8C, 0xF8));
                LogInterviewUiStatus("Mic on.");
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
                LogInterviewUiStatus("Mic off.");
            }
        }
        catch (Exception ex)
        {
            _micOn = !_micOn;
            LogInterviewUiStatus($"Mic error: {ex.Message}");
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
                LogInterviewUiStatus("Speaker (computer audio) on.");
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
                LogInterviewUiStatus("Speaker off.");
            }
        }
        catch (Exception ex)
        {
            _speakerOn = !_speakerOn;
            LogInterviewUiStatus($"Speaker error: {ex.Message}");
        }
    }

    private async void AiAnswerButton_Click(object sender, RoutedEventArgs e)
    {
        if (_answerGenerationInFlight)
        {
            LogInterviewUiStatus("Wait for the current answer to finish, then try AI Answer again.");
            return;
        }

        _ = IncrementAiUsageAsync();
        var transcript = (TranscriptTextBlock.Text ?? "").Trim();
        if (string.IsNullOrWhiteSpace(transcript))
        {
            LogInterviewUiStatus("Transcript is empty. Speak or paste text first.");
            return;
        }

        AiAnswerButton.IsEnabled = false;

        var question = QuestionTextBox.Text?.Trim();
        var startedAnswerStream = false;
        try
        {
            await Dispatcher.InvokeAsync(() =>
            {
                LogInterviewUiStatus("Understanding question…");
            });

            const int clarifyMax = 12_000;
            var snippet = transcript.Length <= clarifyMax
                ? transcript
                : transcript[^clarifyMax..].TrimStart();

            var clarified = await TryClarifyTranscriptQuestionAsync(
                snippet,
                BuildClarifyUserContext(question, _sessionExtraContext)).ConfigureAwait(false);

            string latestQuestion;
            string headingForUi;
            if (clarified.HasValue)
            {
                var cleanedHeading = CleanClarifiedQuestionLine(clarified.Value.Heading);
                var cleanedBody = CleanClarifiedBodyLines(clarified.Value.Body);
                latestQuestion = string.IsNullOrWhiteSpace(cleanedBody) ? cleanedHeading : cleanedBody;
                headingForUi = string.IsNullOrWhiteSpace(question)
                    ? SummarizeForAnswerHeading(cleanedHeading)
                    : question.Trim();
            }
            else
            {
                latestQuestion = GetLatestQuestionFromTranscript(transcript);
                if (string.IsNullOrWhiteSpace(latestQuestion))
                {
                    await Dispatcher.InvokeAsync(() =>
                    {
                        LogInterviewUiStatus("Transcript is empty. Speak or paste text first.");
                    });
                    return;
                }

                latestQuestion = CleanClarifiedQuestionLine(latestQuestion);
                headingForUi = string.IsNullOrWhiteSpace(question)
                    ? SummarizeForAnswerHeading(latestQuestion)
                    : question.Trim();
            }

            if (string.IsNullOrWhiteSpace(latestQuestion))
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    LogInterviewUiStatus("Could not detect a question in the transcript.");
                });
                return;
            }

            var userContent = string.IsNullOrWhiteSpace(question)
                ? $"Answer this question (from voice transcription):\n{latestQuestion}"
                : $"Answer this question. Context from user: {question}.\nTranscribed question (cleaned when possible):\n{latestQuestion}";

            await Dispatcher.InvokeAsync(ShowAnswerSectionUi);
            startedAnswerStream = true;
            await GetAnswerAsync(userContent, AnswerFromTranscriptSystemPrompt, headingForUi).ConfigureAwait(false);
        }
        finally
        {
            if (!startedAnswerStream)
                AiAnswerButton.IsEnabled = true;
        }
    }

    private async void AskButton_Click(object sender, RoutedEventArgs e)
    {
        await SubmitQuestionFromInputAsync().ConfigureAwait(false);
    }

    private async Task SubmitQuestionFromInputAsync(bool continueConversation = false)
    {
        var question = QuestionTextBox.Text?.Trim();
        if (string.IsNullOrWhiteSpace(question))
        {
            LogInterviewUiStatus("Type a question first.");
            return;
        }

        await SubmitQuestionAsync(question, continueConversation).ConfigureAwait(false);
    }

    private async Task SubmitQuestionAsync(string question, bool continueConversation)
    {
        if (_answerGenerationInFlight) return;
        if (string.IsNullOrWhiteSpace(question)) return;

        _continueAnswerFromFollowUpChip = continueConversation;
        _ = IncrementAiUsageAsync();
        var typed = question.Trim();
        typed = ExtractLastSinglePrompt(typed);

        // When the user explicitly types a question here, ALWAYS answer that question.
        // The transcript can be included as optional context, but must not override the typed prompt.
        string fullTranscript;
        try
        {
            var sb = new StringBuilder();
            sb.Append(_finalTranscript);
            if (!string.IsNullOrWhiteSpace(_partialMic))
            {
                if (sb.Length > 0) sb.Append(' ');
                sb.Append(_partialMic.Trim());
            }
            if (!string.IsNullOrWhiteSpace(_partialSystem))
            {
                if (sb.Length > 0) sb.Append(' ');
                sb.Append(_partialSystem.Trim());
            }
            fullTranscript = sb.ToString();
        }
        catch
        {
            fullTranscript = (TranscriptTextBlock?.Text ?? string.Empty);
        }

        var transcriptTail = (fullTranscript ?? string.Empty).Trim();
        if (transcriptTail.Length > 3500)
            transcriptTail = transcriptTail[^3500..].TrimStart();

        // Always hard-anchor the model to ONE question only.
        var userContent =
            "Answer ONLY this question (ignore any other questions in transcript/history):\n" +
            $"{typed}\n\n" +
            "Optional context (may be unrelated; use ONLY if it helps):\n" +
            (string.IsNullOrWhiteSpace(transcriptTail) ? "(no transcript context)" : transcriptTail);

        // Clear input after capturing.
        await Dispatcher.InvokeAsync(() =>
        {
            try { QuestionTextBox?.Clear(); } catch { /* ignore */ }
        });

        await GetAnswerAsync(userContent, _systemPrompt, typed).ConfigureAwait(false);
    }

    /// <summary>
    /// Best-effort: when transcription or typed input contains multiple prompts in one line,
    /// keep ONLY the last prompt so we answer the latest question.
    /// Examples:
    /// - "write SQL ... and write C# ..." -> keep "write C# ..."
    /// - "find X. then reverse string" -> keep "reverse string"
    /// </summary>
    private static string ExtractLastSinglePrompt(string text)
    {
        var t = (text ?? string.Empty).Trim();
        if (t.Length == 0) return t;

        // Normalize whitespace to simplify splitting.
        t = t.Replace("\r\n", " ").Replace("\n", " ");
        while (t.Contains("  ", StringComparison.Ordinal))
            t = t.Replace("  ", " ", StringComparison.Ordinal);

        // If it contains multiple "task verbs", it's likely multiple prompts.
        var verbRx = new Regex(@"\b(write|find|implement|create|design|build|explain|describe|solve|reverse|query)\b",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        var verbCount = verbRx.Matches(t).Count;
        if (verbCount <= 1)
            return t.Trim();

        // Prefer splitting on obvious separators first.
        var parts = Regex.Split(t, @"(?i)(?:\.\s+|;\s+|\?\s+|!\s+|\bthen\b\s+|\balso\b\s+|\bnext\b\s+|\band\b\s+)")
            .Select(p => (p ?? string.Empty).Trim())
            .Where(p => p.Length >= 4)
            .ToList();
        if (parts.Count == 0)
            return t.Trim();

        // Pick the last part that still looks like an actionable prompt (has a verb).
        for (var i = parts.Count - 1; i >= 0; i--)
        {
            var p = parts[i];
            if (verbRx.IsMatch(p))
                return p.Trim();
        }

        return parts[^1].Trim();
    }

    private void MainWindow_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (MainContentGrid.Visibility != Visibility.Visible)
            return;

        if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.Oem2)
        {
            QuestionTextBox.Focus();
            QuestionTextBox.SelectAll();
            e.Handled = true;
        }
    }

    private void QuestionTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || Keyboard.Modifiers != ModifierKeys.None)
            return;

        e.Handled = true;
        _ = SubmitQuestionFromInputAsync(continueConversation: false);
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

    /// <summary>Matches greetings / filler often prepended by ASR (aligned with server clarify cleanup).</summary>
    private static readonly Regex LeadInNoiseRegex = new(
        @"^(?:(?:hi|hello|hey|hiya|yo)\b[,!.]?\s*|(?:good\s+(?:morning|afternoon|evening))\b[,!.]?\s*|" +
        @"(?:thanks?|thank\s+you)\b[,!.]?\s*|(?:ok+|okay|alright)\b[,!.]?\s*|(?:um+|uh+)\b[,!.]?\s*)+",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static string StripConversationalLeadInsFromStart(string text)
    {
        var t = (text ?? string.Empty).Trim();
        if (t.Length == 0)
            return t;
        string prev;
        do
        {
            prev = t;
            t = LeadInNoiseRegex.Replace(t, "").TrimStart();
        } while (!string.Equals(prev, t, StringComparison.Ordinal));

        return t.Trim();
    }

    private static string EnsureLeadingLetterUppercase(string text)
    {
        var t = (text ?? string.Empty).Trim();
        if (t.Length == 0)
            return t;
        var i = 0;
        while (i < t.Length && !char.IsLetter(t[i]))
            i++;
        if (i >= t.Length)
            return t;
        if (!char.IsLower(t[i]))
            return t;
        return t[..i] + char.ToUpperInvariant(t[i]) + t[(i + 1)..];
    }

    private static string CleanClarifiedQuestionLine(string line)
    {
        var t = StripConversationalLeadInsFromStart(line);
        return EnsureLeadingLetterUppercase(t).Trim();
    }

    private static string CleanClarifiedBodyLines(string body)
    {
        if (string.IsNullOrWhiteSpace(body))
            return string.Empty;
        var lines = body.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            if (line.Length == 0)
                continue;
            var numbered = Regex.Match(line, @"^(\d+\.\s*)");
            if (numbered.Success)
            {
                var prefix = numbered.Groups[1].Value;
                var rest = line[prefix.Length..];
                rest = StripConversationalLeadInsFromStart(rest);
                lines[i] = prefix + EnsureLeadingLetterUppercase(rest);
            }
            else
            {
                lines[i] = CleanClarifiedQuestionLine(line);
            }
        }

        return string.Join("\n", lines).Trim();
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

    /// <summary>
    /// Uses the API to extract and rephrase the latest question from noisy ASR text.
    /// Returns null on failure so callers can fall back to <see cref="GetLatestQuestionFromTranscript"/>.
    /// </summary>
    private async Task<(string Heading, string Body)?> TryClarifyTranscriptQuestionAsync(
        string transcriptSnippet,
        string? userContext)
    {
        try
        {
            var payload = new DesktopClarifyTranscriptQuestionRequest
            {
                Transcript = transcriptSnippet,
                UserContext = string.IsNullOrWhiteSpace(userContext) ? null : userContext.Trim(),
            };

            using var res = await _apiClient
                .PostAsJsonAsync("desktop/ai/clarify-transcript-question", payload)
                .ConfigureAwait(false);
            if (!res.IsSuccessStatusCode)
                return null;

            var dto = await res.Content
                .ReadFromJsonAsync<DesktopClarifyTranscriptQuestionResponse>(
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                .ConfigureAwait(false);
            if (dto == null || string.IsNullOrWhiteSpace(dto.Heading))
                return null;
            var h = dto.Heading.Trim();
            var b = string.IsNullOrWhiteSpace(dto.Body) ? h : dto.Body.Trim();
            return (h, b);
        }
        catch (Exception ex)
        {
            DesktopLogger.Warn($"TryClarifyTranscriptQuestionAsync: {ex.Message}");
            return null;
        }
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

    /// <summary>First line of model output: "💬 **Question**: ..." — styled in accent; do not duplicate via a separate UI heading.</summary>
    private static bool IsModelQuestionLine(string? line)
    {
        var t = (line ?? string.Empty).TrimStart();
        if (t.Length == 0) return false;
        if (t.StartsWith("💬", StringComparison.Ordinal)) return true;
        return t.StartsWith("**Question**", StringComparison.OrdinalIgnoreCase)
               || Regex.IsMatch(t, @"^\*\*Question\*\*\s*:", RegexOptions.IgnoreCase);
    }

    private static void AppendFormattedQuestionLineParagraph(FlowDocument doc, string line)
    {
        var emojiFont = AnswerEmojiFontFamily;
        var p = new Paragraph
        {
            Margin = new Thickness(0, 0, 0, 12),
            LineHeight = 24,
            Foreground = AnswerQuestionHeadingBrush,
            FontFamily = emojiFont
        };
        var text = line ?? string.Empty;
        if (!MarkdownEmphasisRegex.IsMatch(text))
        {
            p.Inlines.Add(new Run(text)
            {
                FontWeight = FontWeights.SemiBold,
                Foreground = AnswerQuestionHeadingBrush,
                FontSize = 14,
                FontFamily = emojiFont
            });
            doc.Blocks.Add(p);
            return;
        }

        var idx = 0;
        foreach (Match m in MarkdownEmphasisRegex.Matches(text))
        {
            if (m.Index > idx)
            {
                p.Inlines.Add(new Run(text.Substring(idx, m.Index - idx))
                {
                    FontWeight = FontWeights.SemiBold,
                    Foreground = AnswerQuestionHeadingBrush,
                    FontSize = 14,
                    FontFamily = emojiFont
                });
            }

            var inner = m.Groups["bold"].Success ? m.Groups["bold"].Value : m.Groups["mark"].Value;
            p.Inlines.Add(new Run(inner)
            {
                FontWeight = FontWeights.Bold,
                Foreground = AnswerQuestionHeadingBrush,
                FontSize = 14,
                FontFamily = emojiFont
            });
            idx = m.Index + m.Length;
        }

        if (idx < text.Length)
        {
            p.Inlines.Add(new Run(text[idx..])
            {
                FontWeight = FontWeights.SemiBold,
                Foreground = AnswerQuestionHeadingBrush,
                FontSize = 14,
                FontFamily = emojiFont
            });
        }

        doc.Blocks.Add(p);
    }

    private static string FormatCodeLanguageLabel(string? lang)
    {
        if (string.IsNullOrWhiteSpace(lang)) return "Code";
        var k = lang.Trim().ToLowerInvariant();
        return k switch
        {
            "cs" or "csharp" or "c#" => "C#",
            "fs" or "fsharp" => "F#",
            "vb" or "vbnet" => "VB.NET",
            "js" or "javascript" => "JavaScript",
            "ts" or "typescript" => "TypeScript",
            "py" or "python" => "Python",
            "sql" => "SQL",
            "json" => "JSON",
            "xml" => "XML",
            "html" or "htm" => "HTML",
            "css" => "CSS",
            "bash" or "sh" or "shell" => "Bash",
            "cpp" or "c++" or "cxx" => "C++",
            "c" => "C",
            "go" or "golang" => "Go",
            "rs" or "rust" => "Rust",
            "kotlin" or "kt" => "Kotlin",
            "java" => "Java",
            _ => k.Length == 0 ? "Code" : char.ToUpperInvariant(k[0]) + (k.Length > 1 ? k[1..] : "")
        };
    }

    private enum AnswerInlineEmphasis
    {
        None,
        Bold,
        Mark
    }

    private static void AppendFormattedAnswerParagraph(FlowDocument doc, string line)
    {
        var p = new Paragraph
        {
            Margin = new Thickness(0, 0, 0, 8),
            LineHeight = 22,
            Foreground = AnswerBodyBrush,
            FontFamily = AnswerEmojiFontFamily
        };
        var text = line ?? string.Empty;
        if (!MarkdownEmphasisRegex.IsMatch(text))
        {
            AppendGlossaryRuns(p, text, AnswerInlineEmphasis.None, 0);
            doc.Blocks.Add(p);
            return;
        }

        var idx = 0;
        var boldStyleOrdinal = 0;
        foreach (Match m in MarkdownEmphasisRegex.Matches(text))
        {
            if (m.Index > idx)
                AppendGlossaryRuns(p, text.Substring(idx, m.Index - idx), AnswerInlineEmphasis.None, 0);
            if (m.Groups["bold"].Success)
            {
                var styleIdx = boldStyleOrdinal++ % AnswerBoldEmphasisStyles.Length;
                AppendGlossaryRuns(p, m.Groups["bold"].Value, AnswerInlineEmphasis.Bold, styleIdx);
            }
            else if (m.Groups["mark"].Success)
                AppendGlossaryRuns(p, m.Groups["mark"].Value, AnswerInlineEmphasis.Mark, 0);
            idx = m.Index + m.Length;
        }

        if (idx < text.Length)
            AppendGlossaryRuns(p, text[idx..], AnswerInlineEmphasis.None, 0);

        doc.Blocks.Add(p);
    }

    /// <param name="boldStyleIndex">Palette index for <see cref="AnswerInlineEmphasis.Bold"/>; ignored for other modes.</param>
    private static void AppendGlossaryRuns(Paragraph p, string segment, AnswerInlineEmphasis emphasis, int boldStyleIndex)
    {
        if (segment.Length == 0) return;
        if (!InterviewGlossaryRegex.Value.IsMatch(segment))
        {
            p.Inlines.Add(CreateAnswerRun(segment, emphasis, false, boldStyleIndex));
            return;
        }

        var rx = InterviewGlossaryRegex.Value;
        var last = 0;
        foreach (Match m in rx.Matches(segment))
        {
            if (m.Index > last)
                p.Inlines.Add(CreateAnswerRun(segment[last..m.Index], emphasis, false, boldStyleIndex));
            AddGlossaryKeywordRun(p, m.Value);
            last = m.Index + m.Length;
        }

        if (last < segment.Length)
            p.Inlines.Add(CreateAnswerRun(segment[last..], emphasis, false, boldStyleIndex));
    }

    /// <summary>Inline glossary highlight using <see cref="Run"/> so baseline matches body text (no floating pill boxes).</summary>
    private static void AddGlossaryKeywordRun(Paragraph p, string term)
    {
        p.Inlines.Add(new Run(term)
        {
            FontFamily = AnswerEmojiFontFamily,
            FontSize = 13,
            FontWeight = FontWeights.SemiBold,
            Foreground = KeywordPillFgBrush,
            Background = GlossaryHighlightBgBrush
        });
    }

    private static Run CreateAnswerRun(string text, AnswerInlineEmphasis emphasis, bool glossary, int boldStyleIndex = 0)
    {
        var emojiFont = new FontFamily("Segoe UI, Segoe UI Emoji");
        if (glossary)
        {
            return new Run(text)
            {
                FontFamily = emojiFont,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = KeywordPillFgBrush,
                Background = GlossaryHighlightBgBrush
            };
        }

        return emphasis switch
        {
            AnswerInlineEmphasis.Bold => new Run(text)
            {
                FontFamily = emojiFont,
                FontWeight = FontWeights.Bold,
                FontSize = 14,
                Foreground = AnswerBoldEmphasisStyles[boldStyleIndex % AnswerBoldEmphasisStyles.Length].Fg,
                Background = AnswerBoldEmphasisStyles[boldStyleIndex % AnswerBoldEmphasisStyles.Length].Bg
            },
            AnswerInlineEmphasis.Mark => new Run(text)
            {
                FontFamily = emojiFont,
                FontWeight = FontWeights.SemiBold,
                FontSize = 14,
                Foreground = AnswerMarkFgBrush,
                Background = AnswerMarkBgBrush
            },
            _ => new Run(text)
            {
                FontFamily = emojiFont,
                FontWeight = FontWeights.Medium,
                FontSize = 14,
                Foreground = AnswerBodyBrush
            }
        };
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
            LogInterviewUiStatus(source == AudioQuestionSource.Interviewer
                ? "Auto answer: interviewer question detected..."
                : "Auto answer: question detected on mic...");

            await IncrementAiUsageAsync();

            var prefix = source == AudioQuestionSource.Interviewer
                ? "Answer this question (transcribed from the interviewer's audio). Focus only on the question(s) below:"
                : "Answer this question (transcribed from your microphone). Focus only on the question(s) below:";

            const int clarifyMax = 12_000;
            var snippet = trimmed.Length <= clarifyMax
                ? trimmed
                : trimmed[^clarifyMax..].TrimStart();
            var clarified = await TryClarifyTranscriptQuestionAsync(
                snippet,
                BuildClarifyUserContext(null, _sessionExtraContext)).ConfigureAwait(false);

            string userContent;
            string headingForUi;
            if (clarified.HasValue)
            {
                var cleanedHeading = CleanClarifiedQuestionLine(clarified.Value.Heading);
                var cleanedBody = CleanClarifiedBodyLines(clarified.Value.Body);
                var payload = string.IsNullOrWhiteSpace(cleanedBody) ? cleanedHeading : cleanedBody;
                userContent = $"{prefix}\n\n{payload}";
                headingForUi = SummarizeForAnswerHeading(cleanedHeading);
            }
            else
            {
                var segments = SplitTranscriptIntoQuestionSegments(trimmed);
                for (var i = 0; i < segments.Count; i++)
                    segments[i] = CleanClarifiedQuestionLine(segments[i]);
                userContent = BuildAutoAnswerPromptFromSegments(prefix, segments);
                headingForUi = FormatSplitQuestionsForAnswerHeading(segments);
            }

            await GetAnswerAsync(userContent, AnswerFromTranscriptSystemPrompt, headingForUi).ConfigureAwait(false);
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
        e.Handled = true;
        try
        {
            _finalTranscript.Clear();
            _finalTranscriptDetail.Clear();
            _partialMic = string.Empty;
            _partialSystem = string.Empty;
            _lastTranscriptAppendUtc = DateTimeOffset.MinValue;

            ResetAutoAnswerTransientState();
            UpdateTranscriptDisplay();
        }
        catch
        {
            /* ignore */
        }
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

        _showAskAiAnythingRow = false;
        UpdateAskAiAnythingRowVisibility();
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

        _showAskAiAnythingRow = false;
        UpdateAskAiAnythingRowVisibility();
        UpdateAiAnswerBodyMaxHeight();
        UpdateAnswerHistoryNav();
    }

    private void SetAnswerSectionRowHeight(bool collapsed)
    {
        if (AnswerCardRow == null || AnswerCardBorder == null || InterviewMiddleScrollViewer == null)
            return;

        if (collapsed)
        {
            AnswerCardRow.Height = new GridLength(0);
            AnswerCardRow.MinHeight = 0;
            AnswerCardBorder.Visibility = Visibility.Collapsed;
            HideTranscriptDetailCard();
            InterviewMiddleScrollViewer.Visibility = Visibility.Collapsed;
        }
        else
        {
            AnswerCardRow.Height = new GridLength(1, GridUnitType.Star);
            AnswerCardRow.MinHeight = 64;
            AnswerCardBorder.Visibility = Visibility.Visible;
            InterviewMiddleScrollViewer.Visibility = Visibility.Visible;
        }
    }

    private void ShowAnswerSectionUi()
    {
        if (FindName("AnswerSectionPanel") is not UIElement panel)
            return;
        panel.Visibility = Visibility.Visible;
        if (AnswerCardBorder != null)
            AnswerCardBorder.Visibility = Visibility.Visible;
        // As soon as the user opens the answer panel, show the "Ask AI anything" row
        // (even before the first answer finishes streaming).
        _showAskAiAnythingRow = true;
        UpdateAskAiAnythingRowVisibility();
        SetAnswerSectionRowHeight(collapsed: false);
    }

    private void UpdateAskAiAnythingRowVisibility()
    {
        if (AskAiAnythingRow == null) return;
        AskAiAnythingRow.Visibility = _showAskAiAnythingRow ? Visibility.Visible : Visibility.Collapsed;
    }

    private static string? NormalizeSessionExtraContext(string? value)
    {
        var t = (value ?? string.Empty).Trim();
        if (t.Length == 0) return null;
        return t.Length > 2000 ? t[..2000] : t;
    }

    private void SetSessionExtraContext(string? value) =>
        _sessionExtraContext = NormalizeSessionExtraContext(value);

    private static string? BuildClarifyUserContext(string? questionBoxText, string? sessionExtraContext)
    {
        var q = (questionBoxText ?? string.Empty).Trim();
        var e = NormalizeSessionExtraContext(sessionExtraContext);
        if (e == null && q.Length == 0) return null;
        if (e == null) return q.Length > 2000 ? q[..2000] : q;
        if (q.Length == 0)
            return "Session instructions from the candidate (apply when relevant):\n" + e;
        var q2 = q.Length > 2000 ? q[..2000] : q;
        return "Session instructions from the candidate (apply when relevant):\n" + e +
               "\n\nTyped note from the candidate (not from speech transcription):\n" + q2;
    }

    private async Task RefreshSessionExtraContextFromServerAsync()
    {
        if (_callSessionId == Guid.Empty)
        {
            SetSessionExtraContext(null);
            _naturalSpeakingMode = false;
            return;
        }

        try
        {
            using var res = await _apiClient.GetAsync($"callsessions/{_callSessionId}").ConfigureAwait(false);
            if (!res.IsSuccessStatusCode) return;
            var dto = await res.Content.ReadFromJsonAsync<CallSessionDto>(
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }).ConfigureAwait(false);
            if (dto != null)
            {
                SetSessionExtraContext(dto.ExtraContext);
                _naturalSpeakingMode = dto.NaturalSpeakingMode;
            }
        }
        catch (Exception ex)
        {
            DesktopLogger.Warn($"RefreshSessionExtraContextFromServerAsync: {ex.Message}");
        }
    }

    private string ComposeExtendedSystemPrompt(string? systemPrompt)
    {
        var basePrompt = string.IsNullOrWhiteSpace(systemPrompt) ? _systemPrompt : systemPrompt;
        var outputLanguage = GetOutputLanguageNameForPrompt(App.Settings.SessionLanguage);
        var extraBlock = string.IsNullOrWhiteSpace(_sessionExtraContext)
            ? string.Empty
            : "\n\nAdditional instructions:\n" + _sessionExtraContext +
              "\n\nFollow these Additional instructions on every answer (style, length, examples, tone), unless they would conflict with answering truthfully from the resume.";
        var naturalSpeakingBlock = _naturalSpeakingMode
            ? "\n\nNATURAL SPEAKING MODE (enabled for this session — follow strictly):" +
              "\nThe candidate will say your answer out loud in a live interview. It must sound like a real person thinking and talking, not like reading an AI script or essay." +
              "\n- Under ⭐️ **Answer**, use **spoken, conversational prose**: short clauses, natural rhythm, mild connectors where appropriate (e.g. “so”, “I think”, “for me”, “honestly”, “what I did was”). Slight imperfection is OK; avoid slick corporate polish." +
              "\n- **Do not** use chatbot openers or tone: avoid “Certainly”, “Great question”, “I'd be happy to”, “Let me walk you through”, “To summarize / In conclusion”, or blog/wiki cadence." +
              "\n- **Do not** sound meta (“In this answer I will…”). Start from substance quickly." +
              "\n- Keep the required section headings (💬 **Question**, ⭐️ **Answer**, icons, etc.) for the UI, but make the **words under Answer** feel improvised and human. For non-coding questions, prefer a few flowing sentences over dense bullet stacks unless bullets truly help." +
              "\n- For code/SQL: still use fenced blocks with language tags; introduce them with one casual sentence, not a formal preamble.\n"
            : string.Empty;
        return basePrompt +
               "\n\nYou are role-playing as the job candidate described in the resume. " +
               "Always answer in FIRST PERSON as that candidate. " +
               "Never mention that you are an AI, assistant, or language model. " +
               "Use the resume details (name, experience, skills, education) to answer questions such as 'What is your name?' or 'Introduce yourself' as the candidate." +
               extraBlock +
               naturalSpeakingBlock +
               $"\n\nLANGUAGE: Respond in {outputLanguage}. If you include code, keep code keywords/identifiers in their original language (usually English), but explain in {outputLanguage}." +
               "\n\nRESPONSE FORMAT AND STYLE (VERY IMPORTANT — follow every time):" +
               "\n" +
               "\n1) USE A FRIENDLY, STYLISH LAYOUT WITH ICONS — Prefer this structure (use these exact section labels when relevant):" +
               "\n💬 **Question**: <repeat the question in one line>" +
               "\n" +
               "\n⭐️ **Answer**:" +
               "\n<2–6 short paragraphs or bullets>" +
               "\n" +
               "\n✅ **Key approaches:** (optional, when multiple options exist)" +
               "\n- <bullet>" +
               "\n- <bullet>" +
               "\n" +
               "\n🧠 **Explanation:** (optional, when it helps)" +
               "\n- <1–3 bullets or a short paragraph>" +
               "\n" +
               "\n🧩 **Example**: (optional, include when code/SQL is needed)" +
               "\n```<language>" +
               "\n<code>" +
               "\n```" +
               "\n" +
               "\nICONS: Use real Unicode emoji only (the app renders them in full color). " +
               "Start section lines with: 💬 ⭐️ ✅ 🧠 🧩 🔍 📌 ⚠️ 💡 👍 🎯 (copy these exact characters). " +
               "Do NOT use symbol-font characters like ★ ☆ ✓ ● ○ ► or ASCII (>, *, [x]) — those stay white in the UI. " +
               "Use 2–6 emoji per answer; do not spam." +
               "\n" +
               "\n2) LEAD WITH THE ANSWER — The first lines under ⭐️ Answer should directly address the question. Do not bury the main message after long setup." +
               "\n" +
               "\n3) INTERVIEW-FRIENDLY TONE — Confident, concise, conversational first person. Avoid stiff meta-phrases ('In summary', 'This response will')." +
               "\n" +
               "\n4) DEPTH WHEN NEEDED — For behavioral questions, add short sub-bullets under ⭐️ Answer (Situation → Action → Result is fine). For simple factual questions, keep it short." +
               "\n" +
               "\n5) CODE / QUERIES — When code or SQL is needed, always use fenced code blocks with a language tag (```sql, ```csharp, ```javascript, etc.). Add brief inline comments only on non-obvious lines." +
               "\n" +
               "\n6) DO NOT ANSWER MULTIPLE QUESTIONS AT ONCE — If the input contains multiple prompts, answer ONLY the most recent / primary one unless explicitly asked to do multiple." +
               "\n" +
               "\n6.5) HIGHLIGHT IMPORTANT TERMS (two colors in the app) — " +
               "(1) Primary: wrap the most critical phrases in **double asterisks** (3–8 per answer). " +
               "(2) Secondary: wrap additional important words in ==double equals== (2–6 per answer) for a different highlight color. " +
               "Do not wrap whole sentences; keep highlights to short phrases." +
               "\n" +
               "\n7) ALTERNATE QUESTIONS (required every time, after the answer only) — When you finish the answer, append EXACTLY one block on new lines. " +
               "Do not add any text after this block. The block must contain three complete interview questions, each on its own line, that are closely related to the topic the user actually asked about " +
               "(rephrasings, narrower follow-ups, or closely adjacent angles on the same subject). They must sound like something the interviewer could have said if speech recognition misheard the original. " +
               "Do NOT use generic coaching labels (no \"More depth\", \"Simpler recap\", \"Trade-offs\", etc.). Each line must be a full question ending with ? when appropriate." +
               "\n" +
               "\n<<<ALT_Q>>>" +
               "\n1. [first related question]" +
               "\n2. [second related question]" +
               "\n3. [third related question]" +
               "\n<<<END_ALT_Q>>>";
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

        await ExecuteAnswerStreamCoreAsync(
            request,
            displayQuestionForUi,
            connectingMessage: "Connecting…",
            generatingMessage: "Generating answer…",
            aiChannel: "transcript").ConfigureAwait(false);
    }

    private async Task ExecuteAnswerStreamCoreAsync(
        HttpRequestMessage request,
        string? displayQuestionForUi,
        string connectingMessage = "Connecting…",
        string generatingMessage = "Generating answer…",
        string? aiChannel = null)
    {
        var answerEpochSnap = Volatile.Read(ref _answerUiEpoch);
        var answerStreamSessionId = _callSessionId;
        var streamDebug = App.Settings.StreamDebugLogs;

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
                LogInterviewUiStatus(connectingMessage);
                // Ensure the answer panel is visible for ALL answer sources so streaming text is visible immediately.
                // (Previously only screenshot flows forced the panel open; typed-input could stream into a hidden area.)
                if (FindName("AnswerSectionPanel") is UIElement pan && pan.Visibility == Visibility.Visible)
                {
                    if (FindName("AnswerSectionSeparator") is UIElement sep)
                        sep.Visibility = Visibility.Visible;
                    SetAnswerSectionRowHeight(collapsed: false);
                }
                else
                {
                    ShowAnswerSectionUi();
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
                if (!AnswerStreamStillValidForSession())
                    return;

                _activeStreamAppendsInPlace = _continueAnswerFromFollowUpChip;
                _continueAnswerFromFollowUpChip = false;
                BeginStreamingAnswerDisplay(_currentAnswerDisplayHeading, _activeStreamAppendsInPlace);
            });
            await Dispatcher.InvokeAsync(() =>
            {
                if (AnswerStreamStillValidForSession())
                    LogInterviewUiStatus(generatingMessage);
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
            var displayFilter = new AltQuestionStreamFilter();
            await using var respStream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
            using var reader = new StreamReader(respStream);

            var sw = streamDebug ? Stopwatch.StartNew() : null;
            long receivedLines = 0;
            long receivedDeltaChars = 0;
            long appendedDisplayChars = 0;
            var lastDebugMs = 0L;

            try
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    if (!AnswerStreamStillValidForSession())
                        return;

                    _answerStreamRenderEpochSnap = answerEpochSnap;
                    _answerStreamRenderCallSessionId = answerStreamSessionId;
                    lock (_answerStreamLiveTextLock)
                    {
                        _answerStreamLivePendingTokens.Clear();
                    }

                    _answerStreamLiveRendering = true;
                    _answerStreamFlushTimer?.Stop();
                    _answerStreamFlushTimer?.Start(); // steady UI pump while streaming
                    _answerStreamCompletionPending = false;
                    _answerStreamCompletionContent = string.Empty;
                    _answerStreamCompletionHeading = null;
                    _answerStreamCompletionAppendContinuation = false;
                    _answerStreamDebugUiFlushCount = 0;
                    Interlocked.Exchange(ref _answerStreamLastDebugUiLogTicks, 0);
                });

                if (streamDebug)
                    DesktopLogger.Info($"[STREAM:CLIENT] start channel={aiChannel ?? "<null>"} session={answerStreamSessionId} epoch={answerEpochSnap}");

                while (await reader.ReadLineAsync().ConfigureAwait(false) is { } line)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    receivedLines++;
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

                    receivedDeltaChars += delta.Length;
                    completionSb.Append(delta);

                    var displayDelta = displayFilter.Feed(delta);
                    if (!string.IsNullOrEmpty(displayDelta))
                    {
                        appendedDisplayChars += displayDelta.Length;
                        lock (_answerStreamLiveTextLock)
                        {
                            _answerStreamLivePendingTokens.Enqueue(displayDelta);
                        }
                        // Ensure timer is running (safe even if already running).
                        _ = Dispatcher.BeginInvoke(new Action(() => _answerStreamFlushTimer?.Start()), DispatcherPriority.Render);
                    }

                    if (streamDebug && sw != null)
                    {
                        var ms = (long)sw.Elapsed.TotalMilliseconds;
                        if (ms - lastDebugMs >= 350)
                        {
                            lastDebugMs = ms;
                            DesktopLogger.Info(
                                $"[STREAM:CLIENT] t={ms}ms lines={receivedLines} deltaChars={receivedDeltaChars} displayChars={appendedDisplayChars} " +
                                $"lastDeltaLen={delta.Length} markerSeen={displayFilter.MarkerSeen}");
                        }
                    }
                }
            }
            finally
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    if (AnswerStreamStillValidForSession())
                    {
                        // Flush any remaining buffered tail (unless ALT block started).
                        var remainder = displayFilter.FlushRemainder();
                        if (!string.IsNullOrEmpty(remainder))
                        {
                            lock (_answerStreamLiveTextLock)
                            {
                                _answerStreamLivePendingTokens.Enqueue(remainder);
                            }
                        }
                        // Keep the typewriter pump running; we'll finalize after the queue drains.
                        _answerStreamFlushTimer?.Start();
                    }
                });

                if (streamDebug && sw != null)
                {
                    sw.Stop();
                    DesktopLogger.Info(
                        $"[STREAM:CLIENT] end t={(long)sw.Elapsed.TotalMilliseconds}ms lines={receivedLines} deltaChars={receivedDeltaChars} displayChars={appendedDisplayChars} uiFlushes={_answerStreamDebugUiFlushCount}");
                }
            }

            var completion = completionSb.ToString();
            var completionClean = StripAltQuestionsForDisplay(completion).TrimEnd();
            var altQs = ExtractAltQuestions(completion);
            await Dispatcher.InvokeAsync(() =>
            {
                if (!AnswerStreamStillValidForSession())
                    return;

                // Defer final rich render until backlog drains so UI doesn't "jump" to full answer.
                _answerStreamCompletionContent = completionClean;
                _answerStreamCompletionHeading = _currentAnswerDisplayHeading;
                _answerStreamCompletionAppendContinuation = _activeStreamAppendsInPlace;
                _answerStreamCompletionPending = true;

                LogInterviewUiStatus("Generating answer…");
                RegisterSuccessfulAnswerInHistory(completionClean, mergeIntoLatest: _activeStreamAppendsInPlace);
                PopulateFollowUpSuggestionsFromAltQuestions(altQs, _currentAnswerDisplayHeading, completionClean, TranscriptTextBlock.Text);
                _showAskAiAnythingRow = completionClean.Length > 12
                                        && !completionClean.TrimStart().StartsWith("Error:", StringComparison.OrdinalIgnoreCase);
                UpdateAskAiAnythingRowVisibility();
                // Ensure the pump runs until drained + final render occurs in the tick handler.
                _answerStreamFlushTimer?.Start();
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
                ClearFollowUpSuggestions();
                ShowAnswerSectionUi();

                _showAskAiAnythingRow = false;
                UpdateAskAiAnythingRowVisibility();
                LogInterviewUiStatus("Error.");
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
        if (_answerGenerationInFlight)
        {
            LogInterviewUiStatus("Wait for the current answer to finish, then try Analyze screen again.");
            return;
        }

        if (!_sessionActive || !_mainUiStarted)
        {
            LogInterviewUiStatus("Start the interview session first, then use Analyze screen.");
            return;
        }

        ScreenshotAiButton.IsEnabled = false;
        var startedAnswerStream = false;
        try
        {
            // Same as "AI Answer": open the answer panel so streamed text is visible (otherwise it updates a hidden area).
            await Dispatcher.InvokeAsync(ShowAnswerSectionUi);

            _ = IncrementAiUsageAsync();
            LogInterviewUiStatus("Analyzing screen…");

            byte[] pngBytes;
            try
            {
                pngBytes = await Task.Run(() => ScreenCaptureHelper.CaptureVirtualScreenToPngBytes()).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    LogInterviewUiStatus($"Screen capture failed: {ex.Message}");
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

            startedAnswerStream = true;
            await ExecuteAnswerStreamCoreAsync(
                request,
                "Screenshot",
                connectingMessage: "Analyzing screen…",
                generatingMessage: "Generating answer…",
                aiChannel: "screenshot").ConfigureAwait(false);
        }
        finally
        {
            if (!startedAnswerStream)
                ScreenshotAiButton.IsEnabled = true;
        }
    }

    private void ClearAiAnswer()
    {
        _streamingAnswerRun = null;
        _streamingAnswerParagraph = null;
        _currentAnswerDisplayHeading = null;
        var emptyDoc = new FlowDocument(new Paragraph()) { FontFamily = AnswerEmojiFontFamily };
        AiAnswerTextBlock.Document = emptyDoc;
        ApplyAnswerEmojiGlyphs(emptyDoc);
        ClearFollowUpSuggestions();
    }

    private void ClearFollowUpSuggestions()
    {
        FollowUpChipsHost.Children.Clear();
        FollowUpSuggestionsPanel.Visibility = Visibility.Collapsed;
    }

    private void PopulateFollowUpSuggestionsFromAltQuestions(
        IReadOnlyList<string> modelQuestions,
        string? heading,
        string answerBodyClean,
        string? transcriptSnapshot = null)
    {
        FollowUpChipsHost.Children.Clear();
        if (modelQuestions == null
            || modelQuestions.Count < 3
            || string.IsNullOrWhiteSpace(answerBodyClean)
            || answerBodyClean.TrimStart().StartsWith("Error:", StringComparison.OrdinalIgnoreCase))
        {
            FollowUpSuggestionsPanel.Visibility = Visibility.Collapsed;
            return;
        }

        var q0 = modelQuestions[0].Trim();
        var q1 = modelQuestions[1].Trim();
        var q2 = modelQuestions[2].Trim();
        if (q0.Length < 8 || q1.Length < 8 || q2.Length < 8)
        {
            FollowUpSuggestionsPanel.Visibility = Visibility.Collapsed;
            return;
        }

        var focusTopic = BuildFollowUpFocusTopic(heading, answerBodyClean, transcriptSnapshot);
        var transcriptTail = TrimTranscriptTailForFollowUpPrompt(transcriptSnapshot);

        Style? chipStyle = null;
        try
        {
            chipStyle = (Style)FindResource("FollowUpChipButtonStyle");
        }
        catch
        {
            /* ignore missing style */
        }

        foreach (var q in new[] { q0, q1, q2 })
        {
            var chipLabel = q.Length > 80 ? q[..77].TrimEnd() + "…" : q;
            var tb = new TextBlock
            {
                Text = chipLabel,
                Foreground = FreezeBrush(0xcb, 0xd5, 0xe1),
                FontSize = 11.5,
                FontWeight = FontWeights.SemiBold,
                TextWrapping = TextWrapping.Wrap,
                MaxWidth = 360,
                VerticalAlignment = VerticalAlignment.Center
            };

            var b = new Button
            {
                Content = tb,
                Style = chipStyle,
                MinHeight = 36,
                Margin = new Thickness(0, 0, 10, 10),
                Padding = new Thickness(12, 8, 14, 8),
                HorizontalContentAlignment = HorizontalAlignment.Left,
                ToolTip = q
            };
            b.Click += async (_, _) =>
            {
                await SubmitQuestionAsync(q, continueConversation: true).ConfigureAwait(false);
            };
            FollowUpChipsHost.Children.Add(b);
        }

        FollowUpSuggestionsPanel.Visibility = Visibility.Visible;
    }

    private static string BuildFollowUpFocusTopic(string? heading, string answerBody, string? transcriptSnapshot)
    {
        var spoken = GetLatestQuestionFromTranscript(transcriptSnapshot ?? string.Empty).Trim();
        if (spoken.Length >= 12 && LooksLikeSubstantiveTranscriptSnippet(spoken))
            return ShortenSingleLine(spoken, 200);

        var h = (heading ?? string.Empty).Trim();
        if (h.Length >= 10 && LooksLikeSubstantiveTranscriptSnippet(h))
            return ShortenSingleLine(h, 200);

        return ShortenSingleLine(ExtractTopicSnippetFromAnswer(answerBody), 200);
    }

    private static bool LooksLikeSubstantiveTranscriptSnippet(string s)
    {
        var t = s.Trim();
        if (t.Length < 10)
            return false;
        if (Regex.IsMatch(t,
                @"^(ok|okay|thanks|thank you|yeah|yes|no|hmm|uh|um|right|sure|got it)\.?$",
                RegexOptions.IgnoreCase))
            return false;
        var words = t.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
        return words >= 3;
    }

    private static string ShortenSingleLine(string s, int maxChars)
    {
        var t = (s ?? string.Empty).Replace("\r\n", " ", StringComparison.Ordinal).Replace("\n", " ", StringComparison.Ordinal).Trim();
        while (t.Contains("  ", StringComparison.Ordinal))
            t = t.Replace("  ", " ", StringComparison.Ordinal);
        t = t.Replace("«", string.Empty, StringComparison.Ordinal).Replace("»", string.Empty, StringComparison.Ordinal).Trim();
        if (t.Length <= maxChars)
            return t.Length > 0 ? t : "this topic";
        return t[..(maxChars - 1)].TrimEnd() + "…";
    }

    private static string TrimTranscriptTailForFollowUpPrompt(string? transcript, int maxChars = 800)
    {
        var t = (transcript ?? string.Empty).Trim();
        if (t.Length == 0)
            return string.Empty;
        if (t.Length > maxChars)
            t = t[^maxChars..].TrimStart();
        return t;
    }

    private static string ExtractTopicSnippetFromAnswer(string answerBody)
    {
        var t = (answerBody ?? string.Empty).Replace("\r\n", "\n", StringComparison.Ordinal);
        foreach (var raw in t.Split('\n'))
        {
            var s = raw.Trim();
            if (s.Length < 12) continue;
            if (s.StartsWith("```", StringComparison.Ordinal)) continue;
            if (s.StartsWith("•", StringComparison.Ordinal) || s.StartsWith('-') || Regex.IsMatch(s, @"^\d+[\.\)]\s"))
                continue;
            return s.Length > 120 ? s[..120].TrimEnd() + "…" : s;
        }

        return "this topic";
    }

    private void BeginStreamingAnswerDisplay(string? boldHeading, bool appendContinuation)
    {
        ClearFollowUpSuggestions();
        _streamingAnswerRun = new Run(string.Empty)
        {
            Foreground = AnswerBodyBrush,
            FontFamily = AnswerEmojiFontFamily
        };

        FlowDocument doc;
        if (appendContinuation && AiAnswerTextBlock.Document != null && AiAnswerTextBlock.Document.Blocks.Count > 0)
        {
            doc = AiAnswerTextBlock.Document;
            doc.FontFamily = AnswerEmojiFontFamily;
            doc.Blocks.Add(new BlockUIContainer(new Border
            {
                Height = 1,
                Background = FreezeBrush(0x3d, 0x52, 0x6b),
                Margin = new Thickness(0, 10, 0, 12),
                SnapsToDevicePixels = true
            }));
        }
        else
        {
            doc = new FlowDocument
            {
                PagePadding = new Thickness(0),
                LineHeight = 22,
                FontFamily = AnswerEmojiFontFamily
            };
        }

        var streamParagraph = new Paragraph
        {
            Margin = new Thickness(0, 0, 0, 6),
            Foreground = AnswerBodyBrush,
            LineHeight = 22,
            FontFamily = AnswerEmojiFontFamily
        };
        _streamingAnswerParagraph = streamParagraph;
        streamParagraph.Inlines.Add(_streamingAnswerRun);
        doc.Blocks.Add(streamParagraph);

        AiAnswerTextBlock.Document = doc;
        ApplyAnswerEmojiGlyphs(doc);
        Dispatcher.BeginInvoke(new Action(() =>
        {
            ScrollViewer.SetCanContentScroll(AiAnswerTextBlock, false);
            ScrollViewer.SetPanningMode(AiAnswerTextBlock, PanningMode.VerticalOnly);
            UpdateAiAnswerBodyMaxHeight();
            ScheduleAnswerLayoutRefresh();
        }), DispatcherPriority.Loaded);
    }

    private void RenderAiAnswer(string content, string? boldHeading = null, bool appendContinuation = false)
    {
        _streamingAnswerRun = null;
        content = StripAltQuestionsForDisplay(content ?? string.Empty).TrimEnd();

        FlowDocument doc;
        if (appendContinuation && AiAnswerTextBlock.Document != null && AiAnswerTextBlock.Document.Blocks.Count > 0)
        {
            doc = AiAnswerTextBlock.Document;
            if (_streamingAnswerParagraph != null)
            {
                try
                {
                    if (doc.Blocks.Contains(_streamingAnswerParagraph))
                        doc.Blocks.Remove(_streamingAnswerParagraph);
                }
                catch
                {
                    /* ignore */
                }
            }

            _streamingAnswerParagraph = null;
            AppendAnswerMarkdownBodyToDocument(doc, content ?? string.Empty);
        }
        else
        {
            _streamingAnswerParagraph = null;
            doc = new FlowDocument
            {
                PagePadding = new Thickness(0),
                LineHeight = 22,
                FontFamily = AnswerEmojiFontFamily
            };
            AppendAnswerMarkdownBodyToDocument(doc, content ?? string.Empty);
        }

        doc.FontFamily = AnswerEmojiFontFamily;
        AiAnswerTextBlock.Document = doc;
        ApplyAnswerEmojiGlyphs(doc);
        Dispatcher.BeginInvoke(new Action(() =>
        {
            ScrollViewer.SetCanContentScroll(AiAnswerTextBlock, false);
            ScrollViewer.SetPanningMode(AiAnswerTextBlock, PanningMode.VerticalOnly);
            UpdateAiAnswerBodyMaxHeight();
            CoerceMainWindowHeightToContent();
            ScheduleAnswerLayoutRefresh();
        }), DispatcherPriority.Loaded);
    }

    private static void AppendAnswerMarkdownBodyToDocument(FlowDocument doc, string content)
    {
        doc.FontFamily = AnswerEmojiFontFamily;
        var normalized = (content ?? string.Empty).Replace("\r\n", "\n");
        var lines = normalized.Split('\n');
        var codeBuffer = new StringBuilder();
        var inCodeBlock = false;
        string? pendingFenceLang = null;

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
                    var afterFence = trimmed.Length > 3 ? trimmed[3..].Trim() : string.Empty;
                    pendingFenceLang = string.IsNullOrEmpty(afterFence) ? null : afterFence;
                }
                else
                {
                    AddCodeBlock(doc, codeBuffer.ToString(), pendingFenceLang);
                    pendingFenceLang = null;
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
                doc.Blocks.Add(new Paragraph(new Run(" "))
                {
                    Margin = new Thickness(0, 4, 0, 4),
                    LineHeight = 8
                });
                continue;
            }

            string text = line;
            if (Regex.IsMatch(trimmed, @"^\d+[\.\)]\s"))
                text = trimmed;
            else if (trimmed.StartsWith("- "))
                text = "• " + trimmed[2..];

            if (IsModelQuestionLine(text))
                AppendFormattedQuestionLineParagraph(doc, text);
            else
                AppendFormattedAnswerParagraph(doc, text);
        }

        if (inCodeBlock && codeBuffer.Length > 0)
            AddCodeBlock(doc, codeBuffer.ToString(), pendingFenceLang);
    }

    private static void AddCodeBlock(FlowDocument doc, string codeText, string? fenceLanguage = null)
    {
        var trimmedCode = (codeText ?? string.Empty).TrimEnd('\r', '\n');
        if (trimmedCode.Length == 0)
            return;

        var lang = NormalizeFenceLanguage(fenceLanguage);
        var codeCard = BuildCodeCardUi(trimmedCode, lang);
        doc.Blocks.Add(new BlockUIContainer(codeCard) { Margin = new Thickness(0, 8, 0, 10) });
    }

    private static void AppendHighlightedCodeLine(Paragraph paragraph, string line)
    {
        var commentIndex = line.IndexOf("//", StringComparison.Ordinal);
        if (commentIndex < 0)
            commentIndex = line.IndexOf("--", StringComparison.Ordinal);
        if (commentIndex < 0 && line.TrimStart().StartsWith("#", StringComparison.Ordinal))
            commentIndex = 0;

        var codePart = commentIndex >= 0 ? line[..commentIndex] : line;
        var commentPart = commentIndex >= 0 ? line[commentIndex..] : string.Empty;

        foreach (var token in Regex.Split(codePart, @"(\W+)"))
        {
            if (string.IsNullOrEmpty(token))
                continue;

            var run = new Run(token) { FontWeight = FontWeights.Normal };
            if (CodeKeywords.Contains(token))
                run.Foreground = CodeFgKeywordBrush;
            else if (Regex.IsMatch(token, "^\".*\"$|^'.*'$"))
                run.Foreground = CodeFgStringBrush;
            else if (Regex.IsMatch(token, @"^\d+$"))
                run.Foreground = CodeFgNumberBrush;
            else
                run.Foreground = CodeFgDefaultBrush;

            paragraph.Inlines.Add(run);
        }

        if (!string.IsNullOrWhiteSpace(commentPart))
        {
            paragraph.Inlines.Add(new Run(commentPart)
            {
                Foreground = CodeFgCommentBrush,
                FontWeight = FontWeights.Normal
            });
        }
    }

    private static string NormalizeFenceLanguage(string? fenceLanguage)
    {
        var raw = (fenceLanguage ?? string.Empty).Trim();
        if (string.IsNullOrEmpty(raw))
            return "CODE";

        // Handle a few common aliases and keep the header short (as shown in the reference).
        var l = raw.ToLowerInvariant();
        return l switch
        {
            "c#" or "csharp" or "cs" => "C#",
            "ts" or "typescript" => "TS",
            "js" or "javascript" => "JS",
            "py" or "python" => "PY",
            "json" => "JSON",
            "sql" or "tsql" or "t-sql" => "SQL",
            "bash" or "sh" or "shell" => "SH",
            "html" => "HTML",
            "css" => "CSS",
            _ => raw.Length <= 10 ? raw.ToUpperInvariant() : raw[..10].ToUpperInvariant()
        };
    }

    private static Border BuildCodeCardUi(string code, string languageLabel)
    {
        var outer = new Border
        {
            Background = CodeCardOuterBgBrush,
            BorderBrush = CodeCardOuterBorderBrush,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(14),
            SnapsToDevicePixels = true,
            UseLayoutRounding = true
        };

        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        // Header (language left, Copy button right).
        var header = new Border
        {
            Background = CodeCardHeaderBgBrush,
            BorderBrush = CodeCardHeaderBorderBrush,
            BorderThickness = new Thickness(0, 0, 0, 1),
            CornerRadius = new CornerRadius(14, 14, 0, 0),
            Padding = new Thickness(12, 7, 12, 7),
            SnapsToDevicePixels = true
        };

        var headerDock = new DockPanel { LastChildFill = false };

        var langText = new TextBlock
        {
            Text = languageLabel,
            Foreground = CodeCardLangBrush,
            FontFamily = new FontFamily("Segoe UI"),
            FontSize = 11.5,
            FontWeight = FontWeights.Bold,
            VerticalAlignment = VerticalAlignment.Center
        };
        DockPanel.SetDock(langText, Dock.Left);

        var copyBtn = new Button
        {
            Cursor = Cursors.Hand,
            FocusVisualStyle = null,
            Padding = new Thickness(0),
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Center,
            Height = 24,
            MinHeight = 24,
            Width = 44,
            MinWidth = 44,
            HorizontalContentAlignment = HorizontalAlignment.Center,
            VerticalContentAlignment = VerticalAlignment.Center
        };

        copyBtn.Template = BuildCodeCopyButtonTemplate();
        copyBtn.Click += (_, __) => CopyToClipboardWithTransientFeedback(copyBtn, code);

        copyBtn.Content = BuildCopyIcon();

        DockPanel.SetDock(copyBtn, Dock.Right);
        headerDock.Children.Add(copyBtn);
        headerDock.Children.Add(langText);
        header.Child = headerDock;

        // Body (monospace, horizontally scrollable, highlighted tokens).
        var bodyHost = new Border
        {
            Background = CodeCardOuterBgBrush,
            CornerRadius = new CornerRadius(0, 0, 14, 14),
            Padding = new Thickness(12, 10, 12, 12),
            SnapsToDevicePixels = true
        };

        var scroll = new ScrollViewer
        {
            HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
            CanContentScroll = false
        };

        var linesPanel = new StackPanel { Orientation = Orientation.Vertical };
        foreach (var rawLine in code.Replace("\r\n", "\n").Split('\n'))
        {
            var line = (rawLine ?? string.Empty).Replace("\t", "    ");
            var tb = new TextBlock
            {
                FontFamily = new FontFamily("Cascadia Mono, Consolas"),
                FontSize = 12.5,
                Foreground = CodeFgDefaultBrush,
                TextWrapping = TextWrapping.NoWrap,
                LineHeight = 18
            };
            TextOptions.SetTextFormattingMode(tb, TextFormattingMode.Display);
            TextOptions.SetTextRenderingMode(tb, TextRenderingMode.ClearType);
            AppendHighlightedCodeLine(tb, line);
            linesPanel.Children.Add(tb);
        }
        scroll.Content = linesPanel;
        bodyHost.Child = scroll;

        Grid.SetRow(header, 0);
        Grid.SetRow(bodyHost, 1);
        grid.Children.Add(header);
        grid.Children.Add(bodyHost);

        outer.Child = grid;
        return outer;
    }

    private static ControlTemplate BuildCodeCopyButtonTemplate()
    {
        var bd = new FrameworkElementFactory(typeof(Border));
        bd.Name = "bd";
        bd.SetValue(Border.CornerRadiusProperty, new CornerRadius(10));
        bd.SetValue(Border.PaddingProperty, new Thickness(8, 3, 8, 3));
        bd.SetValue(Border.BackgroundProperty, Brushes.Transparent);
        bd.SetValue(Border.BorderBrushProperty, CodeCardCopyBorderBrush);
        bd.SetValue(Border.BorderThicknessProperty, new Thickness(1));
        bd.SetValue(Border.SnapsToDevicePixelsProperty, true);

        var presenter = new FrameworkElementFactory(typeof(ContentPresenter));
        presenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
        presenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
        bd.AppendChild(presenter);

        var template = new ControlTemplate(typeof(Button)) { VisualTree = bd };

        var over = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
        over.Setters.Add(new Setter(Border.BackgroundProperty, CodeCardCopyHoverBgBrush, "bd"));
        over.Setters.Add(new Setter(Border.BorderBrushProperty, CodeCardCopyHoverBorderBrush, "bd"));

        var pressed = new Trigger { Property = Button.IsPressedProperty, Value = true };
        pressed.Setters.Add(new Setter(Border.BackgroundProperty, CodeCardCopyPressedBgBrush, "bd"));
        pressed.Setters.Add(new Setter(Border.BorderBrushProperty, CodeCardCopyHoverBorderBrush, "bd"));

        var disabled = new Trigger { Property = UIElement.IsEnabledProperty, Value = false };
        disabled.Setters.Add(new Setter(UIElement.OpacityProperty, 0.55, "bd"));

        template.Triggers.Add(over);
        template.Triggers.Add(pressed);
        template.Triggers.Add(disabled);
        return template;
    }

    private static void CopyToClipboardWithTransientFeedback(Button copyButton, string code)
    {
        try
        {
            Clipboard.SetText(code);
        }
        catch
        {
            // Clipboard can fail (busy / denied). Keep UI silent; copy button remains usable.
            return;
        }

        try
        {
            // Show "Copied" tooltip anchored to the button.
            var tip = copyButton.ToolTip as ToolTip;
            if (tip == null)
            {
                tip = new ToolTip
                {
                    Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom,
                    VerticalOffset = 6,
                    StaysOpen = true,
                    Content = new TextBlock
                    {
                        Text = "Copied",
                        Foreground = Brushes.White,
                        FontFamily = new FontFamily("Segoe UI"),
                        FontSize = 12,
                        FontWeight = FontWeights.SemiBold
                    }
                };
                copyButton.ToolTip = tip;
            }

            tip.IsOpen = true;
            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(900) };
            timer.Tick += (_, __) =>
            {
                timer.Stop();
                tip.IsOpen = false;
            };
            timer.Start();
        }
        catch
        {
            /* ignore */
        }
    }

    private static FrameworkElement BuildCopyIcon()
    {
        // Font Awesome icon (vector, crisp at any scale).
        // "Copy" exists in FontAwesome 4.x; if you want the newer FA6 set, we can switch packages.
        return new FontAwesome.WPF.FontAwesome
        {
            Icon = FontAwesome.WPF.FontAwesomeIcon.Copy,
            Width = 18,
            Height = 18,
            Foreground = CodeCardCopyFgBrush,
            SnapsToDevicePixels = true,
            UseLayoutRounding = true
        };
    }

    private static void AppendHighlightedCodeLine(TextBlock textBlock, string line)
    {
        var commentIndex = line.IndexOf("//", StringComparison.Ordinal);
        if (commentIndex < 0)
            commentIndex = line.IndexOf("--", StringComparison.Ordinal);
        if (commentIndex < 0 && line.TrimStart().StartsWith("#", StringComparison.Ordinal))
            commentIndex = 0;

        var codePart = commentIndex >= 0 ? line[..commentIndex] : line;
        var commentPart = commentIndex >= 0 ? line[commentIndex..] : string.Empty;

        foreach (var token in Regex.Split(codePart, @"(\W+)"))
        {
            if (string.IsNullOrEmpty(token))
                continue;

            var run = new Run(token) { FontWeight = FontWeights.Normal };
            if (CodeKeywords.Contains(token))
                run.Foreground = CodeFgKeywordBrush;
            else if (Regex.IsMatch(token, "^\".*\"$|^'.*'$"))
                run.Foreground = CodeFgStringBrush;
            else if (Regex.IsMatch(token, @"^\d+$"))
                run.Foreground = CodeFgNumberBrush;
            else
                run.Foreground = CodeFgDefaultBrush;

            textBlock.Inlines.Add(run);
        }

        if (!string.IsNullOrWhiteSpace(commentPart))
        {
            textBlock.Inlines.Add(new Run(commentPart)
            {
                Foreground = CodeFgCommentBrush,
                FontWeight = FontWeights.Normal
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
                LogInterviewUiStatus($"Log failed: {msg}");
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
                    SetSessionExtraContext(null);
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