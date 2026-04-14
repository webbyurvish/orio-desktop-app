using System;
using System.IO;

namespace AiInterviewAssistant;

public static class DesktopLogger
{
    private static readonly object _lock = new();
    public static string LogFilePath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SmeedAI",
        "AiInterviewAssistant",
        "desktop.log"
    );

    public static string FallbackLogFilePath { get; } = Path.Combine(
        AppContext.BaseDirectory,
        "desktop.log"
    );

    public static void Info(string message) => Write("INFO", message);
    public static void Warn(string message) => Write("WARN", message);
    public static void Error(string message) => Write("ERROR", message);

    private static void Write(string level, string message)
    {
        var line = $"{DateTime.UtcNow:O} [{level}] {message}{Environment.NewLine}";

        // Try primary location first
        try
        {
            var dir = Path.GetDirectoryName(LogFilePath);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);

            lock (_lock)
            {
                File.AppendAllText(LogFilePath, line);
            }
            return;
        }
        catch
        {
            // ignore and try fallback
        }

        // Fallback: write next to the exe
        try
        {
            lock (_lock)
            {
                File.AppendAllText(FallbackLogFilePath, line);
            }
        }
        catch
        {
            // Never throw from logger
        }
    }
}

