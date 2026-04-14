using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace AiInterviewAssistant;

public sealed class CreditsState : INotifyPropertyChanged
{
    public static CreditsState Current { get; } = new();

    /// <summary>Minimum balance required to start a paid (credit) interview block (matches 0.5 on first activate).</summary>
    public const decimal MinimumCreditsForPaidSessionActivation = 0.5m;

    private decimal? _credits;
    private bool _isKnown;

    public event PropertyChangedEventHandler? PropertyChanged;

    public bool IsKnown => _isKnown;

    public string CreditsText
    {
        get
        {
            if (!_isKnown) return "Loading...";
            if (!_credits.HasValue) return "No Credits";
            var v = _credits.Value;
            if (v <= 0m) return "No Credits";
            // Prefer concise label; e.g. "2.5 Credits"
            return $"{v:0.##} Credits";
        }
    }

    public void SetLoading()
    {
        _isKnown = false;
        _credits = null;
        OnPropertyChanged(nameof(CreditsText));
    }

    public void SetUnknown()
    {
        _isKnown = true;
        _credits = null;
        OnPropertyChanged(nameof(CreditsText));
    }

    public void SetCredits(decimal credits)
    {
        _isKnown = true;
        _credits = credits;
        OnPropertyChanged(nameof(CreditsText));
    }

    /// <summary>
    /// When credit balance is still loading (<see cref="IsKnown"/> false), returns true so the user is not blocked prematurely.
    /// When known, requires at least <see cref="MinimumCreditsForPaidSessionActivation"/>.
    /// </summary>
    public bool HasSufficientCreditsForPaidActivation()
    {
        if (!_isKnown) return true;
        if (!_credits.HasValue) return false;
        return _credits.Value >= MinimumCreditsForPaidSessionActivation;
    }

    /// <summary>Session setup: show Buy credits tile instead of Full session when balance is known and below minimum.</summary>
    public bool ShouldShowBuyCreditsInsteadOfFullSession()
    {
        if (!_isKnown) return false;
        if (!_credits.HasValue) return true;
        return _credits.Value < MinimumCreditsForPaidSessionActivation;
    }

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

