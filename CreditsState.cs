using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace AiInterviewAssistant;

public sealed class CreditsState : INotifyPropertyChanged
{
    public static CreditsState Current { get; } = new();

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

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

