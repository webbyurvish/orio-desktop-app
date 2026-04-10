using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace AiInterviewAssistant;

public sealed class DesktopUserState : INotifyPropertyChanged
{
    public static DesktopUserState Current { get; } = new();

    private string _email = string.Empty;

    public event PropertyChangedEventHandler? PropertyChanged;

    public string Email
    {
        get => _email;
        private set
        {
            if (string.Equals(_email, value, StringComparison.Ordinal)) return;
            _email = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(HasEmail));
        }
    }

    public bool HasEmail => !string.IsNullOrWhiteSpace(_email);

    public void SetEmail(string? email) => Email = (email ?? string.Empty).Trim();

    public void Clear() => Email = string.Empty;

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

