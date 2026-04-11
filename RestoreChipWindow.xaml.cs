using System;
using System.Windows;
using System.Windows.Input;

namespace AiInterviewAssistant;

/// <summary>
/// Small floating circle shown when the main window is “minimized” from the login header.
/// Click to restore the main window.
/// </summary>
public partial class RestoreChipWindow : Window
{
    public RestoreChipWindow()
    {
        InitializeComponent();
    }

    public event EventHandler? RestoreRequested;

    private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        RestoreRequested?.Invoke(this, EventArgs.Empty);
    }

}
