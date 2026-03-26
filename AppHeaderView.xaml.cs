using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace AiInterviewAssistant;

public partial class AppHeaderView : UserControl
{
    public AppHeaderView()
    {
        InitializeComponent();
    }

    public static readonly DependencyProperty CreditsTextProperty =
        DependencyProperty.Register(nameof(CreditsText), typeof(string), typeof(AppHeaderView),
            new PropertyMetadata("No Credits"));

    public static readonly DependencyProperty HeaderBackgroundProperty =
        DependencyProperty.Register(nameof(HeaderBackground), typeof(Brush), typeof(AppHeaderView),
            new PropertyMetadata(new SolidColorBrush(Color.FromRgb(0xE4, 0xE8, 0xF0))));

    public static readonly DependencyProperty HeaderBorderBrushProperty =
        DependencyProperty.Register(nameof(HeaderBorderBrush), typeof(Brush), typeof(AppHeaderView),
            new PropertyMetadata(new SolidColorBrush(Color.FromRgb(0xD5, 0xDA, 0xE3))));

    public string CreditsText
    {
        get => (string)GetValue(CreditsTextProperty);
        set => SetValue(CreditsTextProperty, value);
    }

    public Brush HeaderBackground
    {
        get => (Brush)GetValue(HeaderBackgroundProperty);
        set => SetValue(HeaderBackgroundProperty, value);
    }

    public Brush HeaderBorderBrush
    {
        get => (Brush)GetValue(HeaderBorderBrushProperty);
        set => SetValue(HeaderBorderBrushProperty, value);
    }

    public FrameworkElement MoreMenuAnchorElement => MoreMenuAnchor;
    public FrameworkElement MoveMenuAnchorElement => MoveMenuAnchor;

    public event RoutedEventHandler? MoreClicked;
    public event RoutedEventHandler? MoveClicked;
    public event RoutedEventHandler? MinimizeClicked;
    public event RoutedEventHandler? CloseClicked;
    public event MouseButtonEventHandler? HeaderDragRequested;

    private void MoreMenuAnchor_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        MoreClicked?.Invoke(this, new RoutedEventArgs());
    }

    private void MoveMenuAnchor_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        MoveClicked?.Invoke(this, new RoutedEventArgs());
    }

    private void MinimizeBorder_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        MinimizeClicked?.Invoke(this, new RoutedEventArgs());
    }

    private void CloseBorder_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        CloseClicked?.Invoke(this, new RoutedEventArgs());
    }

    private void HeaderBorder_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ButtonState != MouseButtonState.Pressed) return;
        HeaderDragRequested?.Invoke(this, e);
    }
}
