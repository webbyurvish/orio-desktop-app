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
            new PropertyMetadata(CreateAuroraHeaderBrush()));

    public static readonly DependencyProperty HeaderBorderBrushProperty =
        DependencyProperty.Register(nameof(HeaderBorderBrush), typeof(Brush), typeof(AppHeaderView),
            new PropertyMetadata(CreateAuroraHeaderDividerBrush()));

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

    /// <summary>Web .orio-app: --orio-surface → --orio-elevated</summary>
    private static LinearGradientBrush CreateAuroraHeaderBrush()
    {
        var b = new LinearGradientBrush
        {
            StartPoint = new System.Windows.Point(0, 0),
            EndPoint = new System.Windows.Point(1, 0.35),
            MappingMode = BrushMappingMode.RelativeToBoundingBox,
        };
        b.GradientStops.Add(new GradientStop(Color.FromRgb(0x0C, 0x0C, 0x12), 0));
        b.GradientStops.Add(new GradientStop(Color.FromRgb(0x12, 0x12, 0x1C), 1));
        if (b.CanFreeze) b.Freeze();
        return b;
    }

    /// <summary>Web --orio-border–like hairline (~8% white).</summary>
    private static SolidColorBrush CreateAuroraHeaderDividerBrush()
    {
        var br = new SolidColorBrush(Color.FromArgb(0x14, 0xFF, 0xFF, 0xFF));
        if (br.CanFreeze) br.Freeze();
        return br;
    }

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

    private void LogoImage_ImageFailed(object sender, ExceptionRoutedEventArgs e)
    {
        // If the logo isn't available at site-of-origin, keep the built-in gradient mark.
        LogoImage.Visibility = Visibility.Collapsed;
        LogoFallback.Visibility = Visibility.Visible;
    }
}
