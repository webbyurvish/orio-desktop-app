using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using NAudio.CoreAudioApi;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;

namespace AiInterviewAssistant;

public partial class ActivateSessionView : UserControl
{
    private Window? _hostWindow;
    private readonly DispatcherTimer _uiMeterTimer;

    private readonly MMDeviceEnumerator _deviceEnumerator = new();
    private List<DeviceItem> _micDevices = new();

    private WasapiCapture? _micCapture;
    private MMDevice? _micDevice;
    private volatile float _micPeak01;
    private DateTime _micLastActivityAtUtc = DateTime.MinValue;
    private EventHandler<WaveInEventArgs>? _micDataHandler;
    private DateTime _micTestStartedAtUtc = DateTime.MinValue;
    private bool _micEverWorked;
    private bool _micFailed;

    private WasapiLoopbackCapture? _systemLoopback;
    private MMDevice? _systemDevice;
    private volatile float _systemPeak01;
    private DateTime _systemLastActivityAtUtc = DateTime.MinValue;
    private IWavePlayer? _systemPlayer;
    private EventHandler<WaveInEventArgs>? _systemDataHandler;
    private DateTime _systemTestStartedAtUtc = DateTime.MinValue;
    private bool _systemEverWorked;
    private bool _systemFailed;

    private bool _audioTestRunning;

    private static readonly TimeSpan MicNoSignalFailAfter = TimeSpan.FromSeconds(6);
    private static readonly TimeSpan SystemNoSignalFailAfter = TimeSpan.FromSeconds(4);

    private static readonly Brush OkBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#16A34A"));
    private static readonly Brush BadBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DC2626"));
    private static readonly Brush NeutralBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#64748B"));

    public ActivateSessionView()
    {
        InitializeComponent();
        Loaded += OnLoaded;
        Unloaded += OnUnloaded;

        _uiMeterTimer = new DispatcherTimer(DispatcherPriority.Background)
        {
            Interval = TimeSpan.FromMilliseconds(60)
        };
        _uiMeterTimer.Tick += (_, __) => UpdateMetersUi();
    }

    public event RoutedEventHandler? BackRequested;
    public event RoutedEventHandler? ActivateRequested;
    public event RoutedEventHandler? CloseRequested;
    public event RoutedEventHandler? MinimizeRequested;
    public event Action<StartupWindowSlot>? WindowSlotRequested;
    public event RoutedEventHandler? DashboardRequested;
    public event RoutedEventHandler? LogoutRequested;

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        _hostWindow = Window.GetWindow(this);
        if (_hostWindow != null)
            _hostWindow.PreviewMouseLeftButtonDown += OnWindowPreviewMouseLeftButtonDown;

        MoreOptionsPopup.PlacementTarget = SharedHeader.MoreMenuAnchorElement;
        MoveOptionsPopup.PlacementTarget = SharedHeader.MoveMenuAnchorElement;

        LoadAudioDevices();
        _uiMeterTimer.Start();
    }

    private void OnUnloaded(object sender, RoutedEventArgs e)
    {
        _uiMeterTimer.Stop();
        StopAudioTest();

        if (_hostWindow != null)
        {
            _hostWindow.PreviewMouseLeftButtonDown -= OnWindowPreviewMouseLeftButtonDown;
            _hostWindow = null;
        }
    }

    private void OnWindowPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (!MoreOptionsPopup.IsOpen && !MoveOptionsPopup.IsOpen) return;
        if (e.OriginalSource is not DependencyObject src) return;
        if (IsWithin(src, SharedHeader.MoreMenuAnchorElement) || IsWithin(src, SharedHeader.MoveMenuAnchorElement)) return;
        if (MoreOptionsPopup.Child is DependencyObject morePopupRoot && IsWithin(src, morePopupRoot)) return;
        if (MoveOptionsPopup.Child is DependencyObject movePopupRoot && IsWithin(src, movePopupRoot)) return;
        MoreOptionsPopup.IsOpen = false;
        MoveOptionsPopup.IsOpen = false;
    }

    private static bool IsWithin(DependencyObject? child, DependencyObject? ancestor)
    {
        while (child != null)
        {
            if (ReferenceEquals(child, ancestor)) return true;
            child = VisualTreeHelper.GetParent(child);
        }
        return false;
    }

    private void Header_MoreClicked(object? sender, RoutedEventArgs e)
    {
        MoveOptionsPopup.IsOpen = false;
        MoreOptionsPopup.IsOpen = !MoreOptionsPopup.IsOpen;
    }

    private void Header_MoveClicked(object? sender, RoutedEventArgs e)
    {
        MoreOptionsPopup.IsOpen = false;
        MoveOptionsPopup.IsOpen = false;

        var owner = Window.GetWindow(this);
        if (owner == null) return;

        var picked = MoveOverlayWindow.PickSlot(owner);
        if (picked != null)
            WindowSlotRequested?.Invoke(picked.Value);
    }

    private void EmitWindowSlot(StartupWindowSlot slot, MouseButtonEventArgs e)
    {
        e.Handled = true;
        MoveOptionsPopup.IsOpen = false;
        WindowSlotRequested?.Invoke(slot);
    }

    private void MoveTopLeft_Click(object sender, MouseButtonEventArgs e) => EmitWindowSlot(StartupWindowSlot.TopLeft, e);
    private void MoveTopCenter_Click(object sender, MouseButtonEventArgs e) => EmitWindowSlot(StartupWindowSlot.TopCenter, e);
    private void MoveTopRight_Click(object sender, MouseButtonEventArgs e) => EmitWindowSlot(StartupWindowSlot.TopRight, e);
    private void MoveBottomLeft_Click(object sender, MouseButtonEventArgs e) => EmitWindowSlot(StartupWindowSlot.BottomLeft, e);
    private void MoveBottomCenter_Click(object sender, MouseButtonEventArgs e) => EmitWindowSlot(StartupWindowSlot.BottomCenter, e);
    private void MoveBottomRight_Click(object sender, MouseButtonEventArgs e) => EmitWindowSlot(StartupWindowSlot.BottomRight, e);

    private void MenuDashboard_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        MoreOptionsPopup.IsOpen = false;
        DashboardRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void MenuLogout_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        MoreOptionsPopup.IsOpen = false;
        LogoutRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void Back_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        BackRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void Activate_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        ActivateRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void Header_CloseClicked(object? sender, RoutedEventArgs e)
    {
        CloseRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void Header_MinimizeClicked(object? sender, RoutedEventArgs e)
    {
        MinimizeRequested?.Invoke(this, new RoutedEventArgs());
    }

    private void LoadAudioDevices()
    {
        try
        {
            _micDevices = _deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active)
                .Select(d => new DeviceItem(d))
                .OrderBy(d => d.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            MicDeviceCombo.ItemsSource = _micDevices;
            MicDeviceCombo.DisplayMemberPath = nameof(DeviceItem.Name);
            MicDeviceCombo.SelectedItem = _micDevices.FirstOrDefault(i => i.IsDefault) ?? _micDevices.FirstOrDefault();

            _systemDevice = TryGetDefaultOutputDevice();

            ResetAudioUiToNotTested();
        }
        catch (Exception ex)
        {
            MicStatusText.Text = $"Audio devices error: {ex.Message}";
            SystemStatusText.Text = $"Audio devices error: {ex.Message}";
            MicStatusText.Foreground = BadBrush;
            SystemStatusText.Foreground = BadBrush;
        }
    }

    private void AudioTest_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        if (_audioTestRunning) StopAudioTest();
        else StartAudioTest();
    }

    private void AudioReset_Click(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
        PrepareForDisplay();
    }

    /// <summary>
    /// Call when this screen is shown again (e.g. new session). The control often stays in the tree with only
    /// <see cref="UIElement.Visibility"/> toggled, so <see cref="OnLoaded"/> does not re-run and stale
    /// &quot;Working&quot; state must be cleared explicitly.
    /// </summary>
    public void PrepareForDisplay()
    {
        StopAudioTest();
        _micEverWorked = false;
        _systemEverWorked = false;
        _micFailed = false;
        _systemFailed = false;
        _micLastActivityAtUtc = DateTime.MinValue;
        _systemLastActivityAtUtc = DateTime.MinValue;
        _micTestStartedAtUtc = DateTime.MinValue;
        _systemTestStartedAtUtc = DateTime.MinValue;
        _micPeak01 = 0;
        _systemPeak01 = 0;

        try
        {
            LoadAudioDevices();
        }
        catch (Exception ex)
        {
            MicStatusText.Text = $"Audio devices error: {ex.Message}";
            SystemStatusText.Text = $"Audio devices error: {ex.Message}";
            MicStatusText.Foreground = BadBrush;
            SystemStatusText.Foreground = BadBrush;
        }
    }

    private void StartAudioTest()
    {
        StopAudioTest();

        _audioTestRunning = true;
        AudioTestBtnText.Text = "Stop Test";

        if (!_micEverWorked && !_micFailed)
        {
            MicStatusText.Text = _micDevices.Count == 0 ? "No microphone devices found" : "No signal";
            MicStatusText.Foreground = _micDevices.Count == 0 ? BadBrush : BadBrush;
        }

        if (!_systemEverWorked && !_systemFailed)
        {
            SystemStatusText.Text = _systemDevice == null ? "No output device found" : "No signal";
            SystemStatusText.Foreground = _systemDevice == null ? BadBrush : BadBrush;
        }

        if (!_micFailed && _micDevices.Count > 0)
            StartMicCapture();

        if (!_systemFailed && _systemDevice != null)
            StartSystemLoopbackAndTone();
    }

    private void StopAudioTest()
    {
        _audioTestRunning = false;
        AudioTestBtnText.Text = "Test Audio";

        StopMicCaptureOnly();
        StopSystemCaptureOnly();
    }

    private void StartMicCapture()
    {
        var selected = MicDeviceCombo.SelectedItem as DeviceItem;
        _micDevice = selected?.Device ?? _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications);
        try
        {
            _micPeak01 = 0;
            _micLastActivityAtUtc = DateTime.MinValue;
            _micTestStartedAtUtc = DateTime.UtcNow;

            _micCapture = new WasapiCapture(_micDevice) { ShareMode = AudioClientShareMode.Shared };
            _micDataHandler = (_, args) =>
            {
                if (_micCapture == null) return;
                var peak = ComputePeak01(_micCapture.WaveFormat, args.Buffer, args.BytesRecorded);
                if (peak > _micPeak01) _micPeak01 = peak;
                if (peak > 0.03f) _micLastActivityAtUtc = DateTime.UtcNow;
            };
            _micCapture.DataAvailable += _micDataHandler;
            _micCapture.RecordingStopped += (_, __) => { };
            _micCapture.StartRecording();
        }
        catch (Exception ex)
        {
            _micFailed = true;
            MicStatusText.Text = $"Failed";
            MicStatusText.Foreground = BadBrush;
            DesktopLogger.Warn($"Mic test start failed: {ex.Message}");
        }
    }

    private void StartSystemLoopbackAndTone()
    {
        if (_systemDevice == null) return;
        try
        {
            _systemPlayer?.Stop();
            _systemPlayer?.Dispose();
        }
        catch { /* ignore */ }
        finally
        {
            _systemPlayer = null;
        }

        try
        {
            _systemPeak01 = 0;
            _systemLastActivityAtUtc = DateTime.MinValue;
            _systemTestStartedAtUtc = DateTime.UtcNow;

            _systemLoopback = new WasapiLoopbackCapture(_systemDevice);
            _systemDataHandler = (_, args) =>
            {
                if (_systemLoopback == null) return;
                var peak = ComputePeak01(_systemLoopback.WaveFormat, args.Buffer, args.BytesRecorded);
                if (peak > _systemPeak01) _systemPeak01 = peak;
                if (peak > 0.02f) _systemLastActivityAtUtc = DateTime.UtcNow;
            };
            _systemLoopback.DataAvailable += _systemDataHandler;
            _systemLoopback.RecordingStopped += (_, __) => { };
            _systemLoopback.StartRecording();

            var tone = new SignalGenerator { Gain = 0.10, Frequency = 880, Type = SignalGeneratorType.Sin };
            var take = new OffsetSampleProvider(tone) { Take = TimeSpan.FromSeconds(1.2) };
            var wave = new SampleToWaveProvider16(take);
            _systemPlayer = new WasapiOut(_systemDevice, AudioClientShareMode.Shared, true, 50);
            _systemPlayer.Init(wave);
            _systemPlayer.Play();
        }
        catch (Exception ex)
        {
            _systemFailed = true;
            SystemStatusText.Text = "Failed";
            SystemStatusText.Foreground = BadBrush;
            DesktopLogger.Warn($"System audio test start failed: {ex.Message}");
        }
    }

    private void StopMicCaptureOnly()
    {
        try
        {
            if (_micCapture != null)
            {
                if (_micDataHandler != null)
                    _micCapture.DataAvailable -= _micDataHandler;
                _micCapture.StopRecording();
                _micCapture.Dispose();
            }
        }
        catch { /* ignore */ }
        finally
        {
            _micCapture = null;
            _micDataHandler = null;
            _micPeak01 = 0;
            _micTestStartedAtUtc = DateTime.MinValue;
            MicLevelBar.Value = 0;
        }
    }

    private void StopSystemCaptureOnly()
    {
        try
        {
            _systemPlayer?.Stop();
            _systemPlayer?.Dispose();
        }
        catch { /* ignore */ }
        finally
        {
            _systemPlayer = null;
        }

        try
        {
            if (_systemLoopback != null)
            {
                if (_systemDataHandler != null)
                    _systemLoopback.DataAvailable -= _systemDataHandler;
                _systemLoopback.StopRecording();
                _systemLoopback.Dispose();
            }
        }
        catch { /* ignore */ }
        finally
        {
            _systemLoopback = null;
            _systemDataHandler = null;
            _systemPeak01 = 0;
            _systemTestStartedAtUtc = DateTime.MinValue;
            SystemLevelBar.Value = 0;
        }
    }

    private void UpdateMetersUi()
    {
        // mic
        var micVal = Math.Clamp(_micPeak01, 0f, 1f) * 100f;
        MicLevelBar.Value = micVal;
        _micPeak01 = 0; // decay by resetting peak each tick

        if (_audioTestRunning && !_micEverWorked && !_micFailed && _micDevices.Count > 0)
        {
            var activeAgo = DateTime.UtcNow - _micLastActivityAtUtc;
            if (_micLastActivityAtUtc != DateTime.MinValue && activeAgo < TimeSpan.FromSeconds(1.2))
            {
                _micEverWorked = true;
                MicStatusText.Text = "Working";
                MicStatusText.Foreground = OkBrush;
            }
            else
            {
                var sinceStart = DateTime.UtcNow - _micTestStartedAtUtc;
                if (_micTestStartedAtUtc != DateTime.MinValue && sinceStart >= MicNoSignalFailAfter)
                {
                    _micFailed = true;
                    MicStatusText.Text = "Failed";
                    MicStatusText.Foreground = BadBrush;
                    StopMicCaptureOnly();
                }
                else
                {
                    MicStatusText.Text = "No signal";
                    MicStatusText.Foreground = BadBrush;
                }
            }
        }
        else if (_micEverWorked)
        {
            MicStatusText.Text = "Working";
            MicStatusText.Foreground = OkBrush;
        }

        // system
        var sysVal = Math.Clamp(_systemPeak01, 0f, 1f) * 100f;
        SystemLevelBar.Value = sysVal;
        _systemPeak01 = 0;

        if (_audioTestRunning && !_systemEverWorked && !_systemFailed && _systemDevice != null)
        {
            var activeAgo = DateTime.UtcNow - _systemLastActivityAtUtc;
            if (_systemLastActivityAtUtc != DateTime.MinValue && activeAgo < TimeSpan.FromSeconds(1.2))
            {
                _systemEverWorked = true;
                SystemStatusText.Text = "Working";
                SystemStatusText.Foreground = OkBrush;
            }
            else
            {
                var sinceStart = DateTime.UtcNow - _systemTestStartedAtUtc;
                if (_systemTestStartedAtUtc != DateTime.MinValue && sinceStart >= SystemNoSignalFailAfter)
                {
                    _systemFailed = true;
                    SystemStatusText.Text = "Failed";
                    SystemStatusText.Foreground = BadBrush;
                    StopSystemCaptureOnly();
                }
                else
                {
                    SystemStatusText.Text = "No signal";
                    SystemStatusText.Foreground = BadBrush;
                }
            }
        }
        else if (_systemEverWorked)
        {
            SystemStatusText.Text = "Working";
            SystemStatusText.Foreground = OkBrush;
        }

        // If both streams are done (either worked or failed), stop overall test.
        if (_audioTestRunning)
        {
            bool micDone = _micEverWorked || _micFailed || _micDevices.Count == 0;
            bool sysDone = _systemEverWorked || _systemFailed || _systemDevice == null;
            if (micDone && sysDone)
                StopAudioTest();
        }
    }

    private MMDevice? TryGetDefaultOutputDevice()
    {
        try
        {
            return _deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
        }
        catch
        {
            return null;
        }
    }

    private void ResetAudioUiToNotTested()
    {
        MicLevelBar.Value = 0;
        SystemLevelBar.Value = 0;

        if (_micDevices.Count == 0)
        {
            MicStatusText.Text = "No microphone devices found";
            MicStatusText.Foreground = BadBrush;
        }
        else
        {
            MicStatusText.Text = "Not tested";
            MicStatusText.Foreground = NeutralBrush;
        }

        if (_systemDevice == null)
        {
            SystemStatusText.Text = "No output device found";
            SystemStatusText.Foreground = BadBrush;
        }
        else
        {
            SystemStatusText.Text = "Not tested";
            SystemStatusText.Foreground = NeutralBrush;
        }

        AudioTestBtnText.Text = "Test Audio";
    }

    private static float ComputePeak01(WaveFormat format, byte[] buffer, int bytesRecorded)
    {
        if (bytesRecorded <= 0) return 0f;

        // Most WASAPI captures here are either 32-bit float or 16-bit PCM.
        if (format.Encoding == WaveFormatEncoding.IeeeFloat && format.BitsPerSample == 32)
        {
            int samples = bytesRecorded / 4;
            float peak = 0f;
            for (int n = 0; n < samples; n++)
            {
                float sample = BitConverter.ToSingle(buffer, n * 4);
                float abs = Math.Abs(sample);
                if (abs > peak) peak = abs;
            }
            return Math.Clamp(peak, 0f, 1f);
        }

        if ((format.Encoding == WaveFormatEncoding.Pcm || format.Encoding == WaveFormatEncoding.Extensible) && format.BitsPerSample == 16)
        {
            int samples = bytesRecorded / 2;
            float peak = 0f;
            for (int n = 0; n < samples; n++)
            {
                short s = BitConverter.ToInt16(buffer, n * 2);
                float abs = Math.Abs(s / 32768f);
                if (abs > peak) peak = abs;
            }
            return Math.Clamp(peak, 0f, 1f);
        }

        return 0f;
    }

    private sealed class DeviceItem
    {
        public MMDevice Device { get; }
        public string Name { get; }
        public bool IsDefault { get; }

        public DeviceItem(MMDevice device)
        {
            Device = device;
            Name = device.FriendlyName;

            try
            {
                using var e = new MMDeviceEnumerator();
                if (device.DataFlow == DataFlow.Capture)
                    IsDefault = string.Equals(device.ID, e.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications).ID, StringComparison.OrdinalIgnoreCase);
                else
                    IsDefault = string.Equals(device.ID, e.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia).ID, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                IsDefault = false;
            }
        }

        public override string ToString() => Name;
    }
}
