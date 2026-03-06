using HyperTool.Models;
using HyperTool.Services;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Markup;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Animation;
using Microsoft.UI.Xaml.Shapes;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Media;
using System.Net.Http;
using System.Reflection;
using Windows.ApplicationModel.DataTransfer;
using Windows.Graphics;
using Windows.Media.Core;
using Windows.Media.Playback;
using Windows.UI;
using IOPath = System.IO.Path;

namespace HyperTool.Guest.Views;

internal sealed class GuestMainWindow : Window
{
    private const int SettingsMenuIndex = 2;

    private enum UsbHostSearchStatusKind
    {
        Neutral,
        Running,
        Success,
        Error
    }

    public const int DefaultWindowWidth = 1061;
    public const int DefaultWindowHeight = 865;
    private const string ToolRestartIcon = "↻";
    private const string ToolRestartLabel = "Tool neu starten";
    private const int GuestSplashMinVisibleMs = 900;
    private const int GuestSplashStatusCycleMs = 420;
    private const string UpdateOwner = "koerby";
    private const string UpdateRepo = "HyperTool";
    private const string GuestInstallerAssetHint = "HyperTool-Guest-Setup";
    private const string GuestUsbRuntimeOwner = "vadimgrn";
    private const string GuestUsbRuntimeRepo = "usbip-win2";
    private const string GuestUsbRuntimeAssetHint = "x64-release";
    private const string WinFspRuntimeOwner = "winfsp";
    private const string WinFspRuntimeRepo = "winfsp";
    private const string WinFspRuntimeAssetHint = ".msi";

    private readonly Func<Task<IReadOnlyList<UsbIpDeviceInfo>>> _refreshUsbDevicesAsync;
    private readonly Func<string, Task<int>> _connectUsbAsync;
    private readonly Func<string, Task<int>> _disconnectUsbAsync;
    private readonly Func<GuestConfig, Task> _saveConfigAsync;
    private readonly Func<Task<GuestConfig>> _reloadConfigAsync;
    private readonly Func<string, Task> _restartForThemeChangeAsync;
    private readonly Func<Task> _exitForUpdateInstallAsync;
    private readonly Func<Task<(bool hyperVSocketActive, bool registryServiceOk)>> _runTransportDiagnosticsTestAsync;
    private readonly Func<Task<string?>> _discoverUsbHostAddressAsync;
    private readonly Func<Task<IReadOnlyList<HostSharedFolderDefinition>>> _fetchHostSharedFoldersAsync;
    private bool _isUsbClientAvailable;
    private readonly IUpdateService _updateService = new GitHubUpdateService();
    private static readonly HttpClient UpdateDownloadClient = new();

    private readonly List<Button> _navButtons = [];
    private readonly ContentPresenter _pageContent = new();
    private readonly Border _overlay = new();
    private readonly Border _overlayCard = new();
    private readonly TextBlock _overlayTitle = new();
    private readonly TextBlock _overlayText = new();
    private readonly ProgressBar _overlayProgressBar = new();
    private TextBlock? _reloadOverlayStatusText;
    private readonly RotateTransform _logoRotateTransform = new();
    private readonly Storyboard _overlayAmbientStoryboard = new();

    private readonly ObservableCollection<string> _notifications = [];
    private readonly Border _notificationSummaryBorder = new();
    private readonly Grid _notificationExpandedGrid = new() { Visibility = Visibility.Collapsed };
    private readonly ListView _notificationsListView = new();
    private readonly TextBlock _statusText = new()
    {
        Text = "Bereit.",
        TextWrapping = TextWrapping.NoWrap,
        TextTrimming = TextTrimming.CharacterEllipsis,
        MaxLines = 1
    };
    private readonly TextBlock _updateStatusValueText = new() { Text = "Noch nicht geprüft", TextWrapping = TextWrapping.Wrap, Opacity = 0.9 };
    private Button? _installUpdateButton;
    private readonly Ellipse _usbRuntimeStatusDot = new() { Width = 10, Height = 10, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock _usbRuntimeStatusText = new() { Opacity = 0.9, VerticalAlignment = VerticalAlignment.Center };
    private readonly Border _usbHostFeatureStatusChip = new() { CornerRadius = new CornerRadius(9), MinHeight = 30, BorderThickness = new Thickness(1), Padding = new Thickness(10, 5, 10, 5), HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock _usbHostFeatureStatusChipText = new() { FontSize = 12, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold };
    private readonly Border _usbPageFeatureStatusChip = new() { CornerRadius = new CornerRadius(9), MinHeight = 26, BorderThickness = new Thickness(1), Padding = new Thickness(8, 4, 8, 4), HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock _usbPageFeatureStatusChipText = new() { FontSize = 12, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold };
    private readonly TextBlock _usbDisabledOverlayText = new() { Text = "Deaktiviert", Opacity = 0.34, FontSize = 34, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, IsHitTestVisible = false };
    private readonly TextBlock _diagHyperVSocketText = new() { Text = "Unbekannt", Opacity = 0.9 };
    private readonly TextBlock _diagRegistryServiceText = new() { Text = "Unbekannt", Opacity = 0.9 };
    private readonly TextBlock _diagFallbackText = new() { Text = "Nein", Opacity = 0.9 };
    private readonly Button _toggleLogButton = new();
    private bool _isLogExpanded;
    private string _releaseUrl = "https://github.com/koerby/HyperTool/releases";
    private string _installerDownloadUrl = string.Empty;
    private string _installerFileName = string.Empty;
    private bool _updateCheckSucceeded;
    private bool _updateAvailable;

    private readonly ListView _usbListView = new();
    private readonly ToggleSwitch _usbFeatureEnabledToggleSwitch = new() { Header = null, OffContent = "", OnContent = "" };
    private readonly ObservableCollection<GuestSharedFolderMapping> _sharedFolderMappings = [];
    private readonly ListView _sharedFolderMappingsListView = new();
    private readonly ToggleSwitch _sharedFolderFeatureEnabledToggleSwitch = new() { Header = null, OffContent = "", OnContent = "" };
    private readonly ComboBox _sharedFolderBaseDriveCombo = new() { Width = 90, MinWidth = 90, HorizontalAlignment = HorizontalAlignment.Left };
    private readonly Border _sharedFolderSettingsPanel = new();
    private Button? _sharedFolderSaveDriveButton;
    private readonly TextBlock _sharedFolderReconnectStatusText = new() { Text = "Reconnect: inaktiv · Letzter Lauf: noch keiner", Opacity = 0.84, TextWrapping = TextWrapping.Wrap };
    private readonly TextBlock _sharedFolderStatusText = new() { Text = "Bereit.", Opacity = 0.88, TextWrapping = TextWrapping.Wrap };
    private readonly Border _sharedFolderHostFeatureStatusChip = new() { CornerRadius = new CornerRadius(9), MinHeight = 30, BorderThickness = new Thickness(1), Padding = new Thickness(10, 5, 10, 5), HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock _sharedFolderHostFeatureStatusChipText = new() { FontSize = 12, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold };
    private readonly Border _sharedFolderPageFeatureStatusChip = new() { CornerRadius = new CornerRadius(9), MinHeight = 26, BorderThickness = new Thickness(1), Padding = new Thickness(8, 4, 8, 4), HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock _sharedFolderPageFeatureStatusChipText = new() { FontSize = 12, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold };
    private readonly TextBlock _sharedFolderDisabledOverlayText = new() { Text = "Deaktiviert", Opacity = 0.34, FontSize = 34, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, IsHitTestVisible = false };
    private readonly Ellipse _winFspRuntimeStatusDot = new() { Width = 10, Height = 10, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock _winFspRuntimeStatusText = new() { Opacity = 0.9, VerticalAlignment = VerticalAlignment.Center };
    private Button? _winFspRuntimeInstallButton;
    private Button? _winFspRuntimeRestartButton;
    private readonly GuestWinFspMountService _winFspMountService = GuestWinFspMountRegistry.Instance;
    private readonly SemaphoreSlim _sharedFolderUiOperationGate = new(1, 1);
    private CancellationTokenSource? _sharedFolderAutoApplyCts;
    private bool _suppressSharedFolderAutoApply;
    private bool _suppressUsbFeatureToggle;
    private bool _suppressSharedFolderFeatureToggle;
    private bool _suppressSharedFolderBaseDriveChange;
    private string _sharedFolderLastError = "-";
    private string _savedSharedFolderBaseDriveLetter = "Z";
    private string _pendingSharedFolderBaseDriveLetter = "Z";
    private bool _usbHostFeatureReactivationPending;
    private bool _sharedFolderHostFeatureReactivationPending;
    private int _usbHostFeatureReactivationToken;
    private int _sharedFolderHostFeatureReactivationToken;
    private Button? _usbRefreshButton;
    private Button? _usbConnectButton;
    private Button? _usbDisconnectButton;
    private Button? _usbRuntimeInstallButton;
    private Button? _usbRuntimeRestartButton;
    private readonly ComboBox _themeCombo = new();
    private readonly ToggleSwitch _themeToggle = new();
    private readonly TextBlock _themeText = new();
    private readonly TextBox _usbHostAddressTextBox = new();
    private Button? _usbHostSearchButton;
    private UsbHostSearchStatusKind _usbHostSearchStatusKind = UsbHostSearchStatusKind.Neutral;
    private readonly TextBlock _usbHostSearchStatusText = new()
    {
        Opacity = 0.88,
        VerticalAlignment = VerticalAlignment.Center,
        TextWrapping = TextWrapping.NoWrap,
        Text = "Bereit"
    };
    private readonly TextBlock _usbResolvedHostNameText = new()
    {
        Opacity = 0.88,
        TextWrapping = TextWrapping.Wrap,
        Text = "Ermittelter Hostname: -"
    };
    private readonly CheckBox _startWithWindowsCheckBox = new() { Content = "Mit Windows starten" };
    private readonly CheckBox _startMinimizedCheckBox = new() { Content = "Beim Start minimiert" };
    private readonly CheckBox _minimizeToTrayCheckBox = new() { Content = "Tasktray-Menü aktiv" };
    private readonly CheckBox _checkForUpdatesOnStartupCheckBox = new() { Content = "Beim Start auf Updates prüfen" };
    private readonly CheckBox _useHyperVSocketCheckBox = new() { Content = "Hyper-V Socket verwenden (bevorzugt)" };
    private readonly CheckBox _usbAutoConnectCheckBox = new() { Content = "Auto-Connect für ausgewähltes Gerät" };
    private readonly CheckBox _usbDisconnectOnExitCheckBox = new() { Content = "Disconnect beim Beenden" };
    private readonly Border _usbHostAddressEditorCard = new()
    {
        BorderThickness = new Thickness(1),
        CornerRadius = new CornerRadius(8),
        Padding = new Thickness(10, 8, 10, 8)
    };
    private readonly TextBlock _usbModeHintText = new() { TextWrapping = TextWrapping.Wrap, Opacity = 0.88 };
    private readonly Border _usbTransportModeBadge = new()
    {
        CornerRadius = new CornerRadius(9),
        Padding = new Thickness(10, 5, 10, 5),
        MinHeight = 30,
        BorderThickness = new Thickness(1),
        HorizontalAlignment = HorizontalAlignment.Right,
        VerticalAlignment = VerticalAlignment.Center
    };
    private readonly TextBlock _usbTransportModeBadgeText = new()
    {
        FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
        FontSize = 12,
        Text = "Modus: -"
    };
    private readonly Border _usbHyperVModeBadge = new()
    {
        Width = 152,
        MinHeight = 30,
        CornerRadius = new CornerRadius(9),
        Padding = new Thickness(10, 5, 10, 5),
        BorderThickness = new Thickness(1)
    };
    private readonly Border _usbHyperVModeIconBadge = new()
    {
        Width = 18,
        Height = 18,
        CornerRadius = new CornerRadius(5),
        BorderThickness = new Thickness(1),
        VerticalAlignment = VerticalAlignment.Center
    };
    private readonly SymbolIcon _usbHyperVModeIcon = new()
    {
        Symbol = Symbol.Switch,
        Width = 12,
        Height = 12
    };
    private readonly TextBlock _usbHyperVModeBadgeText = new()
    {
        FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
        FontSize = 12,
        Text = "Hyper-V Socket"
    };
    private readonly Border _usbIpModeBadge = new()
    {
        Width = 152,
        MinHeight = 30,
        CornerRadius = new CornerRadius(9),
        Padding = new Thickness(10, 5, 10, 5),
        BorderThickness = new Thickness(1)
    };
    private readonly Border _usbIpModeIconBadge = new()
    {
        Width = 18,
        Height = 18,
        CornerRadius = new CornerRadius(5),
        BorderThickness = new Thickness(1),
        VerticalAlignment = VerticalAlignment.Center
    };
    private readonly SymbolIcon _usbIpModeIcon = new()
    {
        Symbol = Symbol.World,
        Width = 12,
        Height = 12
    };
    private readonly TextBlock _usbIpModeBadgeText = new()
    {
        FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
        FontSize = 12,
        Text = "IP-Mode"
    };

    private HelpWindow? _helpWindow;
    private Grid? _headerGrid;
    private StackPanel? _headerTitleStack;
    private StackPanel? _headerStatusChipsRow;
    private StackPanel? _headerTitleActions;
    private TextBlock? _headerSubtitleText;
    private bool _isCompactHeaderLayout;
    private int _selectedMenuIndex;
    private GuestConfig _config;
    private IReadOnlyList<UsbIpDeviceInfo> _usbDevices = [];
    private bool _suppressThemeEvents;
    private bool _isThemeRestartInProgress;
    private bool _isThemeToggleHandlerAttached;
    private bool _isUsbTransportToggleHandlerAttached;
    private bool _isUsbModeBadgeHandlersAttached;
    private bool _suppressUsbTransportToggleEvents;
    private bool _suppressUsbAutoConnectToggleEvents;
    private bool _suppressUsbDisconnectOnExitToggleEvents;
    private bool _isHandlingMenuSwitch;
    private bool _isMenuSwitchPromptOpen;
    private CancellationTokenSource? _usbTransportAutoRefreshCts;
    private MediaPlayer? _logoSpinPlayer;
    private List<UIElement>? _startupMainElements;
    private UIElement? _usbPage;
    private UIElement? _sharedFoldersPage;
    private UIElement? _settingsPage;
    private UIElement? _infoPage;

    public GuestMainWindow(
        GuestConfig config,
        Func<Task<IReadOnlyList<UsbIpDeviceInfo>>> refreshUsbDevicesAsync,
        Func<string, Task<int>> connectUsbAsync,
        Func<string, Task<int>> disconnectUsbAsync,
        Func<GuestConfig, Task> saveConfigAsync,
        Func<Task<GuestConfig>> reloadConfigAsync,
        Func<string, Task> restartForThemeChangeAsync,
        Func<Task> exitForUpdateInstallAsync,
        Func<Task<(bool hyperVSocketActive, bool registryServiceOk)>> runTransportDiagnosticsTestAsync,
        Func<Task<string?>> discoverUsbHostAddressAsync,
        Func<Task<IReadOnlyList<HostSharedFolderDefinition>>> fetchHostSharedFoldersAsync,
        bool isUsbClientAvailable)
    {
        _config = config;
        _refreshUsbDevicesAsync = refreshUsbDevicesAsync;
        _connectUsbAsync = connectUsbAsync;
        _disconnectUsbAsync = disconnectUsbAsync;
        _saveConfigAsync = saveConfigAsync;
        _reloadConfigAsync = reloadConfigAsync;
        _restartForThemeChangeAsync = restartForThemeChangeAsync;
        _exitForUpdateInstallAsync = exitForUpdateInstallAsync;
        _runTransportDiagnosticsTestAsync = runTransportDiagnosticsTestAsync;
        _discoverUsbHostAddressAsync = discoverUsbHostAddressAsync;
        _fetchHostSharedFoldersAsync = fetchHostSharedFoldersAsync;
        _isUsbClientAvailable = isUsbClientAvailable;

        Title = "HyperTool Guest";
        ExtendsContentIntoTitleBar = false;
        TryApplyInitialWindowSize();

        var initialTheme = GuestConfigService.NormalizeTheme(config.Ui.Theme);
        ApplyThemePalette(initialTheme == "dark");

        _themeCombo.Items.Add("dark");
        _themeCombo.Items.Add("light");

        Content = BuildLayout();
        ApplyConfigToControls();
        ApplyTheme(config.Ui.Theme);
        TryApplyWindowIcon();
        SizeChanged += (_, _) => UpdateHeaderResponsiveLayout();
        UpdateHeaderResponsiveLayout();

        GuestLogger.EntryWritten += OnLoggerEntryWritten;
        Closed += (_, _) => GuestLogger.EntryWritten -= OnLoggerEntryWritten;
    }

    public string CurrentTheme => GuestConfigService.NormalizeTheme((_themeCombo.SelectedItem as string) ?? _config.Ui.Theme);

    public int SelectedMenuIndex => _selectedMenuIndex;

    public void SelectMenuIndex(int index)
    {
        _ = SelectMenuIndexAsync(index);
    }

    private async Task SelectMenuIndexAsync(int index)
    {
        if (_isHandlingMenuSwitch)
        {
            return;
        }

        var normalized = Math.Clamp(index, 0, 3);
        if (_selectedMenuIndex == normalized)
        {
            return;
        }

        _isHandlingMenuSwitch = true;
        try
        {
            if (!await TryHandlePendingSettingsBeforeMenuSwitchAsync(normalized))
            {
                return;
            }

            _selectedMenuIndex = normalized;
            UpdateNavSelection();
            UpdatePageContent();
        }
        finally
        {
            _isHandlingMenuSwitch = false;
        }
    }

    public void ApplyTheme(string theme)
    {
        var normalized = GuestConfigService.NormalizeTheme(theme);
        var isDark = normalized == "dark";

        ApplyThemePalette(isDark);

        if (Content is FrameworkElement root)
        {
            root.RequestedTheme = isDark ? ElementTheme.Dark : ElementTheme.Light;
        }

        _suppressThemeEvents = true;
        _themeCombo.SelectedItem = normalized;
        _themeToggle.IsOn = isDark;
        _suppressThemeEvents = false;
        _themeText.Text = isDark ? "Dunkles Theme" : "Helles Theme";

        _overlay.Background = HyperTool.WinUI.Views.LifecycleVisuals.CreateRootBackgroundBrush();

        _overlayCard.Background = HyperTool.WinUI.Views.LifecycleVisuals.CreateCardSurfaceBrush();
        _overlayCard.BorderBrush = new SolidColorBrush(HyperTool.WinUI.Views.LifecycleVisuals.CardBorder);
        _overlayTitle.Foreground = new SolidColorBrush(HyperTool.WinUI.Views.LifecycleVisuals.TextPrimary);
        _overlayText.Foreground = new SolidColorBrush(HyperTool.WinUI.Views.LifecycleVisuals.TextSecondary);

        ApplyUsbHostSearchStatusColor();

        UpdateTitleBarAppearance(isDark);
    }

    public async Task PlayStartupAnimationAsync()
    {
        var mainElements = _startupMainElements;
        if (mainElements is null && Content is Grid rootGrid)
        {
            mainElements = rootGrid.Children
                .OfType<UIElement>()
                .Where(element => !ReferenceEquals(element, _overlay))
                .ToList();

            foreach (var element in mainElements)
            {
                element.Opacity = 0;
            }
        }

        _overlay.Visibility = Visibility.Visible;
        _overlay.Opacity = 1;
        _overlayProgressBar.Value = 8;

        var startupStart = Stopwatch.StartNew();
        var statusIndex = 0;
        while (startupStart.ElapsedMilliseconds < GuestSplashMinVisibleMs)
        {
            var status = HyperTool.WinUI.Views.LifecycleVisuals.StartupStatusMessages[
                statusIndex % HyperTool.WinUI.Views.LifecycleVisuals.StartupStatusMessages.Length];
            _overlayText.Text = status;

            _overlayProgressBar.Value = Math.Min(92, 8 + (statusIndex * 17));
            statusIndex++;

            var remaining = GuestSplashMinVisibleMs - startupStart.ElapsedMilliseconds;
            var delay = (int)Math.Min(GuestSplashStatusCycleMs, remaining);
            if (delay > 0)
            {
                await Task.Delay(delay);
            }
        }

        _overlayText.Text = "Starte HyperTool Guest Oberfläche …";
        _overlayProgressBar.Value = 100;

        var fadeOut = new DoubleAnimation
        {
            From = 1,
            To = 0,
            Duration = TimeSpan.FromMilliseconds(420),
            EasingFunction = HyperTool.WinUI.Views.LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };

        var story = new Storyboard();
        Storyboard.SetTarget(fadeOut, _overlay);
        Storyboard.SetTargetProperty(fadeOut, "Opacity");
        story.Children.Add(fadeOut);

        if (mainElements is not null)
        {
            foreach (var element in mainElements)
            {
                var fadeInMain = new DoubleAnimation
                {
                    From = 0,
                    To = 1,
                    Duration = TimeSpan.FromMilliseconds(420),
                    EasingFunction = HyperTool.WinUI.Views.LifecycleVisuals.CreateEaseOut(),
                    EnableDependentAnimation = true
                };

                Storyboard.SetTarget(fadeInMain, element);
                Storyboard.SetTargetProperty(fadeInMain, "Opacity");
                story.Children.Add(fadeInMain);
            }
        }

        story.Begin();

        await Task.Delay(460);
        _overlay.Visibility = Visibility.Collapsed;
        _startupMainElements = null;
    }

    public void PrepareStartupSplash()
    {
        if (Content is not Grid rootGrid)
        {
            return;
        }

        _startupMainElements = rootGrid.Children
            .OfType<UIElement>()
            .Where(element => !ReferenceEquals(element, _overlay))
            .ToList();

        foreach (var element in _startupMainElements)
        {
            element.Opacity = 0;
        }

        _overlay.Visibility = Visibility.Visible;
        _overlay.Opacity = 1;
        _overlayText.Text = HyperTool.WinUI.Views.LifecycleVisuals.StartupStatusMessages[0];
        _overlayProgressBar.Value = 8;
    }

    public void PrepareLifecycleGuard(string statusText)
    {
        if (Content is not Grid rootGrid)
        {
            return;
        }

        _startupMainElements = rootGrid.Children
            .OfType<UIElement>()
            .Where(element => !ReferenceEquals(element, _overlay))
            .ToList();

        foreach (var element in _startupMainElements)
        {
            element.Opacity = 0;
        }

        _overlay.Visibility = Visibility.Visible;
        _overlay.Opacity = 1;
        _overlayText.Text = string.IsNullOrWhiteSpace(statusText)
            ? "Design wird neu geladen …"
            : statusText.Trim();
        _overlayProgressBar.Value = 64;
    }

    public async Task DismissLifecycleGuardAsync()
    {
        var mainElements = _startupMainElements;
        var story = new Storyboard();

        var fadeOut = new DoubleAnimation
        {
            From = _overlay.Opacity,
            To = 0,
            Duration = TimeSpan.FromMilliseconds(220),
            EasingFunction = HyperTool.WinUI.Views.LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(fadeOut, _overlay);
        Storyboard.SetTargetProperty(fadeOut, "Opacity");
        story.Children.Add(fadeOut);

        if (mainElements is not null)
        {
            foreach (var element in mainElements)
            {
                var fadeInMain = new DoubleAnimation
                {
                    From = element.Opacity,
                    To = 1,
                    Duration = TimeSpan.FromMilliseconds(220),
                    EasingFunction = HyperTool.WinUI.Views.LifecycleVisuals.CreateEaseOut(),
                    EnableDependentAnimation = true
                };

                Storyboard.SetTarget(fadeInMain, element);
                Storyboard.SetTargetProperty(fadeInMain, "Opacity");
                story.Children.Add(fadeInMain);
            }
        }

        story.Begin();
        await Task.Delay(250);

        _overlay.Visibility = Visibility.Collapsed;
        _startupMainElements = null;
    }

    public async Task PlayExitAnimationAsync()
    {
        _overlay.Visibility = Visibility.Visible;
        _overlay.Opacity = 0;
        _overlayText.Text = "Guest-Dienste werden sicher beendet …";

        var fadeIn = new DoubleAnimation
        {
            From = 0,
            To = 1,
            Duration = TimeSpan.FromMilliseconds(300),
            EnableDependentAnimation = true
        };

        var story = new Storyboard();
        Storyboard.SetTarget(fadeIn, _overlay);
        Storyboard.SetTargetProperty(fadeIn, "Opacity");
        story.Children.Add(fadeIn);
        story.Begin();

        await Task.Delay(330);
    }

    public async Task PlayExitFadeAsync()
    {
        if (Content is not UIElement contentElement)
        {
            return;
        }

        var fadeOut = new DoubleAnimation
        {
            From = 1.0,
            To = 0.0,
            Duration = TimeSpan.FromMilliseconds(220),
            EnableDependentAnimation = true
        };

        var storyboard = new Storyboard();
        Storyboard.SetTarget(fadeOut, contentElement);
        Storyboard.SetTargetProperty(fadeOut, "Opacity");
        storyboard.Children.Add(fadeOut);
        storyboard.Begin();

        await Task.Delay(240);
    }

    public async Task PlayThemeReloadSplashAsync(string targetTheme)
    {
        try
        {
            _overlayAmbientStoryboard.Stop();
        }
        catch
        {
        }

        var reloadOverlay = BuildReloadOverlayContent();
        _overlay.Child = reloadOverlay.Root;
        _reloadOverlayStatusText = reloadOverlay.StatusText;
        _overlay.Visibility = Visibility.Visible;
        _overlay.Opacity = 0;
        _reloadOverlayStatusText.Text = "Layout wird aktualisiert …";

        var fadeIn = new DoubleAnimation
        {
            From = 0,
            To = 1,
            Duration = TimeSpan.FromMilliseconds(130),
            EasingFunction = HyperTool.WinUI.Views.LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };

        var showStoryboard = new Storyboard();
        Storyboard.SetTarget(fadeIn, _overlay);
        Storyboard.SetTargetProperty(fadeIn, "Opacity");
        showStoryboard.Children.Add(fadeIn);
        showStoryboard.Begin();

        await Task.Delay(154);

        _reloadOverlayStatusText.Text = "Design wird neu geladen …";
        await Task.Delay(1000);
    }

    private void TryApplyInitialWindowSize()
    {
        try
        {
            if (AppWindow is not null)
            {
                AppWindow.Resize(new SizeInt32(DefaultWindowWidth, DefaultWindowHeight));
            }
        }
        catch
        {
        }
    }

    public void UpdateUsbDevices(IReadOnlyList<UsbIpDeviceInfo> devices)
    {
        var selectedBeforeRefresh = GetSelectedUsbDevice();
        var selectedKeyBeforeRefresh = BuildUsbSelectionKey(selectedBeforeRefresh);

        _usbDevices = devices;

        if (!_isUsbClientAvailable)
        {
            _usbListView.ItemsSource = new[]
            {
                "USB/IP-Client nicht installiert. USB-Funktionen sind deaktiviert."
            };
            _usbListView.SelectedIndex = -1;
            return;
        }

        if (devices.Count == 0)
        {
            _usbListView.ItemsSource = new[]
            {
                "Aktuell keine Geräte zum Connecten vorhanden."
            };
            _usbListView.SelectedIndex = -1;
            UpdateAutoConnectToggleFromSelection();
            return;
        }

        _usbListView.ItemsSource = devices.Select(item => item.DisplayName).ToList();

        if (!string.IsNullOrWhiteSpace(selectedKeyBeforeRefresh))
        {
            var restoredIndex = devices
                .Select((device, index) => new { device, index })
                .FirstOrDefault(entry => string.Equals(BuildUsbSelectionKey(entry.device), selectedKeyBeforeRefresh, StringComparison.OrdinalIgnoreCase))
                ?.index ?? -1;

            _usbListView.SelectedIndex = restoredIndex;
        }
        else if (_usbListView.SelectedIndex < 0 && devices.Count > 0)
        {
            _usbListView.SelectedIndex = 0;
        }

        UpdateAutoConnectToggleFromSelection();
    }

    private static string BuildUsbSelectionKey(UsbIpDeviceInfo? device)
    {
        if (device is null)
        {
            return string.Empty;
        }

        if (!string.IsNullOrWhiteSpace(device.BusId))
        {
            return "bus:" + device.BusId.Trim();
        }

        if (!string.IsNullOrWhiteSpace(device.HardwareId))
        {
            return "hw:" + device.HardwareId.Trim();
        }

        if (!string.IsNullOrWhiteSpace(device.Description))
        {
            return "desc:" + device.Description.Trim();
        }

        return string.Empty;
    }

    public void UpdateUsbClientAvailability(bool isUsbClientAvailable)
    {
        _isUsbClientAvailable = isUsbClientAvailable;
        UpdateUsbRuntimeStatusUi();
    }

    public UsbIpDeviceInfo? GetSelectedUsbDevice()
    {
        if (_usbListView.SelectedIndex < 0 || _usbListView.SelectedIndex >= _usbDevices.Count)
        {
            return null;
        }

        return _usbDevices[_usbListView.SelectedIndex];
    }

    public void UpdateTransportDiagnostics(bool hyperVSocketActive, bool registryServicePresent, bool fallbackActive)
    {
        _diagHyperVSocketText.Text = hyperVSocketActive ? "Ja" : "Nein";
        _diagRegistryServiceText.Text = registryServicePresent ? "Ja" : "Nein";
        _diagFallbackText.Text = fallbackActive ? "Ja" : "Nein";
        UpdateUsbTransportHeaderStatus();
    }

    public void UpdateSharedFolderReconnectStatus(bool reconnectActive, DateTimeOffset? lastRunUtc, string summary)
    {
        var lastRunText = lastRunUtc.HasValue
            ? lastRunUtc.Value.ToLocalTime().ToString("HH:mm:ss")
            : "noch keiner";
        var activeText = reconnectActive ? "aktiv" : "inaktiv";
        var normalizedSummary = string.IsNullOrWhiteSpace(summary) ? "-" : summary.Trim();

        _sharedFolderReconnectStatusText.Text = $"Reconnect: {activeText} · Letzter Lauf: {lastRunText} · {normalizedSummary}";
    }

    private async Task RunTransportDiagnosticsTestAsync()
    {
        try
        {
            var (hyperVSocketActive, registryServiceOk) = await _runTransportDiagnosticsTestAsync();
            var ok = hyperVSocketActive && registryServiceOk;

            AppendNotification(ok
                ? "[Info] Hyper-V Socket Test erfolgreich: Verbindung steht und Registry-Service ist erreichbar."
                : $"[Warn] Hyper-V Socket Test: Verbindung={(hyperVSocketActive ? "OK" : "FAIL")}, Registry-Service={(registryServiceOk ? "OK" : "FAIL")}. ");
        }
        catch (Exception ex)
        {
            AppendNotification($"[Error] Hyper-V Socket Test fehlgeschlagen: {ex.Message}");
        }
    }

    private UIElement BuildLayout()
    {
        var root = new Grid
        {
            Background = Application.Current.Resources["PageBackgroundBrush"] as Brush
        };

        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(12) });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(12) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var headerCard = CreateCard(new Thickness(16, 16, 16, 0), 14, 14);
        headerCard.Child = BuildHeader();
        Grid.SetRow(headerCard, 0);
        root.Children.Add(headerCard);

        var contentCard = CreateCard(new Thickness(16, 0, 16, 0), 12, 14);
        contentCard.Child = BuildMainContentGrid();
        Grid.SetRow(contentCard, 2);
        root.Children.Add(contentCard);

        var bottomCard = CreateCard(new Thickness(16, 0, 16, 16), 12, 12);
        bottomCard.Child = BuildBottomArea();
        Grid.SetRow(bottomCard, 4);
        root.Children.Add(bottomCard);

        _overlay.Visibility = Visibility.Collapsed;
        _overlay.Child = BuildOverlayContent();
        Grid.SetRow(_overlay, 0);
        Grid.SetRowSpan(_overlay, 5);
        root.Children.Add(_overlay);

        UpdateNavSelection();
        UpdatePageContent();
        UpdateBusyAndNotificationPanel();

        return root;
    }

    private Grid BuildHeader()
    {
        var headerGrid = new Grid();
        _headerGrid = headerGrid;
        headerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        headerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var titleStack = new StackPanel { Orientation = Orientation.Vertical, Spacing = 2 };
        _headerTitleStack = titleStack;
        var titleRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, VerticalAlignment = VerticalAlignment.Center };
        titleRow.Children.Add(new Image
        {
            Source = new Microsoft.UI.Xaml.Media.Imaging.BitmapImage(new Uri("ms-appx:///Assets/HyperTool.Guest.Icon.Transparent.png")),
            Width = 28,
            Height = 28
        });
        titleRow.Children.Add(new TextBlock
        {
            Text = "HyperTool Guest",
            FontSize = 24,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
        });
        titleStack.Children.Add(titleRow);
        _headerSubtitleText = new TextBlock
        {
            Text = "dein nützlicher Hyper V Helfer",
            Opacity = 0.8,
            Margin = new Thickness(0, 0, 0, 6)
        };
        titleStack.Children.Add(_headerSubtitleText);

        Grid.SetRow(titleStack, 0);
        Grid.SetColumn(titleStack, 0);
        headerGrid.Children.Add(titleStack);

        _usbHostFeatureStatusChip.Child = _usbHostFeatureStatusChipText;
        _sharedFolderHostFeatureStatusChip.Child = _sharedFolderHostFeatureStatusChipText;
        _usbTransportModeBadge.Child = _usbTransportModeBadgeText;

        var statusChipsRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 8,
            Margin = new Thickness(12, 2, 12, 0),
            VerticalAlignment = VerticalAlignment.Top
        };
        _headerStatusChipsRow = statusChipsRow;
        statusChipsRow.Children.Add(_usbHostFeatureStatusChip);
        statusChipsRow.Children.Add(_sharedFolderHostFeatureStatusChip);
        statusChipsRow.Children.Add(_usbTransportModeBadge);
        Grid.SetRow(statusChipsRow, 0);
        Grid.SetColumn(statusChipsRow, 1);
        headerGrid.Children.Add(statusChipsRow);

        var titleActions = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 10,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Top
        };
        _headerTitleActions = titleActions;

        var helpButton = new Button
        {
            Width = 54,
            Height = 54,
            CornerRadius = new CornerRadius(12),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            Padding = new Thickness(0),
            Content = new TextBlock
            {
                Text = "?",
                FontSize = 28,
                FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                TextAlignment = TextAlignment.Center
            }
        };
        helpButton.Click += (_, _) => OpenHelpWindow();
        titleActions.Children.Add(helpButton);

        var logoBorder = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            CornerRadius = new CornerRadius(12),
            Width = 54,
            Height = 54,
            Padding = new Thickness(6)
        };
        var logo = new Image
        {
            Source = new Microsoft.UI.Xaml.Media.Imaging.BitmapImage(new Uri("ms-appx:///Assets/Logo.png")),
            Width = 40,
            Height = 40,
            RenderTransform = _logoRotateTransform,
            RenderTransformOrigin = new Windows.Foundation.Point(0.5, 0.5)
        };
        logoBorder.Child = logo;
        logoBorder.Tapped += (_, _) => RunLogoEasterEgg();
        titleActions.Children.Add(logoBorder);

        Grid.SetRow(titleActions, 0);
        Grid.SetColumn(titleActions, 2);
        headerGrid.Children.Add(titleActions);

        UpdateHeaderResponsiveLayout();
        return headerGrid;
    }

    private void UpdateHeaderResponsiveLayout()
    {
        if (_headerGrid is null || _headerStatusChipsRow is null || _headerTitleStack is null || _headerTitleActions is null)
        {
            return;
        }

        var windowWidth = (Content as FrameworkElement)?.ActualWidth ?? DefaultWindowWidth;
        var titleWidth = _headerTitleStack.ActualWidth;
        var chipsWidth = _headerStatusChipsRow.ActualWidth;
        var actionsWidth = _headerTitleActions.ActualWidth;

        var hasMeasuredHeaderWidths = titleWidth > 0 && chipsWidth > 0 && actionsWidth > 0;
        var requiredTopRowWidth = titleWidth + chipsWidth + actionsWidth + 72;
        var compact = hasMeasuredHeaderWidths
            ? windowWidth < requiredTopRowWidth
            : windowWidth > 0 && windowWidth < 1160;
        _isCompactHeaderLayout = compact;

        if (_headerSubtitleText is not null)
        {
            _headerSubtitleText.Visibility = compact ? Visibility.Collapsed : Visibility.Visible;
        }

        _usbHostFeatureStatusChip.Padding = compact ? new Thickness(8, 4, 8, 4) : new Thickness(10, 5, 10, 5);
        _usbHostFeatureStatusChip.MinHeight = compact ? 26 : 30;
        _usbHostFeatureStatusChipText.FontSize = compact ? 11 : 12;

        _sharedFolderHostFeatureStatusChip.Padding = compact ? new Thickness(8, 4, 8, 4) : new Thickness(10, 5, 10, 5);
        _sharedFolderHostFeatureStatusChip.MinHeight = compact ? 26 : 30;
        _sharedFolderHostFeatureStatusChipText.FontSize = compact ? 11 : 12;

        _usbTransportModeBadge.Padding = compact ? new Thickness(8, 4, 8, 4) : new Thickness(10, 5, 10, 5);
        _usbTransportModeBadge.MinHeight = compact ? 26 : 30;
        _usbTransportModeBadgeText.FontSize = compact ? 11 : 12;

        _headerStatusChipsRow.Orientation = Orientation.Horizontal;
        _headerStatusChipsRow.Spacing = compact ? 7 : 8;
        _headerStatusChipsRow.Margin = compact ? new Thickness(0, 8, 0, 0) : new Thickness(0, 6, 0, 0);

        Grid.SetRow(_headerTitleStack, 0);
        Grid.SetColumn(_headerTitleStack, 0);
        Grid.SetColumnSpan(_headerTitleStack, 1);

        Grid.SetRow(_headerTitleActions, 0);
        Grid.SetColumn(_headerTitleActions, 2);

        Grid.SetRow(_headerStatusChipsRow, 1);
        Grid.SetColumn(_headerStatusChipsRow, 0);
        Grid.SetColumnSpan(_headerStatusChipsRow, 3);
        _headerStatusChipsRow.HorizontalAlignment = HorizontalAlignment.Left;

        UpdateUsbTransportHeaderStatus();
    }

    private Grid BuildMainContentGrid()
    {
        var mainGrid = new Grid();
        mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(126) });
        mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(12) });
        mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var sidebar = new Border
        {
            CornerRadius = new CornerRadius(12),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            Padding = new Thickness(10)
        };

        var sidebarStack = new StackPanel { Spacing = 8 };
        sidebarStack.Children.Add(CreateNavButton("🔌", "USB", 0));
        sidebarStack.Children.Add(CreateNavButton("📁", "Shared Folder", 1));
        sidebarStack.Children.Add(CreateNavButton("⚙", "Einstellungen", 2));
        sidebarStack.Children.Add(CreateNavButton("ℹ", "Info", 3));
        sidebar.Child = sidebarStack;

        mainGrid.Children.Add(sidebar);

        var contentGrid = new Grid();
        contentGrid.Children.Add(_pageContent);

        Grid.SetColumn(contentGrid, 2);
        mainGrid.Children.Add(contentGrid);

        return mainGrid;
    }

    private UIElement BuildBottomArea()
    {
        var bottom = new Grid { RowSpacing = 8 };
        bottom.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        bottom.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        bottom.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var topRow = new Grid { ColumnSpacing = 8 };
        topRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        topRow.Children.Add(new TextBlock { Text = "Notifications", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, VerticalAlignment = VerticalAlignment.Center });
        bottom.Children.Add(topRow);

        var summaryGrid = new Grid { ColumnSpacing = 8 };
        summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        _notificationSummaryBorder.Padding = new Thickness(10);
        _notificationSummaryBorder.CornerRadius = new CornerRadius(8);
        _notificationSummaryBorder.BorderThickness = new Thickness(1);
        _notificationSummaryBorder.BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush;
        _notificationSummaryBorder.Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush;
        _notificationSummaryBorder.Child = _statusText;
        summaryGrid.Children.Add(_notificationSummaryBorder);

        var summaryButtons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        var openLogButton = CreateIconButton("📄", "Log-Ordner öffnen", onClick: (_, _) => OpenLogFile());
        openLogButton.CornerRadius = new CornerRadius(8);
        openLogButton.Padding = new Thickness(8, 2, 8, 2);
        openLogButton.BorderThickness = new Thickness(1);
        openLogButton.BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush;
        openLogButton.Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush;
        summaryButtons.Children.Add(openLogButton);

        _toggleLogButton.CornerRadius = new CornerRadius(8);
        _toggleLogButton.Padding = new Thickness(8, 2, 8, 2);
        _toggleLogButton.BorderThickness = new Thickness(1);
        _toggleLogButton.BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush;
        _toggleLogButton.Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush;
        _toggleLogButton.Click += (_, _) =>
        {
            _isLogExpanded = !_isLogExpanded;
            UpdateBusyAndNotificationPanel();
        };
        summaryButtons.Children.Add(_toggleLogButton);

        Grid.SetColumn(summaryButtons, 1);
        summaryGrid.Children.Add(summaryButtons);

        Grid.SetRow(summaryGrid, 1);
        bottom.Children.Add(summaryGrid);

        _notificationExpandedGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        _notificationExpandedGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(8) });
        _notificationExpandedGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var expandedButtons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        expandedButtons.Children.Add(CreateIconButton("⧉", "Copy", onClick: (_, _) => CopyNotificationsToClipboard()));
        expandedButtons.Children.Add(CreateIconButton("⌫", "Clear", onClick: (_, _) =>
        {
            _notifications.Clear();
            _statusText.Text = "Keine Notifications.";
        }));
        Grid.SetRow(expandedButtons, 0);
        _notificationExpandedGrid.Children.Add(expandedButtons);

        var logListBorder = new Border
        {
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(8),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush
        };
        _notificationsListView.ItemsSource = _notifications;
        _notificationsListView.MaxHeight = 220;
        logListBorder.Child = _notificationsListView;
        Grid.SetRow(logListBorder, 2);
        _notificationExpandedGrid.Children.Add(logListBorder);
        Grid.SetRow(_notificationExpandedGrid, 2);
        bottom.Children.Add(_notificationExpandedGrid);

        return bottom;
    }

    private UIElement BuildOverlayContent()
    {
        var overlayRoot = new Grid
        {
            Background = HyperTool.WinUI.Views.LifecycleVisuals.CreateRootBackgroundBrush()
        };

        var focusLayerPrimary = new Ellipse
        {
            Width = 820,
            Height = 820,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Fill = HyperTool.WinUI.Views.LifecycleVisuals.CreateCenterFocusBrush(HyperTool.WinUI.Views.LifecycleVisuals.BackgroundFocusSecondary)
        };
        overlayRoot.Children.Add(focusLayerPrimary);

        var focusLayerSecondary = new Ellipse
        {
            Width = 620,
            Height = 620,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(88, -52, -88, 52),
            Fill = HyperTool.WinUI.Views.LifecycleVisuals.CreateCenterFocusBrush(HyperTool.WinUI.Views.LifecycleVisuals.BackgroundFocusTertiary),
            Opacity = 0.54
        };
        overlayRoot.Children.Add(focusLayerSecondary);

        overlayRoot.Children.Add(new Rectangle
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment = VerticalAlignment.Stretch,
            Fill = HyperTool.WinUI.Views.LifecycleVisuals.CreateVignetteBrush(0x78)
        });

        BuildOverlayNetworkLayer(overlayRoot);
        BuildOverlayAmbientBands(overlayRoot);

        var splashVersionText = new TextBlock
        {
            Text = HyperTool.WinUI.Views.LifecycleVisuals.ResolveDisplayVersion(ResolveGuestVersionText()),
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Bottom,
            Margin = new Thickness(0, 0, 20, 14),
            FontSize = 12,
            Opacity = 0.72,
            Foreground = new SolidColorBrush(Color.FromArgb(0xC8, 0x9B, 0xB7, 0xD7))
        };
        overlayRoot.Children.Add(splashVersionText);

        var splashCopyrightText = new TextBlock
        {
            Text = "Copyright: koerby",
            HorizontalAlignment = HorizontalAlignment.Left,
            VerticalAlignment = VerticalAlignment.Bottom,
            Margin = new Thickness(20, 0, 0, 14),
            FontSize = 12,
            Opacity = 0.72,
            Foreground = new SolidColorBrush(Color.FromArgb(0xC8, 0x9B, 0xB7, 0xD7))
        };
        overlayRoot.Children.Add(splashCopyrightText);

        _overlayCard.Width = 520;
        _overlayCard.Height = double.NaN;
        _overlayCard.Padding = new Thickness(30, 28, 30, 24);
        _overlayCard.CornerRadius = new CornerRadius(24);
        _overlayCard.HorizontalAlignment = HorizontalAlignment.Center;
        _overlayCard.VerticalAlignment = VerticalAlignment.Center;
        _overlayCard.Background = HyperTool.WinUI.Views.LifecycleVisuals.CreateCardSurfaceBrush();
        _overlayCard.BorderBrush = new SolidColorBrush(HyperTool.WinUI.Views.LifecycleVisuals.CardBorder);
        _overlayCard.BorderThickness = new Thickness(1);
        _overlayCard.Shadow = new ThemeShadow();

        var innerFrame = new Border
        {
            CornerRadius = new CornerRadius(20),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(HyperTool.WinUI.Views.LifecycleVisuals.CardInnerOutline),
            Background = HyperTool.WinUI.Views.LifecycleVisuals.CreateCardInnerBrush(),
            Padding = new Thickness(20, 18, 20, 16)
        };

        _overlayTitle.Text = "HyperTool Guest";
        _overlayTitle.FontSize = 30;
        _overlayTitle.FontWeight = Microsoft.UI.Text.FontWeights.SemiBold;
        _overlayTitle.HorizontalAlignment = HorizontalAlignment.Center;
        _overlayTitle.Foreground = new SolidColorBrush(HyperTool.WinUI.Views.LifecycleVisuals.TextPrimary);

        _overlayText.FontSize = 13;
        _overlayText.HorizontalAlignment = HorizontalAlignment.Center;
        _overlayText.TextAlignment = TextAlignment.Center;
        _overlayText.Opacity = 0.94;
        _overlayText.Margin = new Thickness(0, 6, 0, 2);

        _overlayProgressBar.Height = 10;
        _overlayProgressBar.Minimum = 0;
        _overlayProgressBar.Maximum = 100;
        _overlayProgressBar.Value = 8;
        _overlayProgressBar.CornerRadius = new CornerRadius(5);
        _overlayProgressBar.Foreground = HyperTool.WinUI.Views.LifecycleVisuals.CreateProgressBrush();
        _overlayProgressBar.Background = new SolidColorBrush(HyperTool.WinUI.Views.LifecycleVisuals.ProgressTrack);
        _overlayProgressBar.Margin = new Thickness(8, 6, 8, 0);

        var stack = new StackPanel
        {
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center,
            Spacing = 12,
            Children =
            {
                new Grid
                {
                    Width = 122,
                    Height = 122,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Children =
                    {
                        new Ellipse
                        {
                            Width = 122,
                            Height = 122,
                            Fill = new RadialGradientBrush
                            {
                                Center = new Windows.Foundation.Point(0.5, 0.5),
                                GradientOrigin = new Windows.Foundation.Point(0.5, 0.5),
                                RadiusX = 0.5,
                                RadiusY = 0.5,
                                GradientStops =
                                {
                                    new GradientStop { Color = Color.FromArgb(0x22, 0x72, 0xC4, 0xFF), Offset = 0.0 },
                                    new GradientStop { Color = Color.FromArgb(0x00, 0x72, 0xC4, 0xFF), Offset = 1.0 }
                                }
                            },
                            Opacity = 0.30
                        },
                        new Ellipse
                        {
                            Width = 108,
                            Height = 108,
                            StrokeThickness = 1.1,
                            Stroke = new SolidColorBrush(Color.FromArgb(0x78, 0x88, 0xD1, 0xFF)),
                            Opacity = 0.36
                        },
                        new Border
                        {
                            Width = 98,
                            Height = 98,
                            CornerRadius = new CornerRadius(49),
                            Background = new SolidColorBrush(Color.FromArgb(0x64, 0x66, 0xC3, 0xFF)),
                            Opacity = 0.28
                        },
                        new Image
                        {
                            Source = new Microsoft.UI.Xaml.Media.Imaging.BitmapImage(new Uri("ms-appx:///Assets/HyperTool.Guest.Icon.Transparent.png")),
                            Width = 68,
                            Height = 68,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center
                        }
                    }
                },
                _overlayTitle,
                _overlayText,
                _overlayProgressBar
            }
        };

        innerFrame.Child = stack;
        _overlayCard.Child = innerFrame;

        overlayRoot.Children.Add(_overlayCard);

        try
        {
            _overlayAmbientStoryboard.Begin();
        }
        catch
        {
        }

        return overlayRoot;
    }

    private (UIElement Root, TextBlock StatusText) BuildReloadOverlayContent()
    {
        var overlayRoot = new Grid
        {
            Background = HyperTool.WinUI.Views.LifecycleVisuals.CreateRootBackgroundBrush()
        };

        overlayRoot.Children.Add(new TextBlock
        {
            Text = "Copyright: koerby",
            HorizontalAlignment = HorizontalAlignment.Left,
            VerticalAlignment = VerticalAlignment.Bottom,
            Margin = new Thickness(20, 0, 0, 14),
            FontSize = 12,
            Opacity = 0.72,
            Foreground = new SolidColorBrush(Windows.UI.Color.FromArgb(0xC8, 0x9B, 0xB7, 0xD7))
        });

        overlayRoot.Children.Add(new TextBlock
        {
            Text = ResolveGuestVersionText(),
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Bottom,
            Margin = new Thickness(0, 0, 20, 14),
            FontSize = 12,
            Opacity = 0.72,
            Foreground = new SolidColorBrush(Windows.UI.Color.FromArgb(0xC8, 0x9B, 0xB7, 0xD7))
        });

        overlayRoot.Children.Add(new Rectangle
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment = VerticalAlignment.Stretch,
            Fill = HyperTool.WinUI.Views.LifecycleVisuals.CreateVignetteBrush(0x78)
        });

        var statusText = new TextBlock
        {
            FontSize = 13,
            HorizontalAlignment = HorizontalAlignment.Center,
            TextAlignment = TextAlignment.Center,
            Foreground = new SolidColorBrush(HyperTool.WinUI.Views.LifecycleVisuals.TextSecondary),
            Opacity = 0.95,
            Margin = new Thickness(0)
        };

        var card = new Border
        {
            Width = 420,
            Height = double.NaN,
            Padding = new Thickness(24, 22, 24, 18),
            CornerRadius = new CornerRadius(20),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Background = HyperTool.WinUI.Views.LifecycleVisuals.CreateCardSurfaceBrush(),
            BorderBrush = new SolidColorBrush(HyperTool.WinUI.Views.LifecycleVisuals.CardBorder),
            BorderThickness = new Thickness(1),
            Shadow = new ThemeShadow()
        };

        card.Child = new StackPanel
        {
            Spacing = 10,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Children =
            {
                new Image
                {
                    Source = new Microsoft.UI.Xaml.Media.Imaging.BitmapImage(new Uri("ms-appx:///Assets/HyperTool.Guest.Icon.Transparent.png")),
                    Width = 54,
                    Height = 54,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    Opacity = 0.95
                },
                new TextBlock
                {
                    Text = "HyperTool Guest",
                    HorizontalAlignment = HorizontalAlignment.Center,
                    FontSize = 24,
                    FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                    Foreground = new SolidColorBrush(HyperTool.WinUI.Views.LifecycleVisuals.TextPrimary)
                },
                statusText,
                new ProgressRing
                {
                    Width = 30,
                    Height = 30,
                    IsActive = true,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Foreground = new SolidColorBrush(Color.FromArgb(0xFF, 0x8D, 0xCF, 0xFF)),
                    Margin = new Thickness(0, 2, 0, 0)
                }
            }
        };

        overlayRoot.Children.Add(card);
        return (overlayRoot, statusText);
    }

    private void BuildOverlayAmbientBands(Grid root)
    {
        var canvas = new Canvas { IsHitTestVisible = false };

        var band1 = CreateOverlayMovingBand(660, 44, 0.06, -14, -800, 1180, 7000, 0);
        var band2 = CreateOverlayMovingBand(520, 32, 0.04, -10, -760, 1140, 7600, 1200);

        Canvas.SetTop(band1, 152);
        Canvas.SetTop(band2, 476);

        canvas.Children.Add(band1);
        canvas.Children.Add(band2);

        root.Children.Add(canvas);
    }

    private void BuildOverlayNetworkLayer(Grid root)
    {
        var canvas = new Canvas { IsHitTestVisible = false };

        var nodes = new[]
        {
            (X: 236.0, Y: 260.0, Size: 9.8, Highlight: true, Label: "Host"),
            (X: 504.0, Y: 316.0, Size: 8.6, Highlight: true, Label: "Hyper-V"),
            (X: 642.0, Y: 294.0, Size: 8.4, Highlight: false, Label: (string?)null),
            (X: 780.0, Y: 274.0, Size: 9.4, Highlight: true, Label: "VM"),
            (X: 914.0, Y: 312.0, Size: 8.2, Highlight: false, Label: "Client"),
            (X: 1006.0, Y: 346.0, Size: 8.0, Highlight: false, Label: "Target"),
            (X: 422.0, Y: 388.0, Size: 7.2, Highlight: false, Label: (string?)null),
            (X: 838.0, Y: 392.0, Size: 7.2, Highlight: false, Label: (string?)null)
        };

        foreach (var node in nodes)
        {
            var circle = new Ellipse
            {
                Width = node.Size,
                Height = node.Size,
                Fill = new SolidColorBrush(node.Highlight
                    ? Color.FromArgb(0xD6, 0x67, 0xBF, 0xF8)
                    : HyperTool.WinUI.Views.LifecycleVisuals.NodeColor),
                Stroke = new SolidColorBrush(HyperTool.WinUI.Views.LifecycleVisuals.NodeStroke),
                StrokeThickness = node.Highlight ? 1.0 : 0.9,
                Opacity = node.Highlight ? 0.66 : 0.52
            };
            Canvas.SetLeft(circle, node.X - (node.Size / 2));
            Canvas.SetTop(circle, node.Y - (node.Size / 2));
            canvas.Children.Add(circle);

            if (!string.IsNullOrWhiteSpace(node.Label))
            {
                var label = new TextBlock
                {
                    Text = node.Label,
                    FontSize = 11,
                    FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                    Opacity = 0.58,
                    Foreground = new SolidColorBrush(Color.FromArgb(0xC2, 0xA6, 0xC4, 0xE3))
                };
                Canvas.SetLeft(label, node.X - 24);
                Canvas.SetTop(label, node.Y + 13);
                canvas.Children.Add(label);
            }
        }

        var links = new (int From, int To)[]
        {
            (0,4), (4,1), (1,5), (5,2), (2,6), (6,3), (1,7), (7,3)
        };

        for (var i = 0; i < links.Length; i++)
        {
            var (from, to) = links[i];
            if (from < 0 || from >= nodes.Length || to < 0 || to >= nodes.Length)
            {
                continue;
            }

            var a = nodes[from];
            var b = nodes[to];

            var line = new Line
            {
                X1 = a.X,
                Y1 = a.Y,
                X2 = b.X,
                Y2 = b.Y,
                Stroke = new SolidColorBrush(HyperTool.WinUI.Views.LifecycleVisuals.LineColor),
                StrokeThickness = 0.8,
                Opacity = 0.0
            };
            canvas.Children.Add(line);

            var lineAppear = new DoubleAnimation
            {
                From = 0.0,
                To = 0.34,
                Duration = TimeSpan.FromMilliseconds(340),
                BeginTime = TimeSpan.FromMilliseconds(170 + (i * 140)),
                FillBehavior = FillBehavior.HoldEnd,
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(lineAppear, line);
            Storyboard.SetTargetProperty(lineAppear, "Opacity");
            _overlayAmbientStoryboard.Children.Add(lineAppear);

            var pulse = new Ellipse
            {
                Width = 3,
                Height = 3,
                Fill = new SolidColorBrush(HyperTool.WinUI.Views.LifecycleVisuals.PulseColor),
                Opacity = 0.0
            };
            Canvas.SetLeft(pulse, a.X - 1.5);
            Canvas.SetTop(pulse, a.Y - 1.5);
            canvas.Children.Add(pulse);

            var pulseFade = new DoubleAnimation
            {
                From = 0.0,
                To = 0.28,
                Duration = TimeSpan.FromMilliseconds(220),
                BeginTime = TimeSpan.FromMilliseconds(560 + (i * 120)),
                FillBehavior = FillBehavior.HoldEnd,
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(pulseFade, pulse);
            Storyboard.SetTargetProperty(pulseFade, "Opacity");
            _overlayAmbientStoryboard.Children.Add(pulseFade);

            var pulseX = new DoubleAnimation
            {
                From = a.X - 1.5,
                To = b.X - 1.5,
                Duration = TimeSpan.FromMilliseconds(1360 + (i * 90)),
                BeginTime = TimeSpan.FromMilliseconds(760 + (i * 100)),
                RepeatBehavior = RepeatBehavior.Forever,
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(pulseX, pulse);
            Storyboard.SetTargetProperty(pulseX, "(Canvas.Left)");
            _overlayAmbientStoryboard.Children.Add(pulseX);

            var pulseY = new DoubleAnimation
            {
                From = a.Y - 1.5,
                To = b.Y - 1.5,
                Duration = TimeSpan.FromMilliseconds(1360 + (i * 90)),
                BeginTime = TimeSpan.FromMilliseconds(760 + (i * 100)),
                RepeatBehavior = RepeatBehavior.Forever,
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(pulseY, pulse);
            Storyboard.SetTargetProperty(pulseY, "(Canvas.Top)");
            _overlayAmbientStoryboard.Children.Add(pulseY);
        }

        root.Children.Add(canvas);
    }

    private Rectangle CreateOverlayMovingBand(double width, double height, double opacity, double rotation, double fromX, double toX, int durationMs, int beginMs)
    {
        var band = new Rectangle
        {
            Width = width,
            Height = height,
            RadiusX = height / 2,
            RadiusY = height / 2,
            Opacity = opacity,
            Fill = new LinearGradientBrush
            {
                StartPoint = new Windows.Foundation.Point(0, 0.5),
                EndPoint = new Windows.Foundation.Point(1, 0.5),
                GradientStops =
                {
                    new GradientStop { Color = Color.FromArgb(0x00, 0x63, 0xC1, 0xFF), Offset = 0 },
                    new GradientStop { Color = Color.FromArgb(0xFF, 0x63, 0xC1, 0xFF), Offset = 0.5 },
                    new GradientStop { Color = Color.FromArgb(0x00, 0x63, 0xC1, 0xFF), Offset = 1 }
                }
            },
            RenderTransform = new CompositeTransform { Rotation = rotation, TranslateX = fromX }
        };

        var move = new DoubleAnimation
        {
            From = fromX,
            To = toX,
            Duration = TimeSpan.FromMilliseconds(durationMs),
            BeginTime = TimeSpan.FromMilliseconds(beginMs),
            RepeatBehavior = RepeatBehavior.Forever,
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(move, band);
        Storyboard.SetTargetProperty(move, "(UIElement.RenderTransform).(CompositeTransform.TranslateX)");
        _overlayAmbientStoryboard.Children.Add(move);

        return band;
    }

    private UIElement BuildUsbPage()
    {
        var root = new Grid { RowSpacing = 10 };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var actionsCard = new Border
        {
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush,
            Padding = new Thickness(10)
        };

        var actionsStack = new StackPanel { Spacing = 5 };

        var headerRow = new Grid { ColumnSpacing = 10 };
        headerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var headerLeft = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 14,
            VerticalAlignment = VerticalAlignment.Center
        };

        _usbFeatureEnabledToggleSwitch.Toggled += OnUsbFeatureToggleChanged;
        _usbFeatureEnabledToggleSwitch.IsOn = _config.Usb?.Enabled != false;
        _usbFeatureEnabledToggleSwitch.MinWidth = 54;
        _usbFeatureEnabledToggleSwitch.VerticalAlignment = VerticalAlignment.Center;
        headerLeft.Children.Add(_usbFeatureEnabledToggleSwitch);

        headerLeft.Children.Add(new TextBlock
        {
            Text = "USB-Share (USB/IP-WIN2 Client im Hintergrund)",
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            VerticalAlignment = VerticalAlignment.Center
        });
        headerRow.Children.Add(headerLeft);

        _usbPageFeatureStatusChip.Child = _usbPageFeatureStatusChipText;
        Grid.SetColumn(_usbPageFeatureStatusChip, 1);
        headerRow.Children.Add(_usbPageFeatureStatusChip);

        actionsStack.Children.Add(headerRow);

        _usbRuntimeInstallButton = CreateIconButton("⬇", "Installation usbip-win2", onClick: async (_, _) => await InstallGuestUsbRuntimeAsync());
        _usbRuntimeRestartButton = CreateIconButton(ToolRestartIcon, ToolRestartLabel, onClick: async (_, _) => await RestartGuestToolAsync());
        _usbRuntimeInstallButton.Visibility = Visibility.Collapsed;
        _usbRuntimeRestartButton.Visibility = Visibility.Collapsed;
        var usbRuntimeActionsRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        usbRuntimeActionsRow.Children.Add(_usbRuntimeInstallButton);
        usbRuntimeActionsRow.Children.Add(_usbRuntimeRestartButton);
        actionsStack.Children.Add(usbRuntimeActionsRow);

        var actionRow = new Grid { ColumnSpacing = 8, VerticalAlignment = VerticalAlignment.Center };
        actionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        _usbRefreshButton = CreateIconButton("⟳", "Refresh", onClick: async (_, _) => await RefreshUsbAsync());
        _usbConnectButton = CreateIconButton("🔌", "Connect", onClick: async (_, _) => await ConnectUsbAsync());
        _usbDisconnectButton = CreateIconButton("⏏", "Disconnect", onClick: async (_, _) => await DisconnectUsbAsync());

        _usbRefreshButton.IsEnabled = IsUsbRefreshAvailable();
        _usbConnectButton.IsEnabled = IsUsbFeatureUsable();
        _usbDisconnectButton.IsEnabled = IsUsbFeatureUsable();

        Grid.SetColumn(_usbRefreshButton, 0);
        Grid.SetColumn(_usbConnectButton, 1);
        Grid.SetColumn(_usbDisconnectButton, 2);
        Grid.SetColumn(_usbAutoConnectCheckBox, 3);
        Grid.SetColumn(_usbDisconnectOnExitCheckBox, 4);

        actionRow.Children.Add(_usbRefreshButton);
        actionRow.Children.Add(_usbConnectButton);
        actionRow.Children.Add(_usbDisconnectButton);

        _usbAutoConnectCheckBox.IsEnabled = IsUsbFeatureUsable();
        _usbAutoConnectCheckBox.Margin = new Thickness(6, 0, 0, 0);
        _usbAutoConnectCheckBox.VerticalAlignment = VerticalAlignment.Center;
        _usbAutoConnectCheckBox.HorizontalAlignment = HorizontalAlignment.Left;
        _usbAutoConnectCheckBox.Checked += async (_, _) => await SetSelectedUsbDeviceAutoConnectAsync(true);
        _usbAutoConnectCheckBox.Unchecked += async (_, _) => await SetSelectedUsbDeviceAutoConnectAsync(false);
        actionRow.Children.Add(_usbAutoConnectCheckBox);

        _usbDisconnectOnExitCheckBox.IsEnabled = IsUsbFeatureUsable();
        _usbDisconnectOnExitCheckBox.Margin = new Thickness(6, 0, 0, 0);
        _usbDisconnectOnExitCheckBox.VerticalAlignment = VerticalAlignment.Center;
        _usbDisconnectOnExitCheckBox.HorizontalAlignment = HorizontalAlignment.Left;
        _usbDisconnectOnExitCheckBox.IsChecked = _config.Usb?.DisconnectOnExit != false;
        _usbDisconnectOnExitCheckBox.Checked += async (_, _) => await SetUsbDisconnectOnExitAsync(true);
        _usbDisconnectOnExitCheckBox.Unchecked += async (_, _) => await SetUsbDisconnectOnExitAsync(false);
        actionRow.Children.Add(_usbDisconnectOnExitCheckBox);

        actionsStack.Children.Add(actionRow);

        if (!_isUsbClientAvailable)
        {
            actionsStack.Children.Add(new TextBlock
            {
                Text = "USB/IP-Client (usbip-win2) ist nicht installiert. USB-Funktionen sind deaktiviert. Quelle: github.com/vadimgrn/usbip-win2",
                TextWrapping = TextWrapping.Wrap,
                Opacity = 0.88,
                Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
            });
        }

        actionsStack.Children.Add(new TextBlock
        {
            Text = "Hinweis: Für USB-Share muss am Host 'RemoteFX USB Device Redirection' in den Richtlinien deaktiviert sein.",
            TextWrapping = TextWrapping.Wrap,
            Opacity = 0.88,
            Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
        });

        UpdateUsbRuntimeStatusUi();

        actionsCard.Child = actionsStack;
        root.Children.Add(actionsCard);

        var listBorder = new Border
        {
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            Padding = new Thickness(8),
        };

        _usbDisabledOverlayText.Foreground = Application.Current.Resources["TextMutedBrush"] as Brush;
        var listBody = new Grid();
        listBody.Children.Add(_usbListView);
        listBody.Children.Add(_usbDisabledOverlayText);

        var usbStatusOverlay = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Bottom,
            Margin = new Thickness(8, 8, 10, 6),
            Opacity = 0.76,
            IsHitTestVisible = false
        };
        usbStatusOverlay.Children.Add(_usbRuntimeStatusDot);
        usbStatusOverlay.Children.Add(_usbRuntimeStatusText);
        listBody.Children.Add(usbStatusOverlay);

        listBorder.Child = listBody;

        _usbListView.SelectionChanged += (_, _) => UpdateAutoConnectToggleFromSelection();

        Grid.SetRow(listBorder, 1);
        root.Children.Add(listBorder);

        return root;
    }

    private UIElement BuildSharedFoldersPage()
    {
        var root = new Grid { RowSpacing = 10 };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var editorCard = new Border
        {
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush,
            Padding = new Thickness(10)
        };

        var editorStack = new StackPanel { Spacing = 6 };

        var headerRow = new Grid { ColumnSpacing = 10 };
        headerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var headerLeft = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 14,
            VerticalAlignment = VerticalAlignment.Center
        };

        _sharedFolderFeatureEnabledToggleSwitch.Toggled += OnSharedFolderFeatureToggleChanged;
        _sharedFolderFeatureEnabledToggleSwitch.IsOn = _config.SharedFolders?.Enabled == true;
        _sharedFolderFeatureEnabledToggleSwitch.MinWidth = 54;
        _sharedFolderFeatureEnabledToggleSwitch.VerticalAlignment = VerticalAlignment.Center;
        headerLeft.Children.Add(_sharedFolderFeatureEnabledToggleSwitch);

        var titleText = new TextBlock
        {
            Text = "Shared Folder (Netzlaufwerk über WinFsp im Hintergrund)",
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            VerticalAlignment = VerticalAlignment.Center
        };
        headerLeft.Children.Add(titleText);
        headerRow.Children.Add(headerLeft);

        _sharedFolderPageFeatureStatusChip.Child = _sharedFolderPageFeatureStatusChipText;
        Grid.SetColumn(_sharedFolderPageFeatureStatusChip, 1);
        headerRow.Children.Add(_sharedFolderPageFeatureStatusChip);

        editorStack.Children.Add(headerRow);

        _winFspRuntimeInstallButton = CreateIconButton("⬇", "WinFsp nachinstallieren", onClick: async (_, _) => await InstallGuestWinFspRuntimeAsync());
        _winFspRuntimeRestartButton = CreateIconButton(ToolRestartIcon, ToolRestartLabel, onClick: async (_, _) => await RestartGuestToolAsync());
        var winFspRuntimeActionsRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        winFspRuntimeActionsRow.Children.Add(_winFspRuntimeInstallButton);
        winFspRuntimeActionsRow.Children.Add(_winFspRuntimeRestartButton);
        editorStack.Children.Add(winFspRuntimeActionsRow);

        var settingsStack = new StackPanel { Spacing = 6 };

        if (_sharedFolderBaseDriveCombo.Items.Count == 0)
        {
            for (var drive = 'D'; drive <= 'Z'; drive++)
            {
                _sharedFolderBaseDriveCombo.Items.Add(drive.ToString());
            }
        }

        _sharedFolderBaseDriveCombo.SelectionChanged -= OnSharedFolderBaseDriveSelectionChanged;
        _sharedFolderBaseDriveCombo.SelectionChanged += OnSharedFolderBaseDriveSelectionChanged;

        var actionsRow = new Grid
        {
            ColumnSpacing = 8,
            VerticalAlignment = VerticalAlignment.Center
        };
        actionsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var refreshHostListButton = CreateIconButton("⟳", "Refresh", onClick: async (_, _) => await SyncSharedFoldersFromHostAsync(forceHostFeatureRefresh: true));
        Grid.SetColumn(refreshHostListButton, 0);
        actionsRow.Children.Add(refreshHostListButton);

        var selfTestButton = CreateIconButton("🧪", "Selftest", onClick: async (_, _) => await RunSharedFolderSelfTestAsync());
        Grid.SetColumn(selfTestButton, 1);
        actionsRow.Children.Add(selfTestButton);

        var driveLabel = new TextBlock
        {
            Text = "Netzlaufwerk",
            VerticalAlignment = VerticalAlignment.Center,
            Opacity = 0.9
        };
        Grid.SetColumn(driveLabel, 2);
        actionsRow.Children.Add(driveLabel);

        Grid.SetColumn(_sharedFolderBaseDriveCombo, 3);
        actionsRow.Children.Add(_sharedFolderBaseDriveCombo);

        _sharedFolderSaveDriveButton = CreateIconButton("💾", "Speichern", onClick: async (_, _) => await SaveSharedFolderBaseDriveAsync());
        Grid.SetColumn(_sharedFolderSaveDriveButton, 4);
        actionsRow.Children.Add(_sharedFolderSaveDriveButton);

        settingsStack.Children.Add(actionsRow);

        var hasWinFspRuntime = IsWinFspRuntimeInstalled();
        if (_winFspRuntimeInstallButton is not null)
        {
            _winFspRuntimeInstallButton.Visibility = hasWinFspRuntime ? Visibility.Collapsed : Visibility.Visible;
        }

        if (!hasWinFspRuntime)
        {
            settingsStack.Children.Add(new TextBlock
            {
                Text = "WinFsp ist nicht installiert. hypertool-file Shared-Folder sind ohne WinFsp nicht nutzbar.",
                Opacity = 0.88,
                TextWrapping = TextWrapping.Wrap,
                Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
            });
        }

        settingsStack.Children.Add(new TextBlock
        {
            Text = "Shared-Folder läuft rein über Hyper-V Socket / HyperTool File Service (kein IP-Fallback).",
            Opacity = 0.72,
            TextWrapping = TextWrapping.Wrap
        });

        _sharedFolderSettingsPanel.Child = settingsStack;
        editorStack.Children.Add(_sharedFolderSettingsPanel);

        editorStack.Children.Add(_sharedFolderStatusText);
        editorCard.Child = editorStack;
        root.Children.Add(editorCard);

        UpdateWinFspRuntimeStatusUi();

        var listCard = new Border
        {
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            Padding = new Thickness(8)
        };

        _sharedFolderMappingsListView.ItemsSource = _sharedFolderMappings;
        _sharedFolderMappingsListView.ItemTemplate = CreateSharedFolderMappingTemplate();
        _sharedFolderMappingsListView.SelectionMode = ListViewSelectionMode.None;
        _sharedFolderMappingsListView.IsItemClickEnabled = false;

        _sharedFolderDisabledOverlayText.Foreground = Application.Current.Resources["TextMutedBrush"] as Brush;
        var listBody = new Grid();
        listBody.Children.Add(_sharedFolderMappingsListView);
        listBody.Children.Add(_sharedFolderDisabledOverlayText);

        var sharedFolderStatusOverlay = new StackPanel
        {
            Spacing = 4,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Bottom,
            Margin = new Thickness(8, 8, 10, 6),
            Opacity = 0.74,
            IsHitTestVisible = false
        };
        var winFspOverlayRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, HorizontalAlignment = HorizontalAlignment.Right };
        winFspOverlayRow.Children.Add(_winFspRuntimeStatusDot);
        winFspOverlayRow.Children.Add(_winFspRuntimeStatusText);
        sharedFolderStatusOverlay.Children.Add(winFspOverlayRow);

        _sharedFolderReconnectStatusText.HorizontalAlignment = HorizontalAlignment.Right;
        _sharedFolderReconnectStatusText.VerticalAlignment = VerticalAlignment.Center;
        _sharedFolderReconnectStatusText.Margin = new Thickness(0);
        _sharedFolderReconnectStatusText.Opacity = 1.0;
        _sharedFolderReconnectStatusText.IsHitTestVisible = false;
        sharedFolderStatusOverlay.Children.Add(_sharedFolderReconnectStatusText);
        listBody.Children.Add(sharedFolderStatusOverlay);

        listCard.Child = listBody;
        Grid.SetRow(listCard, 1);
        root.Children.Add(listCard);

        RefreshSharedFolderMappingsFromConfig();
        UpdateSharedFolderFeatureUi();
        UpdateHostDiscoveryPresentation();
        _ = SyncSharedFoldersFromHostAsync(forceHostFeatureRefresh: false);
        _ = RefreshSharedFolderMountStatesSafeAsync();
        if (!IsSharedFolderFeatureEnabled())
        {
            _ = ApplySharedFolderSelectionAsync(autoTriggered: true);
        }

        return root;
    }

    private static DataTemplate CreateSharedFolderMappingTemplate()
    {
        const string templateXaml = """
<DataTemplate xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'>
    <Grid ColumnSpacing='8' Margin='4,2,4,2'>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width='36'/>
            <ColumnDefinition Width='32'/>
            <ColumnDefinition Width='*'/>
            <ColumnDefinition Width='90'/>
        </Grid.ColumnDefinitions>
        <CheckBox IsChecked='{Binding Enabled, Mode=TwoWay}' VerticalAlignment='Center' HorizontalAlignment='Left'/>
        <TextBlock Grid.Column='1' Text='{Binding MountStateDot}' VerticalAlignment='Center' HorizontalAlignment='Center' FontSize='12'/>
        <StackPanel Grid.Column='2' Spacing='1' VerticalAlignment='Center'>
            <TextBlock Text='{Binding Label}' FontWeight='SemiBold' Opacity='0.95' TextTrimming='CharacterEllipsis' TextWrapping='NoWrap'/>
            <TextBlock Text='{Binding SharePath}' Opacity='0.72' FontSize='11' TextTrimming='CharacterEllipsis' TextWrapping='NoWrap'/>
        </StackPanel>
        <TextBlock Grid.Column='3' Text='{Binding MountStateText}' Opacity='0.82' TextTrimming='CharacterEllipsis' TextWrapping='NoWrap'/>
    </Grid>
</DataTemplate>
""";

        return (DataTemplate)XamlReader.Load(templateXaml);
    }

    private void RefreshSharedFolderMappingsFromConfig()
    {
        _suppressSharedFolderAutoApply = true;
        try
        {
            foreach (var existing in _sharedFolderMappings)
            {
                existing.PropertyChanged -= OnSharedFolderMappingPropertyChanged;
            }

        _config.SharedFolders ??= new GuestSharedFolderSettings();
        _config.SharedFolders.Mappings ??= [];
        _suppressSharedFolderFeatureToggle = true;
        try
        {
            _sharedFolderFeatureEnabledToggleSwitch.IsOn = _config.SharedFolders.Enabled;
        }
        finally
        {
            _suppressSharedFolderFeatureToggle = false;
        }

        _suppressSharedFolderBaseDriveChange = true;
        try
        {
            _savedSharedFolderBaseDriveLetter = GuestConfigService.NormalizeDriveLetter(_config.SharedFolders.BaseDriveLetter);
            _pendingSharedFolderBaseDriveLetter = _savedSharedFolderBaseDriveLetter;
            _sharedFolderBaseDriveCombo.SelectedItem = _pendingSharedFolderBaseDriveLetter;
        }
        finally
        {
            _suppressSharedFolderBaseDriveChange = false;
        }

        var normalizedMappings = new List<GuestSharedFolderMapping>();
        foreach (var mapping in _config.SharedFolders.Mappings)
        {
            if (mapping is null)
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(mapping.Id))
            {
                mapping.Id = Guid.NewGuid().ToString("N");
            }

            var normalizedDriveLetter = GuestConfigService.NormalizeDriveLetter(mapping.DriveLetter);
            var normalizedSharePath = (mapping.SharePath ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalizedSharePath))
            {
                continue;
            }

            normalizedMappings.Add(new GuestSharedFolderMapping
            {
                Id = mapping.Id,
                Label = ResolveMappingLabel(mapping.Label, normalizedSharePath),
                SharePath = normalizedSharePath,
                DriveLetter = normalizedDriveLetter,
                Persistent = true,
                Enabled = mapping.Enabled,
                MountStateDot = mapping.Enabled ? "🔴" : "⚪",
                MountStateText = mapping.Enabled ? "getrennt" : "deaktiviert"
            });
        }

        _sharedFolderMappings.Clear();
        foreach (var mapping in normalizedMappings)
        {
            _sharedFolderMappings.Add(mapping);
            mapping.PropertyChanged += OnSharedFolderMappingPropertyChanged;
        }
        }
        finally
        {
            _suppressSharedFolderAutoApply = false;
        }
    }

    private void OnSharedFolderMappingPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (_suppressSharedFolderAutoApply)
        {
            return;
        }

        if (sender is not GuestSharedFolderMapping)
        {
            return;
        }

        if (!string.Equals(e.PropertyName, nameof(GuestSharedFolderMapping.Enabled), StringComparison.Ordinal))
        {
            return;
        }

        ScheduleSharedFolderAutoApply();
    }

    private void OnSharedFolderFeatureToggleChanged(object sender, RoutedEventArgs e)
    {
        if (_suppressSharedFolderFeatureToggle)
        {
            return;
        }

        if (!IsWinFspRuntimeInstalled())
        {
            _suppressSharedFolderFeatureToggle = true;
            try
            {
                _sharedFolderFeatureEnabledToggleSwitch.IsOn = false;
            }
            finally
            {
                _suppressSharedFolderFeatureToggle = false;
            }

            _config.SharedFolders ??= new GuestSharedFolderSettings();
            _config.SharedFolders.Enabled = false;
            _ = SaveConfigQuietlyAsync();
            UpdateSharedFolderFeatureUi();
            _sharedFolderStatusText.Text = "Shared-Folder Funktion deaktiviert: WinFsp Runtime ist nicht installiert.";
            return;
        }

        var enabled = _sharedFolderFeatureEnabledToggleSwitch.IsOn;
        _config.SharedFolders ??= new GuestSharedFolderSettings();
        _config.SharedFolders.Enabled = enabled;
        UpdateSharedFolderFeatureUi();
        ScheduleSharedFolderAutoApply();
    }

    private async void OnUsbFeatureToggleChanged(object sender, RoutedEventArgs e)
    {
        if (_suppressUsbFeatureToggle)
        {
            return;
        }

        if (!_isUsbClientAvailable)
        {
            _suppressUsbFeatureToggle = true;
            try
            {
                _usbFeatureEnabledToggleSwitch.IsOn = false;
            }
            finally
            {
                _suppressUsbFeatureToggle = false;
            }

            _config.Usb ??= new GuestUsbSettings();
            _config.Usb.Enabled = false;
            await SaveConfigQuietlyAsync();
            UpdateUsbRuntimeStatusUi();
            AppendNotification("[Warn] USB-Share bleibt deaktiviert, bis usbip-win2 installiert ist.");
            return;
        }

        var enabled = _usbFeatureEnabledToggleSwitch.IsOn;
        _config.Usb ??= new GuestUsbSettings();
        _config.Usb.Enabled = enabled;
        UpdateUsbRuntimeStatusUi();

        await _saveConfigAsync(_config);

        if (!enabled)
        {
            UpdateUsbDevices(Array.Empty<UsbIpDeviceInfo>());
            AppendNotification("[Info] USB Funktion lokal deaktiviert.");
            return;
        }

        AppendNotification("[Info] USB Funktion lokal aktiviert. Aktualisiere Geräteansicht …");
        await RefreshUsbAsync();
    }

    private void OnSharedFolderBaseDriveSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressSharedFolderBaseDriveChange)
        {
            return;
        }

        if (_sharedFolderBaseDriveCombo.SelectedItem is not string selectedLetter || string.IsNullOrWhiteSpace(selectedLetter))
        {
            return;
        }

        _pendingSharedFolderBaseDriveLetter = GuestConfigService.NormalizeDriveLetter(selectedLetter);
        MarkSharedFolderSelectionPending();
    }

    private void MarkSharedFolderSelectionPending()
    {
        if (string.Equals(_pendingSharedFolderBaseDriveLetter, _savedSharedFolderBaseDriveLetter, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        _sharedFolderStatusText.Text = $"Laufwerkswechsel vorgemerkt: {_savedSharedFolderBaseDriveLetter}: → {_pendingSharedFolderBaseDriveLetter}:. Mit 'Laufwerk speichern' übernehmen.";
    }

    private void UpdateSharedFolderFeatureUi()
    {
        var hostEnabled = IsSharedFolderFeatureEnabledByHost();
        var localEnabled = IsSharedFolderFeatureLocallyEnabled();
        var winFspAvailable = IsWinFspRuntimeInstalled();

        if (!winFspAvailable && localEnabled)
        {
            _config.SharedFolders ??= new GuestSharedFolderSettings();
            _config.SharedFolders.Enabled = false;
            localEnabled = false;
            _ = SaveConfigQuietlyAsync();
        }

        var effectiveEnabled = hostEnabled && localEnabled && winFspAvailable;
        var interactive = effectiveEnabled && !_sharedFolderHostFeatureReactivationPending;

        _sharedFolderFeatureEnabledToggleSwitch.IsEnabled = hostEnabled && winFspAvailable;
        _sharedFolderSettingsPanel.IsHitTestVisible = true;
        _sharedFolderSettingsPanel.Opacity = 1.0;
        _sharedFolderMappingsListView.IsEnabled = interactive;
        _sharedFolderBaseDriveCombo.IsEnabled = interactive;
        if (_sharedFolderSaveDriveButton is not null)
        {
            _sharedFolderSaveDriveButton.IsEnabled = interactive;
        }
        _sharedFolderDisabledOverlayText.Visibility = interactive ? Visibility.Collapsed : Visibility.Visible;

        var sharedFolderChipPalette = ResolveHostFeatureChipPalette(effectiveEnabled);
        _sharedFolderHostFeatureStatusChip.Background = sharedFolderChipPalette.chipBackground;
        _sharedFolderHostFeatureStatusChip.BorderBrush = sharedFolderChipPalette.chipBorder;
        _sharedFolderHostFeatureStatusChipText.Text = effectiveEnabled ? "Share Aktiv" : "Share Inaktiv";
        _sharedFolderHostFeatureStatusChipText.Foreground = sharedFolderChipPalette.textForeground;

        _sharedFolderPageFeatureStatusChip.Background = sharedFolderChipPalette.chipBackground;
        _sharedFolderPageFeatureStatusChip.BorderBrush = sharedFolderChipPalette.chipBorder;
        _sharedFolderPageFeatureStatusChipText.Text = effectiveEnabled ? "Aktiv" : "Inaktiv";
        _sharedFolderPageFeatureStatusChipText.Foreground = sharedFolderChipPalette.textForeground;

        if (!hostEnabled)
        {
            _sharedFolderStatusText.Text = "Shared-Folder sind durch den Host deaktiviert.";
        }
        else if (!winFspAvailable)
        {
            _sharedFolderStatusText.Text = "Shared-Folder sind deaktiviert: WinFsp Runtime ist nicht installiert.";
        }
        else if (!localEnabled)
        {
            _sharedFolderStatusText.Text = "Shared-Folder Funktion lokal deaktiviert.";
        }
    }

    private bool IsSharedFolderFeatureEnabled()
    {
        return IsSharedFolderFeatureEnabledByHost() && IsSharedFolderFeatureLocallyEnabled() && IsWinFspRuntimeInstalled();
    }

    private bool IsSharedFolderFeatureLocallyEnabled()
    {
        return _config.SharedFolders?.Enabled ?? true;
    }

    private bool IsSharedFolderFeatureEnabledByHost()
    {
        return _config.SharedFolders?.HostFeatureEnabled != false;
    }

    private bool IsUsbFeatureEnabledByHost()
    {
        return _config.Usb?.HostFeatureEnabled != false;
    }

    private bool IsUsbFeatureLocallyEnabled()
    {
        return _config.Usb?.Enabled != false;
    }

    private bool IsUsbRefreshAvailable()
    {
        return _isUsbClientAvailable && IsUsbFeatureLocallyEnabled();
    }

    private bool IsUsbFeatureUsable()
    {
        return _isUsbClientAvailable && IsUsbFeatureEnabledByHost() && IsUsbFeatureLocallyEnabled();
    }

    private string GetSharedFolderBaseDriveLetter()
    {
        _config.SharedFolders ??= new GuestSharedFolderSettings();
        _savedSharedFolderBaseDriveLetter = GuestConfigService.NormalizeDriveLetter(_savedSharedFolderBaseDriveLetter);
        _config.SharedFolders.BaseDriveLetter = _savedSharedFolderBaseDriveLetter;
        return _savedSharedFolderBaseDriveLetter;
    }

    private async Task SaveSharedFolderBaseDriveAsync()
    {
        if (string.Equals(_pendingSharedFolderBaseDriveLetter, _savedSharedFolderBaseDriveLetter, StringComparison.OrdinalIgnoreCase))
        {
            _sharedFolderStatusText.Text = $"Laufwerksbuchstabe unverändert ({_savedSharedFolderBaseDriveLetter}:).";
            return;
        }

        var previousBaseDriveLetter = _savedSharedFolderBaseDriveLetter;
        _savedSharedFolderBaseDriveLetter = _pendingSharedFolderBaseDriveLetter;
        _config.SharedFolders ??= new GuestSharedFolderSettings();
        _config.SharedFolders.BaseDriveLetter = _savedSharedFolderBaseDriveLetter;

        await SaveSharedFolderMappingsToConfigAsync();
        await _winFspMountService.UnmountAsync(previousBaseDriveLetter, CancellationToken.None);
        await ApplySharedFolderSelectionAsync(autoTriggered: false);
    }

    private void ScheduleSharedFolderAutoApply()
    {
        try
        {
            _sharedFolderAutoApplyCts?.Cancel();
        }
        catch
        {
        }

        _sharedFolderAutoApplyCts?.Dispose();
        _sharedFolderAutoApplyCts = new CancellationTokenSource();
        var token = _sharedFolderAutoApplyCts.Token;

        _ = RunSharedFolderAutoApplyAsync(token);
    }

    private async Task RunSharedFolderAutoApplyAsync(CancellationToken cancellationToken)
    {
        try
        {
            await Task.Delay(180, cancellationToken);
            if (cancellationToken.IsCancellationRequested)
            {
                return;
            }

            await ApplySharedFolderSelectionAsync(autoTriggered: true);
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            _sharedFolderLastError = ex.Message;
            _sharedFolderStatusText.Text = $"Automatische Shared-Folder-Anwendung fehlgeschlagen: {ex.Message}";
            GuestLogger.Warn("sharedfolders.autoapply.failed", ex.Message, new
            {
                exceptionType = ex.GetType().FullName
            });
        }
    }

    private async Task RefreshSharedFolderMountStatesSafeAsync()
    {
        try
        {
            await RefreshSharedFolderMountStatesAsync();
        }
        catch (Exception ex)
        {
            _sharedFolderLastError = ex.Message;
            GuestLogger.Warn("sharedfolders.status.refresh_failed", ex.Message, new
            {
                exceptionType = ex.GetType().FullName
            });
        }
    }

    private async Task RefreshSharedFolderMountStatesAsync()
    {
        var mappingsSnapshot = _sharedFolderMappings.ToList();
        var featureEnabled = IsSharedFolderFeatureEnabled();
        var baseDriveLetter = GetSharedFolderBaseDriveLetter();
        var isCatalogMounted = featureEnabled && _winFspMountService.IsMounted(baseDriveLetter);

        foreach (var mapping in mappingsSnapshot)
        {
            if (!featureEnabled)
            {
                mapping.MountStateDot = "⚪";
                mapping.MountStateText = "deaktiviert";
                continue;
            }

            if (!mapping.Enabled)
            {
                mapping.MountStateDot = "⚪";
                mapping.MountStateText = "deaktiviert";
                continue;
            }

            try
            {
                mapping.MountStateDot = isCatalogMounted ? "🟢" : "🔴";
                mapping.MountStateText = isCatalogMounted
                    ? $"verbunden ({baseDriveLetter}:)"
                    : $"getrennt ({baseDriveLetter}:)";
            }
            catch (Exception ex)
            {
                mapping.MountStateDot = "🔴";
                mapping.MountStateText = "getrennt";
                GuestLogger.Warn("sharedfolders.status.query_failed", ex.Message, new
                {
                    mapping.Id,
                    mapping.Label,
                    mapping.SharePath,
                    mapping.DriveLetter,
                    exceptionType = ex.GetType().FullName
                });
                _sharedFolderLastError = ex.Message;
            }
        }

        _sharedFolderMappingsListView.ItemsSource = null;
        _sharedFolderMappingsListView.ItemsSource = _sharedFolderMappings;
    }

    private async Task ApplySharedFolderSelectionAsync(bool autoTriggered = false)
    {
        var lockAcquired = autoTriggered
            ? await _sharedFolderUiOperationGate.WaitAsync(0)
            : await _sharedFolderUiOperationGate.WaitAsync(TimeSpan.FromSeconds(8));

        if (!lockAcquired)
        {
            _sharedFolderStatusText.Text = autoTriggered
                ? "Änderung übernommen, sobald laufende Shared-Folder Aktion fertig ist."
                : "Shared-Folder Aktion läuft noch. Bitte in ein paar Sekunden erneut versuchen.";
            return;
        }

        try
        {
            _suppressSharedFolderAutoApply = true;
            if (!ValidateSharedFolderMappings(out var validationError))
            {
                _sharedFolderStatusText.Text = validationError;
                return;
            }

            await SaveSharedFolderMappingsToConfigAsync();

            var prefix = autoTriggered ? "Automatisch angewendet" : "Auswahl angewendet";
            var baseDriveLetter = GetSharedFolderBaseDriveLetter();

            if (!IsSharedFolderFeatureEnabled())
            {
                var legacyDriveLetters = _sharedFolderMappings
                    .Select(item => GuestConfigService.NormalizeDriveLetter(item.DriveLetter))
                    .Append(baseDriveLetter)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                await _winFspMountService.UnmountManyAsync(legacyDriveLetters, CancellationToken.None);
                await RefreshSharedFolderMountStatesSafeAsync();
                _sharedFolderStatusText.Text = $"{prefix}: Shared-Folder Funktion deaktiviert · Laufwerk getrennt.";
                _sharedFolderLastError = "-";
                return;
            }

            var enabledMappings = _sharedFolderMappings
                .Where(item => item.Enabled && !string.IsNullOrWhiteSpace(item.SharePath))
                .Select(item => new GuestSharedFolderMapping
                {
                    Id = item.Id,
                    Label = item.Label,
                    SharePath = item.SharePath,
                    DriveLetter = baseDriveLetter,
                    Persistent = true,
                    Enabled = true
                })
                .ToList();

            if (enabledMappings.Count == 0)
            {
                await _winFspMountService.EnsureCatalogMountedAsync(baseDriveLetter, enabledMappings, CancellationToken.None);
                await RefreshSharedFolderMountStatesSafeAsync();
                _sharedFolderStatusText.Text = $"{prefix}: Laufwerk {baseDriveLetter}: bleibt verbunden · keine aktiven Shares.";
                _sharedFolderLastError = "-";
                return;
            }

            try
            {
                await _winFspMountService.EnsureCatalogMountedAsync(baseDriveLetter, enabledMappings, CancellationToken.None);
                await RefreshSharedFolderMountStatesSafeAsync();
                _sharedFolderStatusText.Text = $"{prefix}: 1 Laufwerk ({baseDriveLetter}:) aktiv mit {enabledMappings.Count} Share-Ordnern.";
                _sharedFolderLastError = "-";

                GuestLogger.Info("sharedfolders.apply.catalog_ready", "Shared-Folder Katalog-Mount aktiv.", new
                {
                    driveLetter = baseDriveLetter,
                    count = enabledMappings.Count,
                    mode = "hypertool-file-catalog"
                });
            }
            catch (Exception ex)
            {
                await RefreshSharedFolderMountStatesSafeAsync();
                _sharedFolderLastError = ex.Message;
                _sharedFolderStatusText.Text = $"{prefix}: Fehler beim Katalog-Mount ({baseDriveLetter}:). {ex.Message}";
                GuestLogger.Warn("sharedfolders.apply.catalog_failed", ex.Message, new
                {
                    driveLetter = baseDriveLetter,
                    count = enabledMappings.Count,
                    exceptionType = ex.GetType().FullName
                });
            }
        }
        finally
        {
            _suppressSharedFolderAutoApply = false;
            _sharedFolderUiOperationGate.Release();
        }
    }

    private string? ResolveSharedFolderHostTarget()
    {
        var hostName = (_config.Usb?.HostName ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(hostName))
        {
            return hostName;
        }

        return (_config.Usb?.HostAddress ?? string.Empty).Trim();
    }

    private void UpdateHostDiscoveryPresentation()
    {
        var hostName = (_config.Usb?.HostName ?? string.Empty).Trim();
        var hostAddress = (_config.Usb?.HostAddress ?? string.Empty).Trim();

        if (!string.IsNullOrWhiteSpace(hostName))
        {
            _usbResolvedHostNameText.Text = $"Ermittelter Hostname: {hostName}";
            return;
        }

        if (!string.IsNullOrWhiteSpace(hostAddress))
        {
            _usbResolvedHostNameText.Text = $"Kein Hostname gefunden · Fallback-Ziel: {hostAddress}";
            return;
        }

        _usbResolvedHostNameText.Text = "Ermittelter Hostname: -";
    }

    private async Task SaveSharedFolderMappingsToConfigAsync()
    {
        _config.SharedFolders ??= new GuestSharedFolderSettings();
        _config.SharedFolders.Enabled = _sharedFolderFeatureEnabledToggleSwitch.IsOn;
        _config.SharedFolders.BaseDriveLetter = GetSharedFolderBaseDriveLetter();
        _config.SharedFolders.Mappings = _sharedFolderMappings
            .Select(item => new GuestSharedFolderMapping
            {
                Id = string.IsNullOrWhiteSpace(item.Id) ? Guid.NewGuid().ToString("N") : item.Id,
                Label = ResolveMappingLabel(item.Label, item.SharePath),
                SharePath = item.SharePath,
                DriveLetter = _config.SharedFolders.BaseDriveLetter,
                Persistent = true,
                Enabled = item.Enabled
            })
            .ToList();

        await _saveConfigAsync(_config);
    }

    private static void EnsureUniqueDriveLetters(IList<GuestSharedFolderMapping> mappings)
    {
        var usedLetters = new HashSet<char>();

        foreach (var mapping in mappings)
        {
            var normalizedLetter = GuestConfigService.NormalizeDriveLetter(mapping.DriveLetter)[0];
            if (usedLetters.Contains(normalizedLetter))
            {
                mapping.DriveLetter = GetNextAvailableDriveLetter(usedLetters);
                normalizedLetter = mapping.DriveLetter[0];
            }

            usedLetters.Add(normalizedLetter);
            mapping.DriveLetter = normalizedLetter.ToString();
            mapping.Persistent = true;
        }
    }

    private static string GetNextAvailableDriveLetter(ISet<char> usedLetters)
    {
        for (var letter = 'Z'; letter >= 'D'; letter--)
        {
            if (!usedLetters.Contains(letter))
            {
                return letter.ToString();
            }
        }

        return "Z";
    }

    private bool ValidateSharedFolderMappings(out string error)
    {
        error = string.Empty;

        _config.SharedFolders ??= new GuestSharedFolderSettings();
        _config.SharedFolders.BaseDriveLetter = GuestConfigService.NormalizeDriveLetter(_config.SharedFolders.BaseDriveLetter);

        var enabledMappings = _sharedFolderMappings.Where(mapping => mapping.Enabled).ToList();
        if (enabledMappings.Count == 0)
        {
            return true;
        }

        foreach (var mapping in _sharedFolderMappings)
        {
            mapping.DriveLetter = _config.SharedFolders.BaseDriveLetter;
            mapping.Persistent = true;
        }

        return true;
    }

    private async Task SyncSharedFoldersFromHostAsync(bool forceHostFeatureRefresh = true)
    {
        if (!await _sharedFolderUiOperationGate.WaitAsync(TimeSpan.FromSeconds(8)))
        {
            _sharedFolderStatusText.Text = "Shared-Folder Aktion läuft noch. Bitte in ein paar Sekunden erneut versuchen.";
            return;
        }

        try
        {
            if (forceHostFeatureRefresh)
            {
                await RefreshHostFeatureAvailabilityFromSocketAsync();
            }

            var hostFolders = await _fetchHostSharedFoldersAsync();
            var existingByShareName = _sharedFolderMappings
                .Where(mapping => !string.IsNullOrWhiteSpace(mapping.SharePath))
                .Select(mapping => new
                {
                    Mapping = mapping,
                    ShareName = ExtractShareNameFromUncPath(mapping.SharePath)
                })
                .Where(item => !string.IsNullOrWhiteSpace(item.ShareName))
                .GroupBy(item => item.ShareName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First().Mapping, StringComparer.OrdinalIgnoreCase);

            var resolvedTarget = ResolveSharedFolderHostTarget();
            var baseDriveLetter = GetSharedFolderBaseDriveLetter();
            var synchronizedMappings = new List<GuestSharedFolderMapping>();
            var importedCount = 0;

            foreach (var hostFolder in hostFolders.Where(item => item.Enabled && !string.IsNullOrWhiteSpace(item.ShareName)))
            {
                var normalizedShareName = hostFolder.ShareName.Trim();
                var normalizedLabel = string.IsNullOrWhiteSpace(hostFolder.Label)
                    ? normalizedShareName
                    : hostFolder.Label.Trim();
                var sharePath = BuildCatalogSharePath(normalizedShareName, resolvedTarget);

                if (existingByShareName.TryGetValue(normalizedShareName, out var existing))
                {
                    synchronizedMappings.Add(new GuestSharedFolderMapping
                    {
                        Id = string.IsNullOrWhiteSpace(existing.Id) ? Guid.NewGuid().ToString("N") : existing.Id,
                        Label = normalizedLabel,
                        SharePath = sharePath,
                        DriveLetter = baseDriveLetter,
                        Persistent = true,
                        Enabled = existing.Enabled,
                        MountStateDot = existing.MountStateDot,
                        MountStateText = existing.MountStateText
                    });
                    continue;
                }

                synchronizedMappings.Add(new GuestSharedFolderMapping
                {
                    Id = Guid.NewGuid().ToString("N"),
                    Label = normalizedLabel,
                    SharePath = sharePath,
                    DriveLetter = baseDriveLetter,
                    Persistent = true,
                    Enabled = false
                });

                importedCount++;
            }

            var removedCount = Math.Max(0, _sharedFolderMappings.Count - synchronizedMappings.Count);

            _suppressSharedFolderAutoApply = true;
            try
            {
                foreach (var existing in _sharedFolderMappings)
                {
                    existing.PropertyChanged -= OnSharedFolderMappingPropertyChanged;
                }

                _sharedFolderMappings.Clear();
                foreach (var mapping in synchronizedMappings)
                {
                    _sharedFolderMappings.Add(mapping);
                    mapping.PropertyChanged += OnSharedFolderMappingPropertyChanged;
                }
            }
            finally
            {
                _suppressSharedFolderAutoApply = false;
            }

            _config.SharedFolders ??= new GuestSharedFolderSettings();
            await SaveSharedFolderMappingsToConfigAsync();
            await RefreshSharedFolderMountStatesSafeAsync();
            UpdateHostDiscoveryPresentation();

            if (synchronizedMappings.Count == 0)
            {
                _sharedFolderStatusText.Text = "Keine Shared-Folder vom Host empfangen.";
                return;
            }

            var targetSuffix = string.IsNullOrWhiteSpace(resolvedTarget)
                ? string.Empty
                : $" · Ziel: {resolvedTarget}";

            _sharedFolderStatusText.Text = removedCount == 0
                ? (importedCount == 0
                    ? $"Host-Liste geladen, keine Änderungen.{targetSuffix}"
                    : $"Host-Liste geladen: {importedCount} neue Shared-Folder.{targetSuffix}")
                : $"Host-Liste geladen: {importedCount} neu, {removedCount} veraltete Einträge entfernt.{targetSuffix}";
        }
        catch (Exception ex)
        {
            _sharedFolderStatusText.Text = $"Host-Liste konnte nicht geladen werden: {ex.Message}";
            AppendNotification($"[Warn] Host Shared-Folder Sync fehlgeschlagen: {ex.Message}");
            _sharedFolderLastError = ex.Message;
        }
        finally
        {
            _sharedFolderUiOperationGate.Release();
        }
    }

    private static string BuildCatalogSharePath(string shareName, string? hostTarget)
    {
        var normalizedShareName = (shareName ?? string.Empty).Trim();
        var normalizedHostTarget = (hostTarget ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalizedHostTarget))
        {
            normalizedHostTarget = "HOST";
        }

        return $"\\\\{normalizedHostTarget}\\{normalizedShareName}";
    }

    private static string ExtractShareNameFromUncPath(string? uncPath)
    {
        var normalizedPath = (uncPath ?? string.Empty).Trim();
        if (!normalizedPath.StartsWith("\\\\", StringComparison.Ordinal))
        {
            return string.Empty;
        }

        var withoutPrefix = normalizedPath[2..];
        var firstSlash = withoutPrefix.IndexOf('\\');
        if (firstSlash <= 0)
        {
            return string.Empty;
        }

        var rest = withoutPrefix[(firstSlash + 1)..];
        var secondSlash = rest.IndexOf('\\');
        var shareName = (secondSlash >= 0 ? rest[..secondSlash] : rest).Trim();
        return shareName;
    }

    private static string ResolveMappingLabel(string? preferredLabel, string? sharePath)
    {
        var normalizedLabel = (preferredLabel ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(normalizedLabel))
        {
            return normalizedLabel;
        }

        var extractedShareName = ExtractShareNameFromUncPath(sharePath);
        if (!string.IsNullOrWhiteSpace(extractedShareName))
        {
            return extractedShareName;
        }

        return (sharePath ?? string.Empty).Trim();
    }

    private async Task RunSharedFolderSelfTestAsync()
    {
        if (!await _sharedFolderUiOperationGate.WaitAsync(TimeSpan.FromSeconds(8)))
        {
            _sharedFolderStatusText.Text = "Shared-Folder Aktion läuft noch. Bitte in ein paar Sekunden erneut versuchen.";
            return;
        }

        try
        {
            var hostFolders = await _fetchHostSharedFoldersAsync();
            var hostShareNames = new HashSet<string>(
                hostFolders
                    .Where(folder => folder.Enabled && !string.IsNullOrWhiteSpace(folder.ShareName))
                    .Select(folder => folder.ShareName.Trim()),
                StringComparer.OrdinalIgnoreCase);

            var enabledMappings = _sharedFolderMappings.Where(mapping => mapping.Enabled).ToList();
            var sharePresentCount = 0;
            var mappingPresentCount = 0;
            var featureEnabled = IsSharedFolderFeatureEnabled();
            var baseDriveLetter = GetSharedFolderBaseDriveLetter();
            var mounted = featureEnabled && _winFspMountService.IsMounted(baseDriveLetter);

            foreach (var mapping in enabledMappings)
            {
                var shareName = ExtractShareNameFromUncPath(mapping.SharePath);

                if (!string.IsNullOrWhiteSpace(shareName) && hostShareNames.Contains(shareName))
                {
                    sharePresentCount++;
                }

                if (mounted)
                {
                    mappingPresentCount++;
                }
            }

            var sharePresentText = enabledMappings.Count == 0
                ? "n/a"
                : (sharePresentCount == enabledMappings.Count ? "Ja" : $"Teilweise ({sharePresentCount}/{enabledMappings.Count})");
            var mappingPresentText = enabledMappings.Count == 0
                ? "n/a"
                : (mappingPresentCount == enabledMappings.Count ? "Ja" : $"Teilweise ({mappingPresentCount}/{enabledMappings.Count})");

            _sharedFolderStatusText.Text = $"Self-Test · Share vorhanden: {sharePresentText} · Mapping vorhanden: {mappingPresentText} · Letzter Fehler: {_sharedFolderLastError}";
        }
        catch (Exception ex)
        {
            _sharedFolderLastError = ex.Message;
            _sharedFolderStatusText.Text = $"Self-Test fehlgeschlagen: {ex.Message}";
            GuestLogger.Warn("sharedfolders.selftest.failed", ex.Message, new
            {
                exceptionType = ex.GetType().FullName
            });
        }
        finally
        {
            _sharedFolderUiOperationGate.Release();
        }
    }

    private UIElement BuildSettingsPage()
    {
        var root = new StackPanel { Spacing = 12 };

        var headingCard = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush,
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(14)
        };
        var headingGrid = new Grid { ColumnSpacing = 10, VerticalAlignment = VerticalAlignment.Center };
        headingGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headingGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var headingStack = new StackPanel
        {
            Spacing = 4,
            Children =
            {
                new TextBlock { Text = "Konfiguration", FontSize = 20, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold },
                new TextBlock { Text = "Wichtige Einstellungen übersichtlich und schnell erreichbar.", Opacity = 0.9 }
            }
        };
        headingGrid.Children.Add(headingStack);

        var configHeaderActions = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 8,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Center
        };

        configHeaderActions.Children.Add(CreateIconButton("💾", "Speichern", onClick: async (_, _) => await SaveSettingsAsync()));
        configHeaderActions.Children.Add(CreateIconButton("⟳", "Neu laden", onClick: async (_, _) =>
        {
            try
            {
                await ReloadConfigFromDiskAsync("[Info] Einstellungen aus Datei neu geladen.");
            }
            catch (Exception ex)
            {
                AppendNotification($"[Error] Konfiguration konnte nicht neu geladen werden: {ex.Message}");
            }
        }));
        configHeaderActions.Children.Add(CreateIconButton(ToolRestartIcon, ToolRestartLabel, onClick: async (_, _) => await RestartGuestToolAsync()));

        Grid.SetColumn(configHeaderActions, 1);
        headingGrid.Children.Add(configHeaderActions);

        headingCard.Child = headingGrid;
        root.Children.Add(headingCard);

        var systemSection = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(14)
        };

        var systemStack = new StackPanel { Spacing = 10 };
        systemStack.Children.Add(new TextBlock { Text = "System & Updates", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, FontSize = 16 });

        var quickTogglesGrid = new Grid { ColumnSpacing = 26, RowSpacing = 6, HorizontalAlignment = HorizontalAlignment.Left };
        quickTogglesGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        quickTogglesGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        quickTogglesGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        quickTogglesGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        quickTogglesGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        _minimizeToTrayCheckBox.Margin = new Thickness(0);
        Grid.SetColumn(_minimizeToTrayCheckBox, 0);
        Grid.SetRow(_minimizeToTrayCheckBox, 0);
        quickTogglesGrid.Children.Add(_minimizeToTrayCheckBox);

        _startMinimizedCheckBox.Margin = new Thickness(0);
        Grid.SetColumn(_startMinimizedCheckBox, 1);
        Grid.SetRow(_startMinimizedCheckBox, 0);
        quickTogglesGrid.Children.Add(_startMinimizedCheckBox);

        _startWithWindowsCheckBox.Margin = new Thickness(0);
        Grid.SetColumn(_startWithWindowsCheckBox, 0);
        Grid.SetRow(_startWithWindowsCheckBox, 1);
        quickTogglesGrid.Children.Add(_startWithWindowsCheckBox);

        _checkForUpdatesOnStartupCheckBox.Margin = new Thickness(0);
        _checkForUpdatesOnStartupCheckBox.HorizontalAlignment = HorizontalAlignment.Left;
        Grid.SetColumn(_checkForUpdatesOnStartupCheckBox, 1);
        Grid.SetRow(_checkForUpdatesOnStartupCheckBox, 1);
        quickTogglesGrid.Children.Add(_checkForUpdatesOnStartupCheckBox);

        systemStack.Children.Add(quickTogglesGrid);

        var themeRow = new Grid { ColumnSpacing = 8, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0) };
        themeRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
        themeRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        themeRow.Children.Add(new TextBlock
        {
            Text = "Dark Mode",
            VerticalAlignment = VerticalAlignment.Center,
            Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
        });

        _themeToggle.OnContent = "An";
        _themeToggle.OffContent = "Aus";
        _themeToggle.HorizontalAlignment = HorizontalAlignment.Left;
        _themeToggle.MinWidth = 86;
        if (!_isThemeToggleHandlerAttached)
        {
            _themeToggle.Toggled += async (_, _) =>
            {
                if (_suppressThemeEvents)
                {
                    return;
                }

                _themeCombo.SelectedItem = _themeToggle.IsOn ? "dark" : "light";
                await ApplyThemeAndRestartImmediatelyAsync();
            };
            _isThemeToggleHandlerAttached = true;
        }
        Grid.SetColumn(_themeToggle, 1);
        themeRow.Children.Add(_themeToggle);

        systemStack.Children.Add(themeRow);
        systemSection.Child = systemStack;
        root.Children.Add(systemSection);

        var usbSection = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(14)
        };

        var usbStack = new StackPanel { Spacing = 8 };

        _usbHyperVModeIconBadge.Child = _usbHyperVModeIcon;
        _usbIpModeIconBadge.Child = _usbIpModeIcon;

        var hyperVModeBadgeContent = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            VerticalAlignment = VerticalAlignment.Center
        };
        hyperVModeBadgeContent.Children.Add(_usbHyperVModeIconBadge);
        hyperVModeBadgeContent.Children.Add(_usbHyperVModeBadgeText);
        _usbHyperVModeBadge.Child = hyperVModeBadgeContent;

        var ipModeBadgeContent = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            VerticalAlignment = VerticalAlignment.Center
        };
        ipModeBadgeContent.Children.Add(_usbIpModeIconBadge);
        ipModeBadgeContent.Children.Add(_usbIpModeBadgeText);
        _usbIpModeBadge.Child = ipModeBadgeContent;

        if (!_isUsbModeBadgeHandlersAttached)
        {
            _usbHyperVModeBadge.Tapped += (_, _) =>
            {
                if (_useHyperVSocketCheckBox.IsChecked != true)
                {
                    _useHyperVSocketCheckBox.IsChecked = true;
                }
            };

            _usbIpModeBadge.Tapped += (_, _) =>
            {
                if (_useHyperVSocketCheckBox.IsChecked != false)
                {
                    _useHyperVSocketCheckBox.IsChecked = false;
                }
            };

            _isUsbModeBadgeHandlersAttached = true;
        }

        var modeBadgeRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 8,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Top
        };
        modeBadgeRow.Children.Add(_usbHyperVModeBadge);
        modeBadgeRow.Children.Add(_usbIpModeBadge);

        var usbHeaderGrid = new Grid { ColumnSpacing = 10 };
        usbHeaderGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        usbHeaderGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var usbHeaderTextStack = new StackPanel { Spacing = 4 };
        var usbHeaderTitleRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            VerticalAlignment = VerticalAlignment.Center
        };
        usbHeaderTitleRow.Children.Add(new TextBlock
        {
            Text = "USB Host-Verbindung",
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            FontSize = 16,
            VerticalAlignment = VerticalAlignment.Center,
            Opacity = 0.9
        });
        usbHeaderTextStack.Children.Add(usbHeaderTitleRow);
        usbHeaderGrid.Children.Add(usbHeaderTextStack);

        Grid.SetColumn(modeBadgeRow, 1);
        usbHeaderGrid.Children.Add(modeBadgeRow);
        usbStack.Children.Add(usbHeaderGrid);

        _useHyperVSocketCheckBox.Margin = new Thickness(0, 2, 0, 0);
        if (!_isUsbTransportToggleHandlerAttached)
        {
            _useHyperVSocketCheckBox.Checked += async (_, _) => await OnUsbTransportModeToggledAsync();
            _useHyperVSocketCheckBox.Unchecked += async (_, _) => await OnUsbTransportModeToggledAsync();
            _isUsbTransportToggleHandlerAttached = true;
        }
        usbStack.Children.Add(_useHyperVSocketCheckBox);

        _usbModeHintText.Foreground = Application.Current.Resources["TextMutedBrush"] as Brush;
        usbStack.Children.Add(_usbModeHintText);

        _usbHostAddressEditorCard.BorderThickness = new Thickness(0);
        _usbHostAddressEditorCard.BorderBrush = new SolidColorBrush(Color.FromArgb(0x00, 0x00, 0x00, 0x00));
        _usbHostAddressEditorCard.Background = new SolidColorBrush(Color.FromArgb(0x00, 0x00, 0x00, 0x00));
        _usbHostAddressEditorCard.Padding = new Thickness(0);

        _usbHostAddressTextBox.PlaceholderText = "Beispiel: HOSTNAME oder 172.25.80.1";
        _usbHostAddressTextBox.MinWidth = 420;
        _usbHostAddressTextBox.MaxWidth = 620;
        _usbHostAddressTextBox.HorizontalAlignment = HorizontalAlignment.Left;
        _usbHostAddressTextBox.CornerRadius = new CornerRadius(8);
        _usbHostAddressTextBox.TextChanged += (_, _) => UpdateUsbTransportModePresentation();

        var usbHostAddressRow = new Grid { ColumnSpacing = 8 };
        usbHostAddressRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        usbHostAddressRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        usbHostAddressRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        usbHostAddressRow.Children.Add(_usbHostAddressTextBox);

        _usbHostSearchButton = CreateIconButton("🔎", "Host suchen", onClick: async (_, _) => await SearchUsbHostAddressAsync());
        _usbHostSearchButton.HorizontalAlignment = HorizontalAlignment.Left;
        _usbHostSearchButton.VerticalAlignment = VerticalAlignment.Center;
        Grid.SetColumn(_usbHostSearchButton, 1);
        usbHostAddressRow.Children.Add(_usbHostSearchButton);

        SetUsbHostSearchStatus("Bereit", UsbHostSearchStatusKind.Neutral);
        Grid.SetColumn(_usbHostSearchStatusText, 2);
        usbHostAddressRow.Children.Add(_usbHostSearchStatusText);

        _usbHostAddressEditorCard.Child = usbHostAddressRow;
        usbStack.Children.Add(_usbHostAddressEditorCard);
        usbStack.Children.Add(_usbResolvedHostNameText);

        UpdateUsbTransportModePresentation();
        UpdateHostDiscoveryPresentation();

        usbSection.Child = usbStack;
        root.Children.Add(usbSection);

        return new ScrollViewer { Content = root };
    }

    private UIElement BuildInfoPage()
    {
        var version = ResolveGuestVersionText();

        var panel = new StackPanel { Spacing = 10 };
        var titleWrap = new Grid();
        titleWrap.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        titleWrap.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        titleWrap.Children.Add(new TextBlock { Text = "Info", FontSize = 20, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });

        var versionWrap = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, VerticalAlignment = VerticalAlignment.Bottom };
        versionWrap.Children.Add(new TextBlock { Text = "Version:", Opacity = 0.9, VerticalAlignment = VerticalAlignment.Bottom });
        versionWrap.Children.Add(new TextBlock { Opacity = 0.9, Text = version });
        Grid.SetColumn(versionWrap, 1);
        titleWrap.Children.Add(versionWrap);

        panel.Children.Add(titleWrap);

        var infoStatusRow = new Grid { ColumnSpacing = 8 };
        infoStatusRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        infoStatusRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var updateWrap = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, VerticalAlignment = VerticalAlignment.Center };
        updateWrap.Children.Add(new TextBlock { Text = "Update-Status:", Opacity = 0.9 });
        updateWrap.Children.Add(_updateStatusValueText);

        var copyrightText = new TextBlock
        {
            Text = "Copyright: koerby",
            Opacity = 0.9,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Center
        };

        Grid.SetColumn(updateWrap, 0);
        infoStatusRow.Children.Add(updateWrap);
        Grid.SetColumn(copyrightText, 1);
        infoStatusRow.Children.Add(copyrightText);
        panel.Children.Add(infoStatusRow);

        var projectCard = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PageBackgroundBrush"] as Brush,
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(10)
        };

        var projectStack = new StackPanel { Spacing = 4 };
        projectStack.Children.Add(new TextBlock { Text = "HyperTool Projekt", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
        projectStack.Children.Add(new TextBlock { Text = "HyperTool Guest wird über GitHub Releases verteilt. Hier findest du Version, Update-Status und Release-Links.", TextWrapping = TextWrapping.Wrap, Opacity = 0.85 });
        projectStack.Children.Add(new TextBlock { Text = "GitHub Owner: koerby", Opacity = 0.9 });
        projectStack.Children.Add(new TextBlock { Text = "GitHub Repo: HyperTool", Opacity = 0.9 });
        projectCard.Child = projectStack;
        panel.Children.Add(projectCard);

        var usbipCard = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PageBackgroundBrush"] as Brush,
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(10)
        };

        var usbipStack = new StackPanel { Spacing = 4 };
        usbipStack.Children.Add(new TextBlock { Text = "Externe USB/IP Quelle", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
        usbipStack.Children.Add(new HyperlinkButton { Content = "Quelle: vadimgrn/usbip-win2", NavigateUri = new Uri("https://github.com/vadimgrn/usbip-win2"), Padding = new Thickness(0), Opacity = 0.9 });
        usbipStack.Children.Add(new TextBlock { Text = "Nutzung in HyperTool: externer CLI-Client ohne eigene GUI-Integration.", TextWrapping = TextWrapping.Wrap, Opacity = 0.85 });
        usbipStack.Children.Add(new TextBlock { Text = "Lizenz/Eigentümer: siehe Original-Repository von vadimgrn.", TextWrapping = TextWrapping.Wrap, Opacity = 0.85 });
        usbipCard.Child = usbipStack;

        var winfspCard = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PageBackgroundBrush"] as Brush,
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(10)
        };

        var winfspStack = new StackPanel { Spacing = 4 };
        winfspStack.Children.Add(new TextBlock { Text = "Externe Shared-Folder Runtime", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
        winfspStack.Children.Add(new HyperlinkButton { Content = $"Quelle: {WinFspRuntimeOwner}/{WinFspRuntimeRepo}", NavigateUri = new Uri($"https://github.com/{WinFspRuntimeOwner}/{WinFspRuntimeRepo}"), Padding = new Thickness(0), Opacity = 0.9 });
        winfspStack.Children.Add(new TextBlock { Text = "Nutzung in HyperTool: externer Runtime-Treiber für Guest Shared-Folder-Mounts.", TextWrapping = TextWrapping.Wrap, Opacity = 0.85 });
        winfspStack.Children.Add(new TextBlock { Text = "Lizenz/Eigentümer: siehe Original-Repository von winfsp.", TextWrapping = TextWrapping.Wrap, Opacity = 0.85 });
        winfspCard.Child = winfspStack;

        var externalSourcesGrid = new Grid
        {
            ColumnSpacing = 10,
            HorizontalAlignment = HorizontalAlignment.Stretch
        };
        externalSourcesGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        externalSourcesGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        Grid.SetColumn(usbipCard, 0);
        externalSourcesGrid.Children.Add(usbipCard);
        Grid.SetColumn(winfspCard, 1);
        externalSourcesGrid.Children.Add(winfspCard);

        panel.Children.Add(externalSourcesGrid);

        var diagnosticsCard = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PageBackgroundBrush"] as Brush,
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(10)
        };

        var diagnosticsStack = new StackPanel
        {
            Spacing = 4,
            Margin = new Thickness(0, 0, 240, 0)
        };
        diagnosticsStack.Children.Add(new TextBlock
        {
            Text = "Transport Diagnose",
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
        });

        var hyperVSocketRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, VerticalAlignment = VerticalAlignment.Center };
        hyperVSocketRow.Children.Add(new TextBlock { Text = "Hyper-V Socket aktiv:", Opacity = 0.9, VerticalAlignment = VerticalAlignment.Center });
        hyperVSocketRow.Children.Add(_diagHyperVSocketText);
        diagnosticsStack.Children.Add(hyperVSocketRow);

        var registryRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, VerticalAlignment = VerticalAlignment.Center };
        registryRow.Children.Add(new TextBlock { Text = "Registry-Service erreichbar:", Opacity = 0.9, VerticalAlignment = VerticalAlignment.Center });
        registryRow.Children.Add(_diagRegistryServiceText);
        diagnosticsStack.Children.Add(registryRow);

        var fallbackRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, VerticalAlignment = VerticalAlignment.Center };
        fallbackRow.Children.Add(new TextBlock { Text = "Fallback auf IP aktiv:", Opacity = 0.9, VerticalAlignment = VerticalAlignment.Center });
        fallbackRow.Children.Add(_diagFallbackText);
        diagnosticsStack.Children.Add(fallbackRow);

        var diagnosticsLayout = new Grid();
        diagnosticsLayout.Children.Add(diagnosticsStack);

        var diagnosticsTestButton = CreateIconButton("🧪", "Hyper-V Socket testen", onClick: async (_, _) => await RunTransportDiagnosticsTestAsync());
        diagnosticsTestButton.HorizontalAlignment = HorizontalAlignment.Right;
        diagnosticsTestButton.VerticalAlignment = VerticalAlignment.Top;
        diagnosticsLayout.Children.Add(diagnosticsTestButton);

        diagnosticsCard.Child = diagnosticsLayout;
        panel.Children.Add(diagnosticsCard);

        var buttonRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        buttonRow.Children.Add(CreateIconButton("🛰", "Update prüfen", onClick: async (_, _) => await CheckForUpdatesAsync()));
        _installUpdateButton = CreateIconButton("⬇", "Update installieren", onClick: async (_, _) => await InstallUpdateAsync());
        _installUpdateButton.IsEnabled = CanInstallUpdate();
        buttonRow.Children.Add(_installUpdateButton);
        buttonRow.Children.Add(CreateIconButton("🌐", "Changelog / Release", onClick: (_, _) => OpenReleasePage()));
        buttonRow.Children.Add(CreateSupportCoffeeButton((_, _) =>
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "https://buymeacoffee.com/koerby",
                UseShellExecute = true
            });
        }));
        panel.Children.Add(buttonRow);

        return new ScrollViewer { Content = panel };
    }

    private static string ResolveGuestVersionText()
    {
        var informationalVersion = Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;

        var raw = string.IsNullOrWhiteSpace(informationalVersion)
            ? Assembly.GetExecutingAssembly().GetName().Version?.ToString()
            : informationalVersion;

        if (string.IsNullOrWhiteSpace(raw))
        {
            return "2.1.0";
        }

        var sanitized = raw.Split('+', 2)[0].Trim();

        if (Version.TryParse(sanitized, out var parsed))
        {
            return parsed.Revision == 0
                ? $"{parsed.Major}.{parsed.Minor}.{parsed.Build}"
                : parsed.ToString();
        }

        return sanitized;
    }

    private void UpdatePageContent()
    {
        _pageContent.Content = _selectedMenuIndex switch
        {
            0 => _usbPage ??= BuildUsbPage(),
            1 => _sharedFoldersPage ??= BuildSharedFoldersPage(),
            2 => GetOrCreateSettingsPage(),
            _ => _infoPage ??= BuildInfoPage()
        };
    }

    private UIElement GetOrCreateSettingsPage()
    {
        _settingsPage ??= BuildSettingsPage();
        ApplyConfigToControls();
        return _settingsPage;
    }

    private async Task<bool> TryHandlePendingSettingsBeforeMenuSwitchAsync(int targetMenuIndex)
    {
        if (_selectedMenuIndex != SettingsMenuIndex || targetMenuIndex == SettingsMenuIndex || !HasPendingSettingsChanges())
        {
            return true;
        }

        if (_isMenuSwitchPromptOpen)
        {
            return false;
        }

        _isMenuSwitchPromptOpen = true;
        try
        {
            var result = await ShowUnsavedSettingsPromptAsync();

            if (result == ContentDialogResult.None)
            {
                return false;
            }

            if (result == ContentDialogResult.Primary)
            {
                await SaveSettingsAsync();
                return !HasPendingSettingsChanges();
            }

            await ReloadConfigFromDiskAsync("[Info] Ungespeicherte Änderungen verworfen. Konfiguration neu geladen.");
            var hasPendingAfterReload = HasPendingSettingsChanges();
            if (hasPendingAfterReload)
            {
                AppendNotification("[Warn] Einstellungen konnten nicht vollständig zurückgesetzt werden. Bitte erneut 'Neu laden' ausführen.");
            }

            return !hasPendingAfterReload;
        }
        catch (Exception ex)
        {
            AppendNotification($"[Error] Konfiguration konnte nicht neu geladen werden: {ex.Message}");
            return false;
        }
        finally
        {
            _isMenuSwitchPromptOpen = false;
        }
    }

    private bool HasPendingSettingsChanges()
    {
        var selectedTheme = GuestConfigService.NormalizeTheme((_themeCombo.SelectedItem as string) ?? _config.Ui.Theme);
        var currentTheme = GuestConfigService.NormalizeTheme(_config.Ui.Theme);
        if (!string.Equals(selectedTheme, currentTheme, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if ((_startWithWindowsCheckBox.IsChecked == true) != _config.Ui.StartWithWindows
            || (_startMinimizedCheckBox.IsChecked == true) != _config.Ui.StartMinimized
            || (_minimizeToTrayCheckBox.IsChecked == true) != _config.Ui.MinimizeToTray
            || (_checkForUpdatesOnStartupCheckBox.IsChecked != false) != _config.Ui.CheckForUpdatesOnStartup)
        {
            return true;
        }

        if ((_usbFeatureEnabledToggleSwitch.IsOn) != (_config.Usb?.Enabled != false)
            || (_usbDisconnectOnExitCheckBox.IsChecked != false) != (_config.Usb?.DisconnectOnExit != false)
            || (_useHyperVSocketCheckBox.IsChecked != false) != (_config.Usb?.UseHyperVSocket != false))
        {
            return true;
        }

        var configuredHost = (_config.Usb?.HostAddress ?? string.Empty).Trim();
        var editorHost = (_usbHostAddressTextBox.Text ?? string.Empty).Trim();
        return !string.Equals(editorHost, configuredHost, StringComparison.OrdinalIgnoreCase);
    }

    private async Task<ContentDialogResult> ShowUnsavedSettingsPromptAsync()
    {
        if (Content is not FrameworkElement root || root.XamlRoot is null)
        {
            return ContentDialogResult.None;
        }

        var dialog = new ContentDialog
        {
            XamlRoot = root.XamlRoot,
            Title = "Ungespeicherte Änderungen",
            Content = "Es gibt ungespeicherte Änderungen in den Einstellungen. Jetzt speichern?",
            PrimaryButtonText = "Speichern",
            SecondaryButtonText = "Nein",
            CloseButtonText = "Abbrechen",
            DefaultButton = ContentDialogButton.Primary
        };

        return await dialog.ShowAsync();
    }

    private async Task ReloadConfigFromDiskAsync(string? successNotification = null)
    {
        var reloadedConfig = await _reloadConfigAsync();
        _config = reloadedConfig ?? new GuestConfig();
        _config.Usb ??= new GuestUsbSettings();
        _config.SharedFolders ??= new GuestSharedFolderSettings();
        _config.Ui ??= new GuestUiSettings();
        _config.FileService ??= new GuestFileServiceSettings();
        ApplyConfigToControls();

        if (!string.IsNullOrWhiteSpace(successNotification))
        {
            AppendNotification(successNotification);
        }
    }

    private void UpdateNavSelection()
    {
        for (var index = 0; index < _navButtons.Count; index++)
        {
            var isSelected = index == _selectedMenuIndex;
            _navButtons[index].Background = isSelected
                ? Application.Current.Resources["AccentSoftBrush"] as Brush
                : Application.Current.Resources["SurfaceSoftBrush"] as Brush;
            _navButtons[index].BorderBrush = isSelected
                ? Application.Current.Resources["AccentBrush"] as Brush
                : Application.Current.Resources["PanelBorderBrush"] as Brush;
        }
    }

    private void ApplyConfigToControls()
    {
        var isDark = string.Equals(GuestConfigService.NormalizeTheme(_config.Ui.Theme), "dark", StringComparison.OrdinalIgnoreCase);

        _suppressThemeEvents = true;
        _themeCombo.SelectedItem = isDark ? "dark" : "light";
        _themeToggle.IsOn = isDark;
        _suppressThemeEvents = false;

        _themeText.Text = isDark ? "Dunkles Theme" : "Helles Theme";

        _usbHostAddressTextBox.Text = (_config.Usb?.HostAddress ?? string.Empty).Trim();
        _startWithWindowsCheckBox.IsChecked = _config.Ui.StartWithWindows;
        _startMinimizedCheckBox.IsChecked = _config.Ui.StartMinimized;
        _minimizeToTrayCheckBox.IsChecked = _config.Ui.MinimizeToTray;
        _checkForUpdatesOnStartupCheckBox.IsChecked = _config.Ui.CheckForUpdatesOnStartup;

        _config.FileService ??= new GuestFileServiceSettings();
        _config.FileService.MappingMode = GuestConfigService.NormalizeMappingMode(_config.FileService.MappingMode);

        _suppressUsbTransportToggleEvents = true;
        try
        {
            _useHyperVSocketCheckBox.IsChecked = _config.Usb?.UseHyperVSocket != false;
        }
        finally
        {
            _suppressUsbTransportToggleEvents = false;
        }

        _suppressUsbFeatureToggle = true;
        try
        {
            _usbFeatureEnabledToggleSwitch.IsOn = _config.Usb?.Enabled != false;
        }
        finally
        {
            _suppressUsbFeatureToggle = false;
        }

        _suppressUsbDisconnectOnExitToggleEvents = true;
        try
        {
            _usbDisconnectOnExitCheckBox.IsChecked = _config.Usb?.DisconnectOnExit != false;
        }
        finally
        {
            _suppressUsbDisconnectOnExitToggleEvents = false;
        }

        UpdateUsbTransportModePresentation();
        UpdateHostDiscoveryPresentation();
        UpdateAutoConnectToggleFromSelection();
        UpdateUsbRuntimeStatusUi();

        RefreshSharedFolderMappingsFromConfig();
        UpdateSharedFolderFeatureUi();
        if (_sharedFoldersPage is not null)
        {
            _sharedFolderMappingsListView.ItemsSource = null;
            _sharedFolderMappingsListView.ItemsSource = _sharedFolderMappings;
            _ = RefreshSharedFolderMountStatesSafeAsync();
        }
    }

    public void UpdateHostFeatureAvailability(bool usbFeatureEnabledByHost, bool sharedFoldersFeatureEnabledByHost, string? hostName)
    {
        var previousUsbFeatureEnabledByHost = IsUsbFeatureEnabledByHost();
        var previousSharedFolderFeatureEnabledByHost = IsSharedFolderFeatureEnabledByHost();

        _config.Usb ??= new GuestUsbSettings();
        _config.SharedFolders ??= new GuestSharedFolderSettings();

        _config.Usb.HostFeatureEnabled = usbFeatureEnabledByHost;
        _config.SharedFolders.HostFeatureEnabled = sharedFoldersFeatureEnabledByHost;

        if (!string.IsNullOrWhiteSpace(hostName))
        {
            _config.Usb.HostName = hostName.Trim();
            _config.Usb.HostAddress = hostName.Trim();
        }

        UpdateHostDiscoveryPresentation();
        UpdateUsbRuntimeStatusUi();
        UpdateSharedFolderFeatureUi();

        if (!usbFeatureEnabledByHost)
        {
            _usbHostFeatureReactivationPending = false;
        }
        else if (!previousUsbFeatureEnabledByHost)
        {
            var token = ++_usbHostFeatureReactivationToken;
            _usbHostFeatureReactivationPending = true;
            UpdateUsbRuntimeStatusUi();
            _ = CompleteUsbHostFeatureReactivationAsync(token);
        }

        if (!sharedFoldersFeatureEnabledByHost)
        {
            _sharedFolderHostFeatureReactivationPending = false;
        }
        else if (!previousSharedFolderFeatureEnabledByHost)
        {
            var token = ++_sharedFolderHostFeatureReactivationToken;
            _sharedFolderHostFeatureReactivationPending = true;
            UpdateSharedFolderFeatureUi();
            _ = CompleteSharedFolderHostFeatureReactivationAsync(token);
        }

        if (!sharedFoldersFeatureEnabledByHost)
        {
            ScheduleSharedFolderAutoApply();
        }
    }

    private async Task SaveSettingsAsync()
    {
        _config.Ui.Theme = (_themeCombo.SelectedItem as string) ?? "dark";
        _config.Ui.StartWithWindows = _startWithWindowsCheckBox.IsChecked == true;
        _config.Ui.StartMinimized = _startMinimizedCheckBox.IsChecked == true;
        _config.Ui.MinimizeToTray = _minimizeToTrayCheckBox.IsChecked == true;
        _config.Ui.CheckForUpdatesOnStartup = _checkForUpdatesOnStartupCheckBox.IsChecked != false;
        _config.Usb ??= new GuestUsbSettings();
        _config.Usb.Enabled = _usbFeatureEnabledToggleSwitch.IsOn;
        _config.Usb.DisconnectOnExit = _usbDisconnectOnExitCheckBox.IsChecked != false;
        _config.Usb.UseHyperVSocket = _useHyperVSocketCheckBox.IsChecked != false;
        _config.Usb.HostAddress = (_usbHostAddressTextBox.Text ?? string.Empty).Trim();

        _config.FileService ??= new GuestFileServiceSettings();
        _config.FileService.Enabled = true;
        _config.FileService.MappingMode = GuestConfigService.NormalizeMappingMode(_config.FileService.MappingMode);

        await _saveConfigAsync(_config);
        ApplyTheme(_config.Ui.Theme);
        UpdateUsbTransportModePresentation();
        UpdateHostDiscoveryPresentation();

        AppendNotification("[Info] Einstellungen gespeichert.");
    }

    private async Task OnUsbTransportModeToggledAsync()
    {
        if (_suppressUsbTransportToggleEvents)
        {
            return;
        }

        _config.Usb ??= new GuestUsbSettings();
        _config.Usb.UseHyperVSocket = _useHyperVSocketCheckBox.IsChecked != false;

        UpdateUsbTransportModePresentation();

        await _saveConfigAsync(_config);
        AppendNotification(_config.Usb.UseHyperVSocket
            ? "[Info] USB Transportmodus: Hyper-V Socket bevorzugt."
            : "[Info] USB Transportmodus: IP-Mode aktiv.");

        if (_config.Usb.UseHyperVSocket)
        {
            ScheduleUsbTransportAutoRefresh();
        }
        else
        {
            CancelPendingUsbTransportAutoRefresh();
        }
    }

    private void CancelPendingUsbTransportAutoRefresh()
    {
        if (_usbTransportAutoRefreshCts is null)
        {
            return;
        }

        try
        {
            _usbTransportAutoRefreshCts.Cancel();
        }
        catch
        {
        }

        _usbTransportAutoRefreshCts.Dispose();
        _usbTransportAutoRefreshCts = null;
    }

    private void ScheduleUsbTransportAutoRefresh()
    {
        CancelPendingUsbTransportAutoRefresh();

        var cts = new CancellationTokenSource();
        _usbTransportAutoRefreshCts = cts;
        _ = RunUsbTransportAutoRefreshAsync(cts.Token);
    }

    private async Task RunUsbTransportAutoRefreshAsync(CancellationToken cancellationToken)
    {
        try
        {
            await Task.Delay(1000, cancellationToken);

            if (cancellationToken.IsCancellationRequested)
            {
                return;
            }

            if (_useHyperVSocketCheckBox.IsChecked != true || _config.Usb?.UseHyperVSocket != true)
            {
                return;
            }

            AppendNotification("[Info] Auto-Refresh nach Hyper-V Socket Aktivierung …");
            await RefreshUsbAsync();
        }
        catch (TaskCanceledException)
        {
        }
        catch (Exception ex)
        {
            AppendNotification($"[Warn] Auto-Refresh nach Hyper-V Socket Aktivierung fehlgeschlagen: {ex.Message}");
        }
        finally
        {
            if (_usbTransportAutoRefreshCts is not null && _usbTransportAutoRefreshCts.Token == cancellationToken)
            {
                _usbTransportAutoRefreshCts.Dispose();
                _usbTransportAutoRefreshCts = null;
            }
        }
    }

    private void UpdateUsbTransportModePresentation()
    {
        var useHyperVSocket = _useHyperVSocketCheckBox.IsChecked != false;
        var hyperVSocketLive = string.Equals(_diagHyperVSocketText.Text, "Ja", StringComparison.OrdinalIgnoreCase)
                              && string.Equals(_diagRegistryServiceText.Text, "Ja", StringComparison.OrdinalIgnoreCase)
                              && !string.Equals(_diagFallbackText.Text, "Ja", StringComparison.OrdinalIgnoreCase);
        var ipModeActive = !useHyperVSocket || !hyperVSocketLive;

        _usbHostAddressEditorCard.Visibility = ipModeActive ? Visibility.Visible : Visibility.Collapsed;
        _usbHostAddressTextBox.IsEnabled = ipModeActive;
        if (_usbHostSearchButton is not null)
        {
            _usbHostSearchButton.IsEnabled = ipModeActive;
        }
        _usbHostSearchStatusText.Visibility = ipModeActive ? Visibility.Visible : Visibility.Collapsed;

        if (useHyperVSocket)
        {
            _usbModeHintText.Text = ipModeActive
                ? "Hyper-V Socket ist aktiviert, aktuell aber nicht aktiv. IP-Mode/Fallback nutzt die Host-Adresse unten."
                : "Hyper-V Socket ist aktiviert. Bei Verfügbarkeitsproblemen wird auf IP zurückgefallen.";
        }
        else
        {
            _usbModeHintText.Text = "IP-Mode ist aktiviert. Die Host-Adresse unten wird für USB Connect verwendet.";
        }

        UpdateUsbTransportHeaderStatus();
    }

    private void UpdateUsbTransportHeaderStatus()
    {
        var useHyperVSocket = _config.Usb?.UseHyperVSocket != false;
        var hyperVSocketReportedActive = string.Equals(_diagHyperVSocketText.Text, "Ja", StringComparison.OrdinalIgnoreCase);
        var registryServicePresent = string.Equals(_diagRegistryServiceText.Text, "Ja", StringComparison.OrdinalIgnoreCase);
        var fallbackActive = string.Equals(_diagFallbackText.Text, "Ja", StringComparison.OrdinalIgnoreCase);
        var hyperVSocketLive = hyperVSocketReportedActive && registryServicePresent && !fallbackActive;

        if (useHyperVSocket)
        {
            _usbTransportModeBadgeText.Text = _isCompactHeaderLayout
                ? (hyperVSocketLive
                    ? "Hyper-V aktiv"
                    : (fallbackActive ? "IP-Fallback" : "Hyper-V bevorzugt"))
                : (hyperVSocketLive
                    ? "Hyper-Socket aktiv"
                    : (fallbackActive ? "IP-Fallback aktiv" : "Hyper-Socket bevorzugt"));
            var palette = fallbackActive
                ? ResolveUsbModePalette(forHyperV: false, isActive: true)
                : ResolveUsbModePalette(forHyperV: true, isActive: hyperVSocketLive);
            _usbTransportModeBadge.Background = palette.chipBackground;
            _usbTransportModeBadge.BorderBrush = palette.chipBorder;
            _usbTransportModeBadgeText.Foreground = palette.textForeground;
            UpdateUsbTransportModeBadges(useHyperVSocket: true, hyperVSocketLive: hyperVSocketLive, fallbackActive: fallbackActive);
            return;
        }

        var configuredHost = (_usbHostAddressTextBox.Text ?? _config.Usb?.HostAddress ?? string.Empty).Trim();
        var ipDisplay = string.IsNullOrWhiteSpace(configuredHost) ? "auto" : configuredHost;

        _usbTransportModeBadgeText.Text = _isCompactHeaderLayout
            ? "IP-Mode"
            : $"IP-Mode: {ipDisplay}";
        var ipPalette = ResolveUsbModePalette(forHyperV: false, isActive: true);
        _usbTransportModeBadge.Background = ipPalette.chipBackground;
        _usbTransportModeBadge.BorderBrush = ipPalette.chipBorder;
        _usbTransportModeBadgeText.Foreground = ipPalette.textForeground;
        UpdateUsbTransportModeBadges(useHyperVSocket: false, hyperVSocketLive: false, fallbackActive: true);
    }

    private void UpdateUsbTransportModeBadges(bool useHyperVSocket, bool hyperVSocketLive, bool fallbackActive)
    {
        var hyperVPalette = ResolveUsbModePalette(forHyperV: true, isActive: useHyperVSocket && hyperVSocketLive);
        _usbHyperVModeBadge.Background = hyperVPalette.chipBackground;
        _usbHyperVModeBadge.BorderBrush = hyperVPalette.chipBorder;
        _usbHyperVModeIconBadge.Background = hyperVPalette.iconBackground;
        _usbHyperVModeIconBadge.BorderBrush = hyperVPalette.iconBorder;
        _usbHyperVModeIcon.Foreground = hyperVPalette.iconForeground;
        _usbHyperVModeBadgeText.Foreground = hyperVPalette.textForeground;

        var ipPalette = ResolveUsbModePalette(forHyperV: false, isActive: !useHyperVSocket || fallbackActive);
        _usbIpModeBadge.Background = ipPalette.chipBackground;
        _usbIpModeBadge.BorderBrush = ipPalette.chipBorder;
        _usbIpModeIconBadge.Background = ipPalette.iconBackground;
        _usbIpModeIconBadge.BorderBrush = ipPalette.iconBorder;
        _usbIpModeIcon.Foreground = ipPalette.iconForeground;
        _usbIpModeBadgeText.Foreground = ipPalette.textForeground;
    }

    private (Brush chipBackground, Brush chipBorder, Brush iconBackground, Brush iconBorder, Brush iconForeground, Brush textForeground) ResolveUsbModePalette(bool forHyperV, bool isActive)
    {
        static SolidColorBrush Brush(byte a, byte r, byte g, byte b) => new(Color.FromArgb(a, r, g, b));

        var isDarkMode = string.Equals(CurrentTheme, "dark", StringComparison.OrdinalIgnoreCase);

        if (!isActive)
        {
            return (
                Application.Current.Resources["SurfaceSoftBrush"] as Brush ?? Brush(0xFF, 0x20, 0x2A, 0x48),
                Application.Current.Resources["PanelBorderBrush"] as Brush ?? Brush(0xFF, 0x44, 0x57, 0x7F),
                Application.Current.Resources["PanelBackgroundBrush"] as Brush ?? Brush(0xFF, 0x18, 0x23, 0x3E),
                Application.Current.Resources["PanelBorderBrush"] as Brush ?? Brush(0xFF, 0x44, 0x57, 0x7F),
                Application.Current.Resources["TextMutedBrush"] as Brush ?? Brush(0xFF, 0xA6, 0xB9, 0xD8),
                Application.Current.Resources["TextMutedBrush"] as Brush ?? Brush(0xFF, 0xA6, 0xB9, 0xD8));
        }

        if (forHyperV)
        {
            if (isDarkMode)
            {
                return (
                    Brush(0xFF, 0x14, 0x3C, 0x2C),
                    Brush(0xFF, 0x43, 0xB5, 0x81),
                    Brush(0xFF, 0x43, 0xB5, 0x81),
                    Brush(0xFF, 0x43, 0xB5, 0x81),
                    Brush(0xFF, 0x09, 0x2D, 0x1E),
                    Brush(0xFF, 0xD9, 0xF6, 0xE8));
            }

            return (
                Brush(0xFF, 0xE8, 0xF8, 0xEF),
                Brush(0xFF, 0x2F, 0x9E, 0x68),
                Brush(0xFF, 0x2F, 0x9E, 0x68),
                Brush(0xFF, 0x2F, 0x9E, 0x68),
                Brush(0xFF, 0xF7, 0xFF, 0xFB),
                Brush(0xFF, 0x0E, 0x4F, 0x31));
        }

        if (isDarkMode)
        {
            return (
                Brush(0xFF, 0x47, 0x31, 0x1B),
                Brush(0xFF, 0xF2, 0x9A, 0x3A),
                Brush(0xFF, 0xF2, 0x9A, 0x3A),
                Brush(0xFF, 0xF2, 0x9A, 0x3A),
                Brush(0xFF, 0x2A, 0x1A, 0x08),
                Brush(0xFF, 0xFF, 0xE9, 0xCC));
        }

        return (
            Brush(0xFF, 0xFF, 0xF1, 0xDF),
            Brush(0xFF, 0xD7, 0x82, 0x2C),
            Brush(0xFF, 0xD7, 0x82, 0x2C),
            Brush(0xFF, 0xD7, 0x82, 0x2C),
            Brush(0xFF, 0xFF, 0xFA, 0xF3),
            Brush(0xFF, 0x6B, 0x3A, 0x0A));
    }

    private (Brush chipBackground, Brush chipBorder, Brush textForeground) ResolveHostFeatureChipPalette(bool isActive)
    {
        var palette = ResolveUsbModePalette(forHyperV: true, isActive: isActive);
        return (palette.chipBackground, palette.chipBorder, palette.textForeground);
    }

    private static string BuildAutoConnectKey(UsbIpDeviceInfo device)
    {
        if (!string.IsNullOrWhiteSpace(device.HardwareId))
        {
            return "hardware:" + device.HardwareId.Trim();
        }

        if (!string.IsNullOrWhiteSpace(device.Description))
        {
            return "description:" + device.Description.Trim();
        }

        if (!string.IsNullOrWhiteSpace(device.BusId))
        {
            return "busid:" + device.BusId.Trim();
        }

        return string.Empty;
    }

    private async Task SetUsbDisconnectOnExitAsync(bool enabled)
    {
        if (_suppressUsbDisconnectOnExitToggleEvents)
        {
            return;
        }

        _config.Usb ??= new GuestUsbSettings();
        if (_config.Usb.DisconnectOnExit == enabled)
        {
            return;
        }

        _config.Usb.DisconnectOnExit = enabled;
        await _saveConfigAsync(_config);
        AppendNotification(enabled
            ? "[Info] USB-Disconnect beim Beenden aktiviert."
            : "[Info] USB-Disconnect beim Beenden deaktiviert.");
    }

    private void UpdateAutoConnectToggleFromSelection()
    {
        var selected = GetSelectedUsbDevice();
        var keys = _config.Usb?.AutoConnectDeviceKeys ?? [];
        var key = selected is null ? string.Empty : BuildAutoConnectKey(selected);

        _suppressUsbAutoConnectToggleEvents = true;
        try
        {
            _usbAutoConnectCheckBox.IsEnabled = IsUsbFeatureUsable() && selected is not null;
            _usbAutoConnectCheckBox.IsChecked = selected is not null
                && !string.IsNullOrWhiteSpace(key)
                && keys.Contains(key, StringComparer.OrdinalIgnoreCase);
        }
        finally
        {
            _suppressUsbAutoConnectToggleEvents = false;
        }
    }

    private async Task SetSelectedUsbDeviceAutoConnectAsync(bool enabled)
    {
        if (_suppressUsbAutoConnectToggleEvents)
        {
            return;
        }

        var selected = GetSelectedUsbDevice();
        if (selected is null)
        {
            return;
        }

        var key = BuildAutoConnectKey(selected);
        if (string.IsNullOrWhiteSpace(key))
        {
            AppendNotification("[Warn] Auto-Connect konnte für dieses Gerät nicht gespeichert werden.");
            UpdateAutoConnectToggleFromSelection();
            return;
        }

        _config.Usb ??= new GuestUsbSettings();
        var keys = _config.Usb.AutoConnectDeviceKeys ?? [];
        var changed = false;

        if (enabled)
        {
            if (!keys.Contains(key, StringComparer.OrdinalIgnoreCase))
            {
                keys.Add(key);
                changed = true;
            }
        }
        else
        {
            changed = keys.RemoveAll(existing => string.Equals(existing, key, StringComparison.OrdinalIgnoreCase)) > 0;
        }

        if (!changed)
        {
            return;
        }

        _config.Usb.AutoConnectDeviceKeys = keys
            .Where(static entry => !string.IsNullOrWhiteSpace(entry))
            .Select(static entry => entry.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        await _saveConfigAsync(_config);
        AppendNotification(enabled
            ? $"[Info] Auto-Connect aktiviert für: {selected.Description}"
            : $"[Info] Auto-Connect deaktiviert für: {selected.Description}");
    }

    public async Task CheckForUpdatesOnStartupIfEnabledAsync()
    {
        if (_config.Ui.CheckForUpdatesOnStartup != true)
        {
            return;
        }

        await CheckForUpdatesAsync();
    }

    private async Task ApplyThemeAndRestartImmediatelyAsync()
    {
        if (_isThemeRestartInProgress)
        {
            return;
        }

        var selectedTheme = GuestConfigService.NormalizeTheme((_themeCombo.SelectedItem as string) ?? _config.Ui.Theme);
        var currentTheme = GuestConfigService.NormalizeTheme(_config.Ui.Theme);
        if (string.Equals(selectedTheme, currentTheme, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        _isThemeRestartInProgress = true;
        try
        {
            _config.Ui.Theme = selectedTheme;
            _themeText.Text = selectedTheme == "dark" ? "Dunkles Theme" : "Helles Theme";

            AppendNotification("[Info] Theme geändert – speichere und starte HyperTool Guest neu …");
            await _saveConfigAsync(_config);
            await _restartForThemeChangeAsync(selectedTheme);
        }
        finally
        {
            _isThemeRestartInProgress = false;
        }
    }

    private async Task RefreshUsbAsync()
    {
        await RefreshHostFeatureAvailabilityFromSocketAsync();

        if (!IsUsbFeatureLocallyEnabled())
        {
            AppendNotification("[Info] USB Connect ist lokal deaktiviert.");
            return;
        }

        if (!IsUsbFeatureEnabledByHost())
        {
            AppendNotification("[Info] USB Connect ist durch den Host deaktiviert.");
            return;
        }

        try
        {
            var list = await _refreshUsbDevicesAsync();
            UpdateUsbDevices(list);
            AppendNotification($"[Info] {list.Count} USB-Gerät(e) geladen.");
        }
        catch (Exception ex)
        {
            AppendNotification($"[Error] USB Refresh fehlgeschlagen: {ex.Message}");
        }
    }

    private async Task SearchUsbHostAddressAsync()
    {
        if (_discoverUsbHostAddressAsync is null)
        {
            return;
        }

        SetUsbHostSearchStatus("Suche läuft …", UsbHostSearchStatusKind.Running);

        if (_usbHostSearchButton is not null)
        {
            _usbHostSearchButton.IsEnabled = false;
        }

        try
        {
            AppendNotification("[Info] Suche Hostname (Hyper-V Socket), sonst IP-Fallback …");

            var discoveredTarget = (await _discoverUsbHostAddressAsync() ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(discoveredTarget))
            {
                SetUsbHostSearchStatus("Kein Host gefunden", UsbHostSearchStatusKind.Error);
                AppendNotification("[Warn] Kein Hostname/IP gefunden. Stelle sicher, dass die Host-App läuft.");
                return;
            }

            _usbHostAddressTextBox.Text = discoveredTarget;
            _config.Usb ??= new GuestUsbSettings();
            _config.Usb.HostAddress = discoveredTarget;
            await _saveConfigAsync(_config);

            UpdateUsbTransportModePresentation();
            UpdateHostDiscoveryPresentation();

            var discoveredHostName = (_config.Usb.HostName ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(discoveredHostName))
            {
                SetUsbHostSearchStatus($"Hostname gefunden: {discoveredHostName}", UsbHostSearchStatusKind.Success);
                AppendNotification($"[Info] Hostname gefunden: {discoveredHostName}");
            }
            else
            {
                SetUsbHostSearchStatus($"IP-Fallback gefunden: {discoveredTarget}", UsbHostSearchStatusKind.Success);
                AppendNotification($"[Info] Kein Hostname gefunden, IP-Fallback: {discoveredTarget}");
            }
        }
        catch (Exception ex)
        {
            SetUsbHostSearchStatus("Suche fehlgeschlagen", UsbHostSearchStatusKind.Error);
            AppendNotification($"[Warn] Host-Suche fehlgeschlagen: {ex.Message}");
        }
        finally
        {
            if (_usbHostSearchButton is not null)
            {
                _usbHostSearchButton.IsEnabled = _usbHostAddressEditorCard.Visibility == Visibility.Visible;
            }
        }
    }

    private void SetUsbHostSearchStatus(string text, UsbHostSearchStatusKind statusKind)
    {
        _usbHostSearchStatusKind = statusKind;
        _usbHostSearchStatusText.Text = text;
        ApplyUsbHostSearchStatusColor();
    }

    private void ApplyUsbHostSearchStatusColor()
    {
        Brush? statusBrush = _usbHostSearchStatusKind switch
        {
            UsbHostSearchStatusKind.Success => Application.Current.Resources["SystemFillColorSuccessBrush"] as Brush,
            UsbHostSearchStatusKind.Error => Application.Current.Resources["SystemFillColorCriticalBrush"] as Brush,
            UsbHostSearchStatusKind.Running => Application.Current.Resources["AccentStrongBrush"] as Brush,
            _ => Application.Current.Resources["TextMutedBrush"] as Brush
        };

        _usbHostSearchStatusText.Foreground = statusBrush
            ?? Application.Current.Resources["TextMutedBrush"] as Brush;
    }

    private async Task ConnectUsbAsync()
    {
        if (!IsUsbFeatureLocallyEnabled())
        {
            AppendNotification("[Warn] USB Connect ist lokal deaktiviert.");
            return;
        }

        if (!IsUsbFeatureEnabledByHost())
        {
            AppendNotification("[Warn] USB Connect ist durch den Host deaktiviert.");
            return;
        }

        var selected = GetSelectedUsbDevice();
        if (selected is null)
        {
            AppendNotification("[Warn] Bitte ein USB-Gerät auswählen.");
            return;
        }

        var code = await _connectUsbAsync(selected.BusId);
        AppendNotification(code == 0
            ? "[Info] USB Host-Attach erfolgreich."
            : $"[Error] USB Host-Attach fehlgeschlagen (Code {code}).");

        await RefreshUsbAsync();
    }

    private async Task DisconnectUsbAsync()
    {
        if (!IsUsbFeatureLocallyEnabled())
        {
            AppendNotification("[Warn] USB Connect ist lokal deaktiviert.");
            return;
        }

        if (!IsUsbFeatureEnabledByHost())
        {
            AppendNotification("[Warn] USB Connect ist durch den Host deaktiviert.");
            return;
        }

        var selected = GetSelectedUsbDevice();
        if (selected is null)
        {
            AppendNotification("[Warn] Bitte ein USB-Gerät auswählen.");
            return;
        }

        var code = await _disconnectUsbAsync(selected.BusId);
        AppendNotification(code == 0
            ? "[Info] USB Host-Detach erfolgreich."
            : $"[Error] USB Host-Detach fehlgeschlagen (Code {code}).");

        if (code == 0)
        {
            AppendNotification("[Info] Aktualisiere USB-Liste in 3 Sekunden …");
            await Task.Delay(TimeSpan.FromSeconds(3));
        }

        await RefreshUsbAsync();
    }

    private void RunLogoEasterEgg()
    {
        _logoRotateTransform.Angle = 0;

        var storyboard = new Storyboard();
        var animation = new DoubleAnimation
        {
            From = 0,
            To = 720,
            Duration = TimeSpan.FromMilliseconds(1400),
            EnableDependentAnimation = true
        };

        Storyboard.SetTarget(animation, _logoRotateTransform);
        Storyboard.SetTargetProperty(animation, nameof(RotateTransform.Angle));
        storyboard.Children.Add(animation);
        storyboard.Begin();

        try
        {
            var soundPath = IOPath.Combine(AppContext.BaseDirectory, "Assets", "logo-spin.wav");
            if (File.Exists(soundPath))
            {
                _logoSpinPlayer?.Dispose();

                var player = new MediaPlayer
                {
                    AudioCategory = MediaPlayerAudioCategory.SoundEffects,
                    Volume = 0.30,
                    Source = MediaSource.CreateFromUri(new Uri(soundPath))
                };

                player.MediaEnded += (_, _) =>
                {
                    try
                    {
                        player.Dispose();
                    }
                    catch
                    {
                    }

                    if (ReferenceEquals(_logoSpinPlayer, player))
                    {
                        _logoSpinPlayer = null;
                    }
                };

                _logoSpinPlayer = player;
                player.Play();
            }
            else
            {
                SystemSounds.Asterisk.Play();
            }
        }
        catch
        {
            SystemSounds.Asterisk.Play();
        }
    }

    private void OpenHelpWindow()
    {
        if (_helpWindow is not null)
        {
            try
            {
                _helpWindow.Activate();
                return;
            }
            catch (Exception ex)
            {
                GuestLogger.Warn("help.reopen.failed", "Vorhandenes Hilfe-Fenster konnte nicht erneut aktiviert werden. Es wird neu erstellt.", new
                {
                    exceptionType = ex.GetType().FullName,
                    ex.Message
                });

                try
                {
                    _helpWindow.Close();
                }
                catch
                {
                }

                _helpWindow = null;
            }
        }

        var repoUrl = "https://github.com/koerby/HyperTool";
        _helpWindow = new HelpWindow(GuestConfigService.DefaultConfigPath, repoUrl, CurrentTheme);
        _helpWindow.Closed += (_, _) => _helpWindow = null;
        _helpWindow.Activate();
    }

    public void CloseAuxiliaryWindows()
    {
        try
        {
            _helpWindow?.Close();
        }
        catch
        {
        }
        finally
        {
            _helpWindow = null;
        }
    }

    private void OpenReleasePage()
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = string.IsNullOrWhiteSpace(_releaseUrl) ? "https://github.com/koerby/HyperTool/releases" : _releaseUrl,
            UseShellExecute = true
        });
    }

    private void OpenUsbipClientRepository()
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = "https://github.com/vadimgrn/usbip-win2",
            UseShellExecute = true
        });
    }

    private void OpenWinFspRepository()
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = $"https://github.com/{WinFspRuntimeOwner}/{WinFspRuntimeRepo}",
            UseShellExecute = true
        });
    }

    private async Task InstallGuestUsbRuntimeAsync()
    {
        if (_usbRuntimeInstallButton is not null)
        {
            _usbRuntimeInstallButton.IsEnabled = false;
        }

        try
        {
            AppendNotification("[Info] usbip-win2 Installer wird vorbereitet...");

            var installerResult = await _updateService.CheckForUpdateAsync(
                GuestUsbRuntimeOwner,
                GuestUsbRuntimeRepo,
                "0.0.0",
                CancellationToken.None,
                GuestUsbRuntimeAssetHint);

            if (!installerResult.Success || string.IsNullOrWhiteSpace(installerResult.InstallerDownloadUrl))
            {
                AppendNotification("[Warn] Installer-Asset konnte nicht automatisch ermittelt werden. Release-Seite wird geöffnet.");
                Process.Start(new ProcessStartInfo
                {
                    FileName = "https://github.com/vadimgrn/usbip-win2/releases/latest",
                    UseShellExecute = true
                });
                return;
            }

            var targetDirectory = IOPath.Combine(IOPath.GetTempPath(), "HyperTool", "runtime-installers");
            Directory.CreateDirectory(targetDirectory);

            var fileName = ResolveInstallerFileName(
                installerResult.InstallerDownloadUrl,
                installerResult.InstallerFileName,
                "usbip-win2-x64.exe");

            var installerPath = IOPath.Combine(targetDirectory, fileName);

            AppendNotification($"[Info] Lade usbip-win2 herunter: {fileName}");
            using (var response = await UpdateDownloadClient.GetAsync(installerResult.InstallerDownloadUrl, CancellationToken.None))
            {
                response.EnsureSuccessStatusCode();
                await using var stream = new FileStream(installerPath, FileMode.Create, FileAccess.Write, FileShare.None);
                await response.Content.CopyToAsync(stream);
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = installerPath,
                UseShellExecute = true,
                Verb = "runas"
            });

            AppendNotification("[Success] usbip-win2 Installer gestartet. Nach Abschluss App neu starten.");
        }
        catch (Exception ex)
        {
            AppendNotification($"[Error] Automatische usbip-win2 Installation fehlgeschlagen: {ex.Message}");
        }
        finally
        {
            if (_usbRuntimeInstallButton is not null)
            {
                _usbRuntimeInstallButton.IsEnabled = true;
            }
        }
    }

    private async Task InstallGuestWinFspRuntimeAsync()
    {
        if (_winFspRuntimeInstallButton is not null)
        {
            _winFspRuntimeInstallButton.IsEnabled = false;
        }

        try
        {
            AppendNotification("[Info] WinFsp Installer wird vorbereitet...");

            var installerResult = await _updateService.CheckForUpdateAsync(
                WinFspRuntimeOwner,
                WinFspRuntimeRepo,
                "0.0.0",
                CancellationToken.None,
                WinFspRuntimeAssetHint);

            if (!installerResult.Success || string.IsNullOrWhiteSpace(installerResult.InstallerDownloadUrl))
            {
                AppendNotification("[Warn] WinFsp Installer konnte nicht automatisch ermittelt werden. Release-Seite wird geöffnet.");
                Process.Start(new ProcessStartInfo
                {
                    FileName = "https://github.com/winfsp/winfsp/releases/latest",
                    UseShellExecute = true
                });
                return;
            }

            var targetDirectory = IOPath.Combine(IOPath.GetTempPath(), "HyperTool", "runtime-installers");
            Directory.CreateDirectory(targetDirectory);

            var fileName = ResolveInstallerFileName(
                installerResult.InstallerDownloadUrl,
                installerResult.InstallerFileName,
                "winfsp-x64.msi");

            var installerPath = IOPath.Combine(targetDirectory, fileName);

            AppendNotification($"[Info] Lade WinFsp herunter: {fileName}");
            using (var response = await UpdateDownloadClient.GetAsync(installerResult.InstallerDownloadUrl, CancellationToken.None))
            {
                response.EnsureSuccessStatusCode();
                await using var stream = new FileStream(installerPath, FileMode.Create, FileAccess.Write, FileShare.None);
                await response.Content.CopyToAsync(stream);
            }

            var extension = IOPath.GetExtension(installerPath);
            if (string.Equals(extension, ".msi", StringComparison.OrdinalIgnoreCase))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "msiexec.exe",
                    Arguments = $"/i \"{installerPath}\" /passive",
                    UseShellExecute = true,
                    Verb = "runas"
                });
            }
            else
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = installerPath,
                    UseShellExecute = true,
                    Verb = "runas"
                });
            }

            AppendNotification("[Success] WinFsp Installer gestartet. Nach Abschluss HyperTool Guest neu starten.");
        }
        catch (Exception ex)
        {
            AppendNotification($"[Error] Automatische WinFsp Installation fehlgeschlagen: {ex.Message}");
        }
        finally
        {
            UpdateWinFspRuntimeStatusUi();
            if (_winFspRuntimeInstallButton is not null)
            {
                _winFspRuntimeInstallButton.IsEnabled = true;
            }
        }
    }

    private void OpenLogFile()
    {
        var logDirectory = string.IsNullOrWhiteSpace(_config.Logging.DirectoryPath)
            ? GuestConfigService.DefaultLogDirectory
            : _config.Logging.DirectoryPath;

        Directory.CreateDirectory(logDirectory);
        Process.Start(new ProcessStartInfo
        {
            FileName = logDirectory,
            UseShellExecute = true
        });
    }

    private void CopyNotificationsToClipboard()
    {
        var text = _notifications.Count == 0
            ? "Keine Notifications vorhanden."
            : string.Join(Environment.NewLine, _notifications.Select(entry => entry ?? string.Empty));

        var package = new DataPackage();
        package.SetText(text);
        Clipboard.SetContent(package);
        AppendNotification("[Info] Notifications in Zwischenablage kopiert.");
    }

    private async Task CheckForUpdatesAsync()
    {
        AppendNotification("[Info] Update-Prüfung gestartet.");

        var currentVersion = ResolveGuestVersionText();
        var result = await _updateService.CheckForUpdateAsync(
            UpdateOwner,
            UpdateRepo,
            currentVersion,
            CancellationToken.None,
            GuestInstallerAssetHint);

        _updateStatusValueText.Text = result.Message;
        _releaseUrl = string.IsNullOrWhiteSpace(result.ReleaseUrl) ? _releaseUrl : result.ReleaseUrl;
        _updateCheckSucceeded = result.Success;
        _updateAvailable = result.Success && result.HasUpdate;

        if (_updateAvailable)
        {
            _installerDownloadUrl = result.InstallerDownloadUrl ?? string.Empty;
            _installerFileName = result.InstallerFileName ?? string.Empty;
        }
        else
        {
            _installerDownloadUrl = string.Empty;
            _installerFileName = string.Empty;
        }

        if (_installUpdateButton is not null)
        {
            _installUpdateButton.IsEnabled = CanInstallUpdate();
        }

        if (!result.Success)
        {
            AppendNotification($"[Warn] {result.Message}");
            return;
        }

        if (result.HasUpdate)
        {
            if (string.IsNullOrWhiteSpace(_installerDownloadUrl))
            {
                AppendNotification("[Warn] Update gefunden, aber kein Guest-Installer-Asset erkannt.");
            }
            else
            {
                AppendNotification($"[Info] {result.Message}");
            }
        }
        else
        {
            AppendNotification("[Info] Bereits aktuell.");
        }
    }

    private bool CanInstallUpdate()
    {
        return _updateCheckSucceeded
            && _updateAvailable
            && !string.IsNullOrWhiteSpace(_installerDownloadUrl);
    }

    private async Task InstallUpdateAsync()
    {
        if (string.IsNullOrWhiteSpace(_installerDownloadUrl))
        {
            AppendNotification("[Warn] Kein Installer-Download verfügbar. Bitte zuerst Update prüfen.");
            return;
        }

        try
        {
            AppendNotification("[Info] Lade Update-Installer herunter...");

            var targetDirectory = IOPath.Combine(IOPath.GetTempPath(), "HyperTool", "updates");
            Directory.CreateDirectory(targetDirectory);

            var fileName = ResolveInstallerFileName(_installerDownloadUrl, _installerFileName);
            var installerPath = IOPath.Combine(targetDirectory, fileName);

            using var response = await UpdateDownloadClient.GetAsync(_installerDownloadUrl, CancellationToken.None);
            response.EnsureSuccessStatusCode();

            await using (var stream = new FileStream(installerPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                await response.Content.CopyToAsync(stream);
            }

            AppendNotification($"[Info] Installer gespeichert: {installerPath}");

            if (!StartInstallerDetached(installerPath))
            {
                throw new InvalidOperationException("Installer konnte nicht gestartet werden.");
            }

            AppendNotification("[Info] Installer gestartet. HyperTool Guest wird jetzt beendet.");
            await Task.Delay(500);
            await _exitForUpdateInstallAsync();
        }
        catch (Exception ex)
        {
            AppendNotification($"[Error] Update-Installation fehlgeschlagen: {ex.Message}");
        }
    }

    private static bool StartInstallerDetached(string installerPath)
    {
        if (string.IsNullOrWhiteSpace(installerPath) || !File.Exists(installerPath))
        {
            return false;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/c start \"\" \"{installerPath}\"",
                UseShellExecute = false,
                CreateNoWindow = true
            });
            return true;
        }
        catch
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = installerPath,
                    UseShellExecute = true
                });
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

    private void UpdateUsbRuntimeStatusUi()
    {
        var hostEnabled = IsUsbFeatureEnabledByHost();
        var localEnabled = IsUsbFeatureLocallyEnabled();
        var isAvailable = _isUsbClientAvailable;

        if (!isAvailable && localEnabled)
        {
            _config.Usb ??= new GuestUsbSettings();
            _config.Usb.Enabled = false;
            localEnabled = false;
            _ = SaveConfigQuietlyAsync();
        }

        var featureUsable = IsUsbFeatureUsable() && !_usbHostFeatureReactivationPending;
        var refreshAvailable = IsUsbRefreshAvailable();

        _usbRuntimeStatusDot.Fill = new SolidColorBrush(isAvailable
            ? Windows.UI.Color.FromArgb(0xFF, 0x32, 0xD7, 0x4B)
            : Windows.UI.Color.FromArgb(0xFF, 0xE8, 0x4A, 0x5F));
        _usbRuntimeStatusText.Text = !isAvailable
            ? "USB/IP-Client: Nicht installiert"
            : (!localEnabled
                ? "USB-Connect: lokal deaktiviert"
                : (!hostEnabled
                    ? "USB-Connect: durch Host deaktiviert"
                    : "USB/IP-Client: Verfügbar"));

        _usbDisabledOverlayText.Visibility = featureUsable ? Visibility.Collapsed : Visibility.Visible;

        var usbEffectiveEnabled = IsUsbFeatureUsable();
        var usbChipPalette = ResolveHostFeatureChipPalette(usbEffectiveEnabled);
        _usbHostFeatureStatusChip.Background = usbChipPalette.chipBackground;
        _usbHostFeatureStatusChip.BorderBrush = usbChipPalette.chipBorder;
        _usbHostFeatureStatusChipText.Text = usbEffectiveEnabled ? "USB Aktiv" : "USB Inaktiv";
        _usbHostFeatureStatusChipText.Foreground = usbChipPalette.textForeground;

        _usbPageFeatureStatusChip.Background = usbChipPalette.chipBackground;
        _usbPageFeatureStatusChip.BorderBrush = usbChipPalette.chipBorder;
        _usbPageFeatureStatusChipText.Text = usbEffectiveEnabled ? "Aktiv" : "Inaktiv";
        _usbPageFeatureStatusChipText.Foreground = usbChipPalette.textForeground;

        _usbFeatureEnabledToggleSwitch.IsEnabled = hostEnabled && isAvailable;

        if (_usbRuntimeInstallButton is not null)
        {
            _usbRuntimeInstallButton.Visibility = !isAvailable ? Visibility.Visible : Visibility.Collapsed;
            _usbRuntimeInstallButton.IsEnabled = true;
        }

        if (_usbRuntimeRestartButton is not null)
        {
            _usbRuntimeRestartButton.Visibility = !isAvailable ? Visibility.Visible : Visibility.Collapsed;
            _usbRuntimeRestartButton.IsEnabled = true;
        }

        if (_usbRefreshButton is not null)
        {
            _usbRefreshButton.IsEnabled = refreshAvailable;
        }

        if (_usbConnectButton is not null)
        {
            _usbConnectButton.IsEnabled = featureUsable;
        }

        if (_usbDisconnectButton is not null)
        {
            _usbDisconnectButton.IsEnabled = featureUsable;
        }

        _usbAutoConnectCheckBox.IsEnabled = featureUsable;
        _usbDisconnectOnExitCheckBox.IsEnabled = featureUsable;
    }

    private async Task CompleteUsbHostFeatureReactivationAsync(int token)
    {
        try
        {
            await Task.Delay(TimeSpan.FromSeconds(1));
            if (token != _usbHostFeatureReactivationToken || !IsUsbFeatureEnabledByHost())
            {
                return;
            }

            await RefreshUsbAsync();
        }
        finally
        {
            if (token == _usbHostFeatureReactivationToken)
            {
                _usbHostFeatureReactivationPending = false;
                UpdateUsbRuntimeStatusUi();
            }
        }
    }

    private async Task CompleteSharedFolderHostFeatureReactivationAsync(int token)
    {
        try
        {
            await Task.Delay(TimeSpan.FromSeconds(1));
            if (token != _sharedFolderHostFeatureReactivationToken || !IsSharedFolderFeatureEnabledByHost())
            {
                return;
            }

            await SyncSharedFoldersFromHostAsync();
            await RefreshSharedFolderMountStatesSafeAsync();
        }
        finally
        {
            if (token == _sharedFolderHostFeatureReactivationToken)
            {
                _sharedFolderHostFeatureReactivationPending = false;
                UpdateSharedFolderFeatureUi();
            }
        }
    }

    private async Task RefreshHostFeatureAvailabilityFromSocketAsync()
    {
        try
        {
            var identity = await new HyperVSocketHostIdentityGuestClient().FetchHostIdentityAsync(CancellationToken.None);
            if (identity is null)
            {
                return;
            }

            UpdateHostFeatureAvailability(
                usbFeatureEnabledByHost: identity.Features?.UsbSharingEnabled != false,
                sharedFoldersFeatureEnabledByHost: identity.Features?.SharedFoldersEnabled != false,
                hostName: identity.HostName);
        }
        catch
        {
        }
    }

    private void UpdateWinFspRuntimeStatusUi()
    {
        var isAvailable = IsWinFspRuntimeInstalled();
        _winFspRuntimeStatusDot.Fill = new SolidColorBrush(isAvailable
            ? Windows.UI.Color.FromArgb(0xFF, 0x32, 0xD7, 0x4B)
            : Windows.UI.Color.FromArgb(0xFF, 0xE8, 0x4A, 0x5F));
        _winFspRuntimeStatusText.Text = isAvailable
            ? "WinFsp Runtime: Verfügbar"
            : "WinFsp Runtime: Nicht installiert";

        if (_winFspRuntimeInstallButton is not null)
        {
            _winFspRuntimeInstallButton.Visibility = isAvailable ? Visibility.Collapsed : Visibility.Visible;
        }

        if (_winFspRuntimeRestartButton is not null)
        {
            _winFspRuntimeRestartButton.Visibility = isAvailable ? Visibility.Collapsed : Visibility.Visible;
        }
    }

    private async Task RestartGuestToolAsync()
    {
        try
        {
            var currentTheme = GuestConfigService.NormalizeTheme((_themeCombo.SelectedItem as string) ?? _config.Ui.Theme);
            AppendNotification("[Info] Tool wird neu geladen …");
            await _restartForThemeChangeAsync(currentTheme);
        }
        catch (Exception ex)
        {
            AppendNotification($"[Error] Neustart fehlgeschlagen: {ex.Message}");
        }
    }

    private static bool IsWinFspRuntimeInstalled()
    {
        try
        {
            var installDir = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\WinFsp", "InstallDir", null) as string;
            if (!string.IsNullOrWhiteSpace(installDir)
                && File.Exists(IOPath.Combine(installDir, "bin", "winfsp-x64.dll")))
            {
                return true;
            }
        }
        catch
        {
        }

        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        if (!string.IsNullOrWhiteSpace(programFiles)
            && File.Exists(IOPath.Combine(programFiles, "WinFsp", "bin", "winfsp-x64.dll")))
        {
            return true;
        }

        var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        if (!string.IsNullOrWhiteSpace(programFilesX86)
            && File.Exists(IOPath.Combine(programFilesX86, "WinFsp", "bin", "winfsp-x64.dll")))
        {
            return true;
        }

        return false;
    }

    private static string ResolveInstallerFileName(string downloadUrl, string? fileName, string defaultFileName = "HyperTool-Guest-Setup.exe")
    {
        if (!string.IsNullOrWhiteSpace(fileName))
        {
            return fileName.Trim();
        }

        if (Uri.TryCreate(downloadUrl, UriKind.Absolute, out var uri))
        {
            var inferred = IOPath.GetFileName(uri.LocalPath);
            if (!string.IsNullOrWhiteSpace(inferred))
            {
                return inferred;
            }
        }

        return defaultFileName;
    }

    private void OnLoggerEntryWritten(string message)
    {
        if (!DispatcherQueue.TryEnqueue(() => AppendNotification(message)))
        {
            AppendNotification(message);
        }
    }

    private async Task SaveConfigQuietlyAsync()
    {
        try
        {
            await _saveConfigAsync(_config);
        }
        catch
        {
        }
    }

    private void AppendNotification(string message)
    {
        _statusText.Text = message;
        _notifications.Insert(0, message);

        while (_notifications.Count > 200)
        {
            _notifications.RemoveAt(_notifications.Count - 1);
        }

        UpdateBusyAndNotificationPanel();
    }

    private void UpdateBusyAndNotificationPanel()
    {
        _toggleLogButton.Content = _isLogExpanded ? "▾ Log einklappen" : "▸ Log ausklappen";
        _notificationExpandedGrid.Visibility = _isLogExpanded ? Visibility.Visible : Visibility.Collapsed;
        _notificationSummaryBorder.Visibility = _isLogExpanded ? Visibility.Collapsed : Visibility.Visible;
    }

    private string _configPathLabel() => $"Aktive Config: {GuestConfigService.DefaultConfigPath}";

    private Button CreateNavButton(string icon, string title, int index)
    {
        var button = new Button
        {
            Padding = new Thickness(8, 8, 8, 8),
            MinHeight = 78,
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush,
            HorizontalAlignment = HorizontalAlignment.Stretch
        };

        var content = new StackPanel { Spacing = 4, HorizontalAlignment = HorizontalAlignment.Center };
        var navIconHost = new Grid
        {
            Width = 36,
            Height = 36,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
        navIconHost.Children.Add(new Viewbox
        {
            Stretch = Stretch.Uniform,
            Margin = new Thickness(1),
            Child = new TextBlock
            {
                Text = icon,
                FontSize = 24,
                FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                TextAlignment = TextAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            }
        });
        content.Children.Add(navIconHost);
        content.Children.Add(new TextBlock
        {
            Text = title,
            FontSize = 11,
            Opacity = 0.9,
            TextWrapping = TextWrapping.Wrap,
            MaxWidth = 72,
            TextAlignment = TextAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center
        });

        button.Content = content;
        button.Click += async (_, _) =>
        {
            await SelectMenuIndexAsync(index);
        };

        _navButtons.Add(button);
        return button;
    }

    private static Button CreateIconButton(string icon, string label, RoutedEventHandler? onClick = null)
    {
        var content = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            VerticalAlignment = VerticalAlignment.Center
        };

        content.Children.Add(new TextBlock
        {
            Text = icon,
            FontSize = 18,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            VerticalAlignment = VerticalAlignment.Center
        });

        content.Children.Add(new TextBlock
        {
            Text = label,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            VerticalAlignment = VerticalAlignment.Center
        });

        var button = new Button
        {
            Content = content,
            Padding = new Thickness(10, 7, 10, 7),
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush
        };

        if (onClick is not null)
        {
            button.Click += onClick;
        }

        return button;
    }

    private static Button CreateSupportCoffeeButton(RoutedEventHandler onClick)
    {
        var textBrush = new SolidColorBrush(Color.FromArgb(0xFF, 0x2E, 0x20, 0x00));
        var content = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            VerticalAlignment = VerticalAlignment.Center
        };

        content.Children.Add(new TextBlock
        {
            Text = "☕",
            FontSize = 18,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            Foreground = textBrush,
            VerticalAlignment = VerticalAlignment.Center
        });

        content.Children.Add(new TextBlock
        {
            Text = "Buy Me a Coffee",
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            Foreground = textBrush,
            VerticalAlignment = VerticalAlignment.Center
        });

        var button = new Button
        {
            Content = content,
            Padding = new Thickness(10, 7, 10, 7),
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(Color.FromArgb(0xFF, 0xC9, 0x9B, 0x00)),
            Background = new SolidColorBrush(Color.FromArgb(0xFF, 0xFF, 0xDD, 0x00))
        };

        button.Click += onClick;
        return button;
    }

    private static Border CreateCard(Thickness margin, double padding, double cornerRadius)
    {
        var gradientBrush = new LinearGradientBrush
        {
            StartPoint = new Windows.Foundation.Point(0, 0),
            EndPoint = new Windows.Foundation.Point(1, 1)
        };

        var topBrush = Application.Current.Resources["SurfaceTopBrush"] as SolidColorBrush;
        var bottomBrush = Application.Current.Resources["SurfaceBottomBrush"] as SolidColorBrush;

        gradientBrush.GradientStops.Add(new GradientStop
        {
            Color = topBrush?.Color ?? Color.FromArgb(0xFA, 0xFF, 0xFF, 0xFF),
            Offset = 0.0
        });
        gradientBrush.GradientStops.Add(new GradientStop
        {
            Color = bottomBrush?.Color ?? Color.FromArgb(0xF0, 0xF2, 0xF8, 0xFF),
            Offset = 1.0
        });

        return new Border
        {
            Margin = margin,
            Padding = new Thickness(padding),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(cornerRadius),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = gradientBrush
        };
    }

    private static void ApplyThemePalette(bool isDark)
    {
        if (Application.Current?.Resources is not ResourceDictionary resources)
        {
            return;
        }

        SetBrushColor(resources, "PageBackgroundBrush", isDark ? Color.FromArgb(0xFF, 0x14, 0x1A, 0x31) : Color.FromArgb(0xFF, 0xF3, 0xF7, 0xFD));
        SetBrushColor(resources, "PanelBackgroundBrush", isDark ? Color.FromArgb(0xFF, 0x1A, 0x22, 0x40) : Color.FromArgb(0xFF, 0xFF, 0xFF, 0xFF));
        SetBrushColor(resources, "PanelBorderBrush", isDark ? Color.FromArgb(0xFF, 0x2A, 0x37, 0x60) : Color.FromArgb(0xFF, 0xC5, 0xD6, 0xEA));
        SetBrushColor(resources, "TextPrimaryBrush", isDark ? Color.FromArgb(0xFF, 0xE8, 0xF0, 0xFF) : Color.FromArgb(0xFF, 0x0C, 0x21, 0x38));
        SetBrushColor(resources, "TextMutedBrush", isDark ? Color.FromArgb(0xFF, 0x9E, 0xB4, 0xDA) : Color.FromArgb(0xFF, 0x2E, 0x4B, 0x69));
        SetBrushColor(resources, "AccentBrush", isDark ? Color.FromArgb(0xFF, 0x6C, 0xC9, 0xFF) : Color.FromArgb(0xFF, 0x1F, 0x79, 0xCC));
        SetBrushColor(resources, "AccentSoftBrush", isDark ? Color.FromArgb(0x3A, 0x7B, 0xC9, 0x66) : Color.FromArgb(0x26, 0x1F, 0x79, 0xCC));
        SetBrushColor(resources, "AccentStrongBrush", isDark ? Color.FromArgb(0xFF, 0x8B, 0xD4, 0xFF) : Color.FromArgb(0xFF, 0x1F, 0x79, 0xCC));
        SetBrushColor(resources, "SurfaceTopBrush", isDark ? Color.FromArgb(0xFF, 0x20, 0x29, 0x49) : Color.FromArgb(0xFF, 0xFF, 0xFF, 0xFF));
        SetBrushColor(resources, "SurfaceBottomBrush", isDark ? Color.FromArgb(0xFF, 0x17, 0x1F, 0x39) : Color.FromArgb(0xFF, 0xF6, 0xFA, 0xFF));
        SetBrushColor(resources, "SurfaceSoftBrush", isDark ? Color.FromArgb(0xFF, 0x1D, 0x26, 0x45) : Color.FromArgb(0xFF, 0xF1, 0xF7, 0xFF));
    }

    private static void SetBrushColor(ResourceDictionary resources, string key, Color color)
    {
        if (resources.TryGetValue(key, out var existingValue) && existingValue is SolidColorBrush existingBrush)
        {
            existingBrush.Color = color;
            return;
        }

        resources[key] = new SolidColorBrush(color);
    }

    private void UpdateTitleBarAppearance(bool isDark)
    {
        try
        {
            if (AppWindow?.TitleBar is not Microsoft.UI.Windowing.AppWindowTitleBar titleBar)
            {
                return;
            }

            if (isDark)
            {
                titleBar.BackgroundColor = Color.FromArgb(0xFF, 0x17, 0x1F, 0x3A);
                titleBar.ForegroundColor = Color.FromArgb(0xFF, 0xE8, 0xF0, 0xFF);
                titleBar.InactiveBackgroundColor = Color.FromArgb(0xFF, 0x14, 0x1A, 0x31);
                titleBar.InactiveForegroundColor = Color.FromArgb(0xFF, 0x98, 0xAE, 0xD3);

                titleBar.ButtonBackgroundColor = Color.FromArgb(0xFF, 0x17, 0x1F, 0x3A);
                titleBar.ButtonForegroundColor = Color.FromArgb(0xFF, 0xE8, 0xF0, 0xFF);
                titleBar.ButtonHoverBackgroundColor = Color.FromArgb(0xFF, 0x22, 0x2D, 0x51);
                titleBar.ButtonHoverForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xFF, 0xFF);
                titleBar.ButtonPressedBackgroundColor = Color.FromArgb(0xFF, 0x2A, 0x36, 0x61);
                titleBar.ButtonPressedForegroundColor = Color.FromArgb(0xFF, 0xFF, 0xFF, 0xFF);
                titleBar.ButtonInactiveBackgroundColor = Color.FromArgb(0xFF, 0x14, 0x1A, 0x31);
                titleBar.ButtonInactiveForegroundColor = Color.FromArgb(0xFF, 0x98, 0xAE, 0xD3);
            }
            else
            {
                titleBar.BackgroundColor = Color.FromArgb(0xFF, 0xD8, 0xE9, 0xFF);
                titleBar.ForegroundColor = Color.FromArgb(0xFF, 0x0F, 0x24, 0x3C);
                titleBar.InactiveBackgroundColor = Color.FromArgb(0xFF, 0xE6, 0xF2, 0xFF);
                titleBar.InactiveForegroundColor = Color.FromArgb(0xFF, 0x4E, 0x66, 0x83);

                titleBar.ButtonBackgroundColor = Color.FromArgb(0xFF, 0xD8, 0xE9, 0xFF);
                titleBar.ButtonForegroundColor = Color.FromArgb(0xFF, 0x0F, 0x24, 0x3C);
                titleBar.ButtonHoverBackgroundColor = Color.FromArgb(0xFF, 0xC7, 0xDE, 0xFC);
                titleBar.ButtonHoverForegroundColor = Color.FromArgb(0xFF, 0x0A, 0x1B, 0x30);
                titleBar.ButtonPressedBackgroundColor = Color.FromArgb(0xFF, 0xBA, 0xD3, 0xF7);
                titleBar.ButtonPressedForegroundColor = Color.FromArgb(0xFF, 0x08, 0x19, 0x2C);
                titleBar.ButtonInactiveBackgroundColor = Color.FromArgb(0xFF, 0xE6, 0xF2, 0xFF);
                titleBar.ButtonInactiveForegroundColor = Color.FromArgb(0xFF, 0x4E, 0x66, 0x83);
            }
        }
        catch
        {
        }
    }

    private void TryApplyWindowIcon()
    {
        try
        {
            if (AppWindow is null)
            {
                return;
            }

            var iconPath = new[]
            {
                IOPath.Combine(AppContext.BaseDirectory, "Assets", "HyperTool.Guest.ico"),
                IOPath.Combine(AppContext.BaseDirectory, "Assets", "HyperTool.ico"),
                IOPath.Combine(AppContext.BaseDirectory, "HyperTool.Guest.ico")
            }.FirstOrDefault(File.Exists);

            if (!string.IsNullOrWhiteSpace(iconPath))
            {
                AppWindow.SetIcon(iconPath);
            }
        }
        catch
        {
        }
    }
}
