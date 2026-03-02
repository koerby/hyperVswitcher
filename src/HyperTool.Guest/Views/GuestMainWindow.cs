using HyperTool.Models;
using HyperTool.Services;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Animation;
using Microsoft.UI.Xaml.Shapes;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Media;
using System.Net.Http;
using System.Reflection;
using Windows.ApplicationModel.DataTransfer;
using Windows.Graphics;
using Windows.UI;
using IOPath = System.IO.Path;

namespace HyperTool.Guest.Views;

internal sealed class GuestMainWindow : Window
{
    public const int DefaultWindowWidth = 1400;
    public const int DefaultWindowHeight = 940;
    private const string UpdateOwner = "koerby";
    private const string UpdateRepo = "HyperTool";
    private const string GuestInstallerAssetHint = "HyperTool-Guest-Setup";

    private readonly Func<Task<IReadOnlyList<UsbIpDeviceInfo>>> _refreshUsbDevicesAsync;
    private readonly Func<string, Task<int>> _connectUsbAsync;
    private readonly Func<string, Task<int>> _disconnectUsbAsync;
    private readonly Func<GuestConfig, Task> _saveConfigAsync;
    private readonly Func<string, Task> _restartForThemeChangeAsync;
    private readonly bool _isUsbClientAvailable;
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
    private readonly TextBlock _statusText = new() { Text = "Bereit.", TextWrapping = TextWrapping.Wrap };
    private readonly TextBlock _updateStatusValueText = new() { Text = "Noch nicht geprüft", TextWrapping = TextWrapping.Wrap, Opacity = 0.9 };
    private readonly Button _toggleLogButton = new();
    private bool _isLogExpanded;
    private string _releaseUrl = "https://github.com/koerby/HyperTool/releases";
    private string _installerDownloadUrl = string.Empty;
    private string _installerFileName = string.Empty;

    private readonly ListView _usbListView = new();
    private Button? _usbRefreshButton;
    private Button? _usbConnectButton;
    private Button? _usbDisconnectButton;
    private readonly ComboBox _themeCombo = new();
    private readonly ToggleSwitch _themeToggle = new();
    private readonly TextBlock _themeText = new();
    private readonly TextBox _usbHostAddressTextBox = new();
    private readonly CheckBox _startWithWindowsCheckBox = new() { Content = "Mit Windows starten" };
    private readonly CheckBox _startMinimizedCheckBox = new() { Content = "Beim Start minimiert" };
    private readonly CheckBox _minimizeToTrayCheckBox = new() { Content = "Tasktray-Menü aktiv" };

    private HelpWindow? _helpWindow;
    private int _selectedMenuIndex;
    private GuestConfig _config;
    private IReadOnlyList<UsbIpDeviceInfo> _usbDevices = [];
    private bool _suppressThemeEvents;
    private bool _isThemeRestartInProgress;
    private bool _isThemeToggleHandlerAttached;
    private List<UIElement>? _startupMainElements;
    private UIElement? _usbPage;
    private UIElement? _settingsPage;
    private UIElement? _infoPage;

    public GuestMainWindow(
        GuestConfig config,
        Func<Task<IReadOnlyList<UsbIpDeviceInfo>>> refreshUsbDevicesAsync,
        Func<string, Task<int>> connectUsbAsync,
        Func<string, Task<int>> disconnectUsbAsync,
        Func<GuestConfig, Task> saveConfigAsync,
        Func<string, Task> restartForThemeChangeAsync,
        bool isUsbClientAvailable)
    {
        _config = config;
        _refreshUsbDevicesAsync = refreshUsbDevicesAsync;
        _connectUsbAsync = connectUsbAsync;
        _disconnectUsbAsync = disconnectUsbAsync;
        _saveConfigAsync = saveConfigAsync;
        _restartForThemeChangeAsync = restartForThemeChangeAsync;
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

        GuestLogger.EntryWritten += OnLoggerEntryWritten;
        Closed += (_, _) => GuestLogger.EntryWritten -= OnLoggerEntryWritten;
    }

    public string CurrentTheme => GuestConfigService.NormalizeTheme((_themeCombo.SelectedItem as string) ?? _config.Ui.Theme);

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
        while (startupStart.ElapsedMilliseconds < HyperTool.WinUI.Views.LifecycleVisuals.SplashMinVisibleMs)
        {
            var status = HyperTool.WinUI.Views.LifecycleVisuals.StartupStatusMessages[
                statusIndex % HyperTool.WinUI.Views.LifecycleVisuals.StartupStatusMessages.Length];
            _overlayText.Text = status;

            _overlayProgressBar.Value = Math.Min(92, 8 + (statusIndex * 17));
            statusIndex++;

            var remaining = HyperTool.WinUI.Views.LifecycleVisuals.SplashMinVisibleMs - startupStart.ElapsedMilliseconds;
            var delay = (int)Math.Min(HyperTool.WinUI.Views.LifecycleVisuals.SplashStatusCycleMs, remaining);
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
        _usbDevices = devices;

        if (!_isUsbClientAvailable)
        {
            _usbListView.ItemsSource = new[]
            {
                "USB/IP-Client nicht installiert. USB-Funktionen sind deaktiviert."
            };
            return;
        }

        _usbListView.ItemsSource = devices.Select(item => item.DisplayName).ToList();
    }

    public UsbIpDeviceInfo? GetSelectedUsbDevice()
    {
        if (_usbListView.SelectedIndex < 0 || _usbListView.SelectedIndex >= _usbDevices.Count)
        {
            return null;
        }

        return _usbDevices[_usbListView.SelectedIndex];
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
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var titleStack = new StackPanel { Orientation = Orientation.Vertical, Spacing = 2 };
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
        titleStack.Children.Add(new TextBlock
        {
            Text = "dein nützlicher Hyper V Helfer",
            Opacity = 0.8,
            Margin = new Thickness(0, 0, 0, 6)
        });

        headerGrid.Children.Add(titleStack);

        var titleActions = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 10,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Top
        };

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
        titleActions.Children.Add(logoBorder);

        Grid.SetColumn(titleActions, 1);
        headerGrid.Children.Add(titleActions);

        return headerGrid;
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
        sidebarStack.Children.Add(CreateNavButton("⚙", "Einstellungen", 1));
        sidebarStack.Children.Add(CreateNavButton("ℹ", "Info", 2));
        sidebar.Child = sidebarStack;

        mainGrid.Children.Add(sidebar);

        var contentGrid = new Grid();
        contentGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        contentGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(10) });
        contentGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var topRow = new Border
        {
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            Padding = new Thickness(12),
            Child = new TextBlock
            {
                Text = "Guest USB Connect & Management",
                FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
            }
        };

        contentGrid.Children.Add(topRow);
        Grid.SetRow(_pageContent, 2);
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

        var topRow = new Grid { ColumnSpacing = 8 };
        topRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        topRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        topRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        topRow.Children.Add(new TextBlock { Text = "Notifications", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, VerticalAlignment = VerticalAlignment.Center });

        var summaryButtons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        summaryButtons.Children.Add(CreateIconButton("📄", "Logdatei öffnen", onClick: (_, _) => OpenLogFile()));
        _toggleLogButton.CornerRadius = new CornerRadius(10);
        _toggleLogButton.Padding = new Thickness(10, 7, 10, 7);
        _toggleLogButton.BorderThickness = new Thickness(1);
        _toggleLogButton.BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush;
        _toggleLogButton.Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush;
        _toggleLogButton.Click += (_, _) =>
        {
            _isLogExpanded = !_isLogExpanded;
            UpdateBusyAndNotificationPanel();
        };
        summaryButtons.Children.Add(_toggleLogButton);

        Grid.SetColumn(summaryButtons, 2);
        topRow.Children.Add(summaryButtons);
        bottom.Children.Add(topRow);

        var summaryGrid = new Grid { ColumnSpacing = 8 };
        _notificationSummaryBorder.Padding = new Thickness(10);
        _notificationSummaryBorder.CornerRadius = new CornerRadius(8);
        _notificationSummaryBorder.BorderThickness = new Thickness(1);
        _notificationSummaryBorder.BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush;
        _notificationSummaryBorder.Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush;
        _notificationSummaryBorder.Child = _statusText;
        summaryGrid.Children.Add(_notificationSummaryBorder);
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
        Grid.SetRow(_notificationExpandedGrid, 1);
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
            (X: 364.0, Y: 298.0, Size: 8.4, Highlight: false, Label: "Mgmt"),
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
            (0,4), (4,1), (1,5), (5,2), (2,6), (6,3), (1,7), (7,8), (8,3)
        };

        for (var i = 0; i < links.Length; i++)
        {
            var (from, to) = links[i];
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

        var actionsStack = new StackPanel { Spacing = 8 };
        actionsStack.Children.Add(new TextBlock
        {
            Text = "USB Host-Connect (Host-Freigaben)",
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
        });

        var actionRow = new Grid { ColumnSpacing = 8 };
        actionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        actionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        actionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        _usbRefreshButton = CreateIconButton("⟳", "Refresh", onClick: async (_, _) => await RefreshUsbAsync());
        _usbConnectButton = CreateIconButton("🔌", "Connect", onClick: async (_, _) => await ConnectUsbAsync());
        _usbDisconnectButton = CreateIconButton("⏏", "Disconnect", onClick: async (_, _) => await DisconnectUsbAsync());

        _usbRefreshButton.HorizontalAlignment = HorizontalAlignment.Stretch;
        _usbConnectButton.HorizontalAlignment = HorizontalAlignment.Stretch;
        _usbDisconnectButton.HorizontalAlignment = HorizontalAlignment.Stretch;

        _usbRefreshButton.IsEnabled = _isUsbClientAvailable;
        _usbConnectButton.IsEnabled = _isUsbClientAvailable;
        _usbDisconnectButton.IsEnabled = _isUsbClientAvailable;

        Grid.SetColumn(_usbRefreshButton, 0);
        Grid.SetColumn(_usbConnectButton, 1);
        Grid.SetColumn(_usbDisconnectButton, 2);

        actionRow.Children.Add(_usbRefreshButton);
        actionRow.Children.Add(_usbConnectButton);
        actionRow.Children.Add(_usbDisconnectButton);
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

        actionsCard.Child = actionsStack;
        root.Children.Add(actionsCard);

        var listBorder = new Border
        {
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            Padding = new Thickness(8),
            Child = _usbListView
        };

        Grid.SetRow(listBorder, 1);
        root.Children.Add(listBorder);

        return root;
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
        headingCard.Child = new StackPanel
        {
            Spacing = 4,
            Children =
            {
                new TextBlock { Text = "Konfiguration", FontSize = 20, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold },
                new TextBlock { Text = "Wichtige Einstellungen übersichtlich und schnell erreichbar.", Opacity = 0.9 }
            }
        };
        root.Children.Add(headingCard);

        var topBarGrid = new Grid
        {
            ColumnDefinitions =
            {
                new ColumnDefinition { Width = GridLength.Auto },
                new ColumnDefinition { Width = GridLength.Auto },
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }
            }
        };

        var saveButton = CreateIconButton("💾", "Speichern", onClick: async (_, _) => await SaveSettingsAsync());
        Grid.SetColumn(saveButton, 0);
        topBarGrid.Children.Add(saveButton);

        var reloadButton = CreateIconButton("⟳", "Neu laden", onClick: (_, _) =>
        {
            ApplyConfigToControls();
            AppendNotification("[Info] Einstellungen aus Config neu geladen.");
        });
        Grid.SetColumn(reloadButton, 1);
        topBarGrid.Children.Add(reloadButton);

        var configPathWrap = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Center
        };
        configPathWrap.Children.Add(new TextBlock
        {
            Text = "Aktive Config:",
            VerticalAlignment = VerticalAlignment.Center,
            Opacity = 0.9,
            Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
        });
        configPathWrap.Children.Add(new TextBlock
        {
            Text = GuestConfigService.DefaultConfigPath,
            VerticalAlignment = VerticalAlignment.Center,
            Opacity = 0.88,
            TextTrimming = TextTrimming.CharacterEllipsis,
            MaxWidth = 560
        });
        Grid.SetColumn(configPathWrap, 2);
        topBarGrid.Children.Add(configPathWrap);

        var topBar = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush,
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(14),
            Child = topBarGrid
        };
        root.Children.Add(topBar);

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

        var quickTogglesGrid = new Grid { ColumnSpacing = 10, RowSpacing = 6 };
        quickTogglesGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        quickTogglesGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        quickTogglesGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        quickTogglesGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        Grid.SetColumn(_minimizeToTrayCheckBox, 0);
        Grid.SetRow(_minimizeToTrayCheckBox, 0);
        quickTogglesGrid.Children.Add(_minimizeToTrayCheckBox);

        Grid.SetColumn(_startMinimizedCheckBox, 1);
        Grid.SetRow(_startMinimizedCheckBox, 0);
        quickTogglesGrid.Children.Add(_startMinimizedCheckBox);

        Grid.SetColumn(_startWithWindowsCheckBox, 0);
        Grid.SetRow(_startWithWindowsCheckBox, 1);
        quickTogglesGrid.Children.Add(_startWithWindowsCheckBox);

        var updateCheck = new CheckBox
        {
            Content = "Beim Start auf Updates prüfen",
            IsChecked = true,
            IsEnabled = false,
            Opacity = 0.7,
            Margin = new Thickness(0)
        };
        Grid.SetColumn(updateCheck, 1);
        Grid.SetRow(updateCheck, 1);
        quickTogglesGrid.Children.Add(updateCheck);

        systemStack.Children.Add(quickTogglesGrid);

        var themeRow = new Grid { ColumnSpacing = 8, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 2, 0, 0) };
        themeRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
        themeRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        themeRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
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
                _themeText.Text = _themeToggle.IsOn ? "Dunkles Theme" : "Helles Theme";
                await ApplyThemeAndRestartImmediatelyAsync();
            };
            _isThemeToggleHandlerAttached = true;
        }
        Grid.SetColumn(_themeToggle, 1);
        themeRow.Children.Add(_themeToggle);

        _themeText.VerticalAlignment = VerticalAlignment.Center;
        _themeText.Foreground = Application.Current.Resources["TextMutedBrush"] as Brush;
        _themeText.Opacity = 0.95;
        Grid.SetColumn(_themeText, 2);
        themeRow.Children.Add(_themeText);

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
        usbStack.Children.Add(new TextBlock
        {
            Text = "USB Host-Verbindung",
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            FontSize = 16
        });
        usbStack.Children.Add(new TextBlock
        {
            Text = "Host-IP oder Hostname für USB Attach eintragen (z.B. 192.168.178.10). Wenn leer, wird der Host aus SharePath verwendet.",
            TextWrapping = TextWrapping.Wrap,
            Opacity = 0.85
        });

        _usbHostAddressTextBox.PlaceholderText = "192.168.178.10 oder host.local";
        _usbHostAddressTextBox.MinWidth = 420;
        _usbHostAddressTextBox.MaxWidth = 620;
        _usbHostAddressTextBox.HorizontalAlignment = HorizontalAlignment.Left;
        _usbHostAddressTextBox.CornerRadius = new CornerRadius(8);
        usbStack.Children.Add(_usbHostAddressTextBox);

        usbSection.Child = usbStack;
        root.Children.Add(usbSection);

        return new ScrollViewer { Content = root };
    }

    private UIElement BuildInfoPage()
    {
        var version = ResolveGuestVersionText();

        var panel = new StackPanel { Spacing = 10 };
        panel.Children.Add(new TextBlock { Text = "Info", FontSize = 20, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });

        var versionWrap = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        versionWrap.Children.Add(new TextBlock { Text = "Version:", Opacity = 0.9 });
        versionWrap.Children.Add(new TextBlock { Opacity = 0.9, Text = version });
        panel.Children.Add(versionWrap);

        var updateWrap = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        updateWrap.Children.Add(new TextBlock { Text = "Update-Status:", Opacity = 0.9 });
        updateWrap.Children.Add(_updateStatusValueText);
        panel.Children.Add(updateWrap);

        var projectCard = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PageBackgroundBrush"] as Brush,
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(10)
        };

        var projectStack = new StackPanel { Spacing = 6 };
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

        var usbipStack = new StackPanel { Spacing = 6 };
        usbipStack.Children.Add(new TextBlock { Text = "Externe USB/IP Quelle", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
        usbipStack.Children.Add(new TextBlock { Text = "Quelle: vadimgrn/usbip-win2", Opacity = 0.9 });
        usbipStack.Children.Add(new TextBlock { Text = "Nutzung in HyperTool: externer CLI-Client ohne eigene GUI-Integration.", TextWrapping = TextWrapping.Wrap, Opacity = 0.85 });
        usbipStack.Children.Add(new TextBlock { Text = "Lizenz/Eigentümer: siehe Original-Repository von vadimgrn.", TextWrapping = TextWrapping.Wrap, Opacity = 0.85 });
        usbipCard.Child = usbipStack;
        panel.Children.Add(usbipCard);

        var buttonRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        buttonRow.Children.Add(CreateIconButton("🛰", "Update prüfen", onClick: async (_, _) => await CheckForUpdatesAsync()));
        buttonRow.Children.Add(CreateIconButton("⬇", "Update installieren", onClick: async (_, _) => await InstallUpdateAsync()));
        buttonRow.Children.Add(CreateIconButton("🌐", "Changelog / Release", onClick: (_, _) => OpenReleasePage()));
        buttonRow.Children.Add(CreateIconButton("🔗", "usbip-win2 Quelle", onClick: (_, _) => OpenUsbipClientRepository()));
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
            1 => GetOrCreateSettingsPage(),
            _ => _infoPage ??= BuildInfoPage()
        };
    }

    private UIElement GetOrCreateSettingsPage()
    {
        _settingsPage ??= BuildSettingsPage();
        ApplyConfigToControls();
        return _settingsPage;
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
    }

    private async Task SaveSettingsAsync()
    {
        _config.Ui.Theme = (_themeCombo.SelectedItem as string) ?? "dark";
        _config.Ui.StartWithWindows = _startWithWindowsCheckBox.IsChecked == true;
        _config.Ui.StartMinimized = _startMinimizedCheckBox.IsChecked == true;
        _config.Ui.MinimizeToTray = _minimizeToTrayCheckBox.IsChecked == true;
        _config.Usb ??= new GuestUsbSettings();
        _config.Usb.HostAddress = (_usbHostAddressTextBox.Text ?? string.Empty).Trim();

        await _saveConfigAsync(_config);
        ApplyTheme(_config.Ui.Theme);

        AppendNotification("[Info] Einstellungen gespeichert.");
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

    private async Task ConnectUsbAsync()
    {
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
                using var player = new SoundPlayer(soundPath);
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
        _installerDownloadUrl = result.InstallerDownloadUrl ?? string.Empty;
        _installerFileName = result.InstallerFileName ?? string.Empty;

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

            Process.Start(new ProcessStartInfo
            {
                FileName = installerPath,
                UseShellExecute = true
            });

            AppendNotification("[Info] Installer gestartet.");
        }
        catch (Exception ex)
        {
            AppendNotification($"[Error] Update-Installation fehlgeschlagen: {ex.Message}");
        }
    }

    private static string ResolveInstallerFileName(string downloadUrl, string? fileName)
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

        return "HyperTool-Guest-Setup.exe";
    }

    private void OnLoggerEntryWritten(string message)
    {
        if (!DispatcherQueue.TryEnqueue(() => AppendNotification(message)))
        {
            AppendNotification(message);
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
        var toggleLabel = _isLogExpanded ? "Log einklappen" : "Log ausklappen";
        _toggleLogButton.Content = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            VerticalAlignment = VerticalAlignment.Center,
            Children =
            {
                new TextBlock
                {
                    Text = "▸",
                    FontSize = 18,
                    FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                    VerticalAlignment = VerticalAlignment.Center
                },
                new TextBlock
                {
                    Text = toggleLabel,
                    FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                    VerticalAlignment = VerticalAlignment.Center
                }
            }
        };
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
        button.Click += (_, _) =>
        {
            if (_selectedMenuIndex == index)
            {
                return;
            }

            _selectedMenuIndex = index;
            UpdateNavSelection();
            UpdatePageContent();
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
