using HyperTool.Services;
using HyperTool.Models;
using HyperTool.ViewModels;
using HyperTool.WinUI.Helpers;
using HyperTool.WinUI.Services;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Markup;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Animation;
using Microsoft.UI.Xaml.Shapes;
using Serilog;
using System.Linq;
using System.ComponentModel;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Media;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Windows.Input;
using Windows.Media.Core;
using Windows.Media.Playback;
using Windows.Graphics;
using Windows.UI;

namespace HyperTool.WinUI.Views;

public sealed class MainWindow : Window
{
    public const int DefaultWindowWidth = 1400;
    public const int DefaultWindowHeight = 950;
    private const string ToolRestartIcon = "↻";
    private const string ToolRestartLabel = "Tool neu starten";
    private const string HostUsbRuntimeOwner = "dorssel";
    private const string HostUsbRuntimeRepo = "usbipd-win";
    private const string HostUsbRuntimeAssetHint = "x64";
    private const string GuestSingleInstancePipeName = "HyperTool.Guest.SingleInstance.Activate";
    private static readonly HttpClient RuntimeInstallerDownloadClient = new();

    private readonly IThemeService _themeService;
    private readonly MainViewModel _viewModel;
    private readonly IHostSharedFolderService _hostSharedFolderService = new HostSharedFolderService();
    private readonly List<Button> _navButtons = [];
    private Button? _usbNavButton;
    private readonly StackPanel _vmChipPanel = new() { Orientation = Orientation.Horizontal, Spacing = 8 };
    private readonly ScrollViewer _vmChipScrollViewer = new()
    {
        HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden,
        VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
        HorizontalScrollMode = ScrollMode.Enabled,
        VerticalScrollMode = ScrollMode.Disabled
    };
    private readonly Border _vmChipLeftFadeOverlay = new() { Width = 26, HorizontalAlignment = HorizontalAlignment.Left, IsHitTestVisible = false, Visibility = Visibility.Collapsed };
    private readonly Border _vmChipRightFadeOverlay = new() { Width = 26, HorizontalAlignment = HorizontalAlignment.Right, IsHitTestVisible = false, Visibility = Visibility.Collapsed };
    private readonly Button _vmChipsLeftButton = new();
    private readonly Button _vmChipsRightButton = new();
    private readonly ContentPresenter _pageContent = new();
    private readonly TextBlock _statusText = new();
    private readonly Border _configurationNoticeBorder = new() { Visibility = Visibility.Collapsed };
    private readonly TextBlock _configurationNoticeText = new() { TextWrapping = TextWrapping.Wrap };
    private readonly ProgressRing _busyRing = new() { Width = 18, Height = 18, Visibility = Visibility.Collapsed };
    private readonly TextBlock _busyText = new() { Visibility = Visibility.Collapsed };
    private readonly ProgressBar _busyProgress = new() { Height = 10, Visibility = Visibility.Collapsed, Minimum = 0, Maximum = 100, CornerRadius = new CornerRadius(6) };
    private readonly Border _busyPercentBadge = new() { Visibility = Visibility.Collapsed };
    private readonly TextBlock _busyPercentText = new();
    private readonly Border _notificationSummaryBorder = new();
    private readonly Grid _notificationExpandedGrid = new() { Visibility = Visibility.Collapsed };
    private readonly ListView _notificationsListView = new();
    private readonly Button _toggleLogButton = new();
    private readonly TreeView _checkpointTreeView = new();
    private readonly ListView _usbDevicesListView = new();
    private readonly ListView _sharedFoldersListView = new();
    private readonly Ellipse _sharedFolderFileServiceStatusDot = new() { Width = 10, Height = 10, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock _sharedFolderFileServiceStatusText = new() { Opacity = 0.9, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock _selectedVmStateText = new() { FontSize = 12, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBox _sharedFolderPathTextBox = new();
    private readonly TextBox _sharedFolderShareNameTextBox = new();
    private readonly CheckBox _sharedFolderReadOnlyCheckBox = new() { Content = "Freigabe nur Lesen" };
    private readonly TextBlock _sharedFolderStatusText = new() { Opacity = 0.88, TextWrapping = TextWrapping.Wrap, Text = "Bereit." };
    private Button? _sharedFolderNavButton;
    private Button? _sharedFolderNewButton;
    private Button? _sharedFolderSaveButton;
    private Button? _sharedFolderDeleteButton;
    private Button? _sharedFolderFileServiceInstallButton;
    private string _sharedFolderEditingId = string.Empty;
    private string _sharedFolderLastError = "-";
    private readonly SemaphoreSlim _sharedFolderEnabledToggleGate = new(1, 1);
    private bool _sharedFolderCollectionHandlersAttached;
    private readonly CheckBox _usbAutoShareCheckBox = new();
    private readonly CheckBox _usbAutoDetachOnDisconnectCheckBox = new();
    private readonly CheckBox _usbUnshareOnExitCheckBox = new();
    private readonly ToggleSwitch _usbHostFeatureEnabledToggleSwitch = new() { Header = null, OffContent = "", OnContent = "" };
    private readonly Border _usbHostFeatureStatusChip = new() { CornerRadius = new CornerRadius(9), MinHeight = 30, BorderThickness = new Thickness(1), Padding = new Thickness(10, 5, 10, 5), HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock _usbHostFeatureStatusChipText = new() { FontSize = 12, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold };
    private readonly TextBlock _usbDisabledOverlayText = new() { Text = "Deaktiviert", Opacity = 0.34, FontSize = 34, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, IsHitTestVisible = false };
    private StackPanel? _usbFeatureControlsPanel;
    private readonly ToggleSwitch _sharedFolderHostFeatureEnabledToggleSwitch = new() { Header = null, OffContent = "", OnContent = "" };
    private readonly Border _sharedFolderHostFeatureStatusChip = new() { CornerRadius = new CornerRadius(9), MinHeight = 30, BorderThickness = new Thickness(1), Padding = new Thickness(10, 5, 10, 5), HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock _sharedFolderHostFeatureStatusChipText = new() { FontSize = 12, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold };
    private readonly TextBlock _sharedFolderDisabledOverlayText = new() { Text = "Deaktiviert", Opacity = 0.34, FontSize = 34, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, IsHitTestVisible = false };
    private readonly Ellipse _usbRuntimeStatusDot = new() { Width = 10, Height = 10, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock _usbRuntimeStatusText = new() { Opacity = 0.9, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock _usbRuntimeHintText = new() { TextWrapping = TextWrapping.Wrap, Opacity = 0.9 };
    private readonly TextBlock _usbRemoteFxPolicyHintText = new() { TextWrapping = TextWrapping.Wrap, Opacity = 0.9 };
    private Button? _usbRuntimeInstallButton;
    private Button? _usbRuntimeRestartButton;
    private Button? _usbRefreshButton;
    private Button? _usbShareButton;
    private Button? _usbUnshareButton;
    private readonly StackPanel _vmAdapterCardsPanel = new() { Spacing = 10 };
    private readonly Dictionary<string, TreeViewNode> _checkpointNodesById = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<TreeViewNode, HyperVCheckpointTreeItem> _checkpointItemsByNode = [];
    private bool _isUpdatingCheckpointTreeSelection;
    private readonly RotateTransform _logoRotateTransform = new();
    private UIElement? _vmPage;
    private UIElement? _snapshotsPage;
    private UIElement? _sharedFoldersPage;
    private UIElement? _configPage;
    private UIElement? _infoPage;
    private UIElement? _usbPage;
    private HelpWindow? _helpWindow;
    private HostNetworkWindow? _hostNetworkWindow;
    private Grid? _windowHost;
    private Border? _themeTransitionOverlay;
    private TextBlock? _themeTransitionStatusText;
    private bool _isMainLayoutLoaded;
    private bool _isThemeRestartInProgress;
    private DispatcherQueueTimer? _startupStatusTimer;
    private TextBlock? _startupStatusText;
    private int _startupStatusIndex;
    private bool _isStartupStatusTransitionRunning;
    private MediaPlayer? _logoSpinPlayer;
    private DispatcherQueueTimer? _vmChipRefreshDebounceTimer;
    private string _vmChipRefreshSignature = string.Empty;
    private bool _suppressHostFeatureToggleEvents;

    public MainWindow(IThemeService themeService, MainViewModel viewModel, bool showStartupSplash = false)
    {
        _themeService = themeService;
        _viewModel = viewModel;
        Title = "HyperTool";
        ExtendsContentIntoTitleBar = false;
        DwmWindowHelper.ApplyRoundedCorners(this);
        WindowHandleProvider.MainWindowHandle = WinRT.Interop.WindowNative.GetWindowHandle(this);
        TryApplyInitialWindowSize();
        TryApplyWindowIcon();

        var mainLayout = BuildLayout();
        UIElement initialMainContent;
        if (showStartupSplash)
        {
            var startupSplash = BuildStartupSplashContent();
            mainLayout.Opacity = 0;

            var transitionHost = new Grid();
            transitionHost.Children.Add(mainLayout);
            transitionHost.Children.Add(startupSplash);

            initialMainContent = transitionHost;
            _isMainLayoutLoaded = false;

            _ = ShowStartupSplashThenLoadMainLayoutAsync(startupSplash, mainLayout, transitionHost);
        }
        else
        {
            initialMainContent = mainLayout;
            _isMainLayoutLoaded = true;
        }

        _windowHost = BuildWindowHost(initialMainContent);
        Content = _windowHost;

        _viewModel.PropertyChanged += OnViewModelPropertyChanged;
        _viewModel.AvailableVms.CollectionChanged += OnAvailableVmsCollectionChanged;
        _viewModel.AvailableVmNetworkAdapters.CollectionChanged += OnVmNetworkAdaptersCollectionChanged;
        _viewModel.AvailableSwitches.CollectionChanged += OnVmNetworkAdaptersCollectionChanged;
        _viewModel.AvailableCheckpointTree.CollectionChanged += OnCheckpointTreeCollectionChanged;
                _vmChipRefreshDebounceTimer = DispatcherQueue.CreateTimer();
                _vmChipRefreshDebounceTimer.Interval = TimeSpan.FromMilliseconds(90);
                _vmChipRefreshDebounceTimer.Tick += (_, _) =>
                {
                    _vmChipRefreshDebounceTimer?.Stop();
                    RefreshVmChips();
                };
        _themeService.ApplyTheme(_viewModel.UiTheme);
        ApplyRequestedTheme();
        UpdateTitleBarAppearance();
        UpdateHostFeatureAvailabilityUi();

        if (!showStartupSplash)
        {
            RequestVmChipsRefresh();
            RefreshVmAdapterCards();
        }

        Closed += OnWindowClosed;
    }

    private async Task ShowStartupSplashThenLoadMainLayoutAsync(UIElement startupSplash, UIElement mainLayout, Grid transitionHost)
    {
        await Task.Delay(TimeSpan.FromMilliseconds(LifecycleVisuals.SplashMinVisibleMs));

        _startupStatusTimer?.Stop();
        _startupStatusTimer = null;

        var fadeOutSplash = new DoubleAnimation
        {
            From = 1.0,
            To = 0.0,
            Duration = TimeSpan.FromMilliseconds(420),
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(fadeOutSplash, startupSplash);
        Storyboard.SetTargetProperty(fadeOutSplash, "Opacity");

        var fadeInMain = new DoubleAnimation
        {
            From = 0.0,
            To = 1.0,
            Duration = TimeSpan.FromMilliseconds(420),
            EasingFunction = LifecycleVisuals.CreateEaseOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(fadeInMain, mainLayout);
        Storyboard.SetTargetProperty(fadeInMain, "Opacity");

        var transition = new Storyboard();
        transition.Children.Add(fadeOutSplash);
        transition.Children.Add(fadeInMain);
        transition.Begin();

        await Task.Delay(460);

        transitionHost.Children.Remove(startupSplash);
        _isMainLayoutLoaded = true;
        RequestVmChipsRefresh();
        RefreshVmAdapterCards();
    }

    private Grid BuildWindowHost(UIElement initialMainContent)
    {
        var host = new Grid();
        host.Children.Add(initialMainContent);

        _themeTransitionOverlay = BuildThemeTransitionOverlay();
        host.Children.Add(_themeTransitionOverlay);

        return host;
    }

    private void SetWindowMainContent(UIElement content)
    {
        if (_windowHost is null)
        {
            Content = content;
            return;
        }

        if (_windowHost.Children.Count == 0)
        {
            _windowHost.Children.Add(content);
            return;
        }

        _windowHost.Children[0] = content;
    }

    private Border BuildThemeTransitionOverlay()
    {
        var overlayRoot = new Border
        {
            Visibility = Visibility.Collapsed,
            Opacity = 0,
            IsHitTestVisible = false,
            Background = new SolidColorBrush(Windows.UI.Color.FromArgb(0xD6, 0x08, 0x11, 0x22))
        };

        var overlayGrid = new Grid
        {
            Background = LifecycleVisuals.CreateRootBackgroundBrush()
        };

        overlayGrid.Children.Add(new TextBlock
        {
            Text = "Copyright: koerby",
            HorizontalAlignment = HorizontalAlignment.Left,
            VerticalAlignment = VerticalAlignment.Bottom,
            Margin = new Thickness(20, 0, 0, 14),
            FontSize = 12,
            Opacity = 0.72,
            Foreground = new SolidColorBrush(Windows.UI.Color.FromArgb(0xC8, 0x9B, 0xB7, 0xD7))
        });

        overlayGrid.Children.Add(new TextBlock
        {
            Text = LifecycleVisuals.ResolveDisplayVersion(_viewModel.AppVersion),
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Bottom,
            Margin = new Thickness(0, 0, 20, 14),
            FontSize = 12,
            Opacity = 0.72,
            Foreground = new SolidColorBrush(Windows.UI.Color.FromArgb(0xC8, 0x9B, 0xB7, 0xD7))
        });

        overlayGrid.Children.Add(new Rectangle
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment = VerticalAlignment.Stretch,
            Fill = LifecycleVisuals.CreateVignetteBrush(0x78)
        });

        var card = new Border
        {
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Width = 420,
            Padding = new Thickness(24, 22, 24, 18),
            CornerRadius = new CornerRadius(20),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(LifecycleVisuals.CardBorder),
            Background = LifecycleVisuals.CreateCardSurfaceBrush(),
            Shadow = new ThemeShadow()
        };

        var stack = new StackPanel
        {
            Spacing = 10,
            HorizontalAlignment = HorizontalAlignment.Stretch
        };

        stack.Children.Add(new Image
        {
            Source = new Microsoft.UI.Xaml.Media.Imaging.BitmapImage(new Uri("ms-appx:///Assets/HyperTool.Icon.Transparent.png")),
            Width = 54,
            Height = 54,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Opacity = 0.95
        });

        stack.Children.Add(new TextBlock
        {
            Text = "HyperTool",
            HorizontalAlignment = HorizontalAlignment.Center,
            FontSize = 24,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            Foreground = new SolidColorBrush(LifecycleVisuals.TextPrimary)
        });

        _themeTransitionStatusText = new TextBlock
        {
            Text = "Layout wird aktualisiert …",
            HorizontalAlignment = HorizontalAlignment.Center,
            TextAlignment = TextAlignment.Center,
            FontSize = 13,
            Foreground = new SolidColorBrush(LifecycleVisuals.TextSecondary),
            Opacity = 0.95
        };
        stack.Children.Add(_themeTransitionStatusText);

        stack.Children.Add(new ProgressRing
        {
            Width = 30,
            Height = 30,
            IsActive = true,
            HorizontalAlignment = HorizontalAlignment.Center,
            Foreground = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, 0x8D, 0xCF, 0xFF)),
            Margin = new Thickness(0, 2, 0, 0)
        });

        card.Child = stack;
        overlayGrid.Children.Add(card);
        overlayRoot.Child = overlayGrid;

        return overlayRoot;
    }

    private async Task ShowThemeTransitionOverlayAsync(string statusText)
    {
        if (_themeTransitionOverlay is null)
        {
            return;
        }

        if (_themeTransitionStatusText is not null)
        {
            _themeTransitionStatusText.Text = statusText;
        }

        _themeTransitionOverlay.IsHitTestVisible = true;
        _themeTransitionOverlay.Visibility = Visibility.Visible;
        await AnimateOpacityAsync(_themeTransitionOverlay, from: _themeTransitionOverlay.Opacity, to: 1.0, durationMs: 130);
    }

    private async Task HideThemeTransitionOverlaySafeAsync()
    {
        if (_themeTransitionOverlay is null)
        {
            return;
        }

        try
        {
            await AnimateOpacityAsync(_themeTransitionOverlay, from: _themeTransitionOverlay.Opacity, to: 0.0, durationMs: 140);
        }
        catch
        {
        }
        finally
        {
            _themeTransitionOverlay.Opacity = 0;
            _themeTransitionOverlay.Visibility = Visibility.Collapsed;
            _themeTransitionOverlay.IsHitTestVisible = false;
        }
    }

    public Task ShowLifecycleGuardAsync(string statusText)
    {
        return ShowThemeTransitionOverlayAsync(statusText);
    }

    public Task HideLifecycleGuardAsync()
    {
        return HideThemeTransitionOverlaySafeAsync();
    }

    private static async Task AnimateOpacityAsync(UIElement target, double from, double to, int durationMs)
    {
        var animation = new DoubleAnimation
        {
            From = from,
            To = to,
            Duration = TimeSpan.FromMilliseconds(durationMs),
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };

        var storyboard = new Storyboard();
        Storyboard.SetTarget(animation, target);
        Storyboard.SetTargetProperty(animation, "Opacity");
        storyboard.Children.Add(animation);
        storyboard.Begin();

        await Task.Delay(durationMs + 24);
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

    private UIElement BuildStartupSplashContent()
    {
        var root = new Grid
        {
            Background = LifecycleVisuals.CreateRootBackgroundBrush()
        };

        var ambientStoryboard = new Storyboard { RepeatBehavior = RepeatBehavior.Forever };
        var sequenceStoryboard = new Storyboard();

        var focusLayerPrimary = new Ellipse
        {
            Width = 840,
            Height = 840,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Fill = LifecycleVisuals.CreateCenterFocusBrush(LifecycleVisuals.BackgroundFocusPrimary)
        };
        root.Children.Add(focusLayerPrimary);

        var focusLayerSecondary = new Ellipse
        {
            Width = 620,
            Height = 620,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(90, -56, -90, 56),
            Fill = LifecycleVisuals.CreateCenterFocusBrush(LifecycleVisuals.BackgroundFocusSecondary),
            Opacity = 0.74
        };
        root.Children.Add(focusLayerSecondary);

        var focusLayerTertiary = new Ellipse
        {
            Width = 460,
            Height = 460,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(-96, 74, 96, -74),
            Fill = LifecycleVisuals.CreateCenterFocusBrush(LifecycleVisuals.BackgroundFocusTertiary),
            Opacity = 0.48
        };
        root.Children.Add(focusLayerTertiary);

        var overlay = new Canvas { IsHitTestVisible = false };
        var lightBand1 = CreateSplashLightBand(700, 48, 0.08, -14, -900, 1280, 5600, 0, ambientStoryboard);
        var lightBand2 = CreateSplashLightBand(560, 36, 0.06, -11, -850, 1240, 6400, 920, ambientStoryboard);
        var lightBand3 = CreateSplashLightBand(500, 30, 0.05, -9, -810, 1200, 7000, 1840, ambientStoryboard);
        Canvas.SetTop(lightBand1, 120);
        Canvas.SetTop(lightBand2, 302);
        Canvas.SetTop(lightBand3, 530);
        overlay.Children.Add(lightBand1);
        overlay.Children.Add(lightBand2);
        overlay.Children.Add(lightBand3);
        root.Children.Add(overlay);

        BuildSplashNetworkLayer(root, ambientStoryboard, sequenceStoryboard);

        var vignette = new Rectangle
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment = VerticalAlignment.Stretch,
            Fill = LifecycleVisuals.CreateVignetteBrush(0x66)
        };
        root.Children.Add(vignette);

        var splashVersionText = new TextBlock
        {
            Text = LifecycleVisuals.ResolveDisplayVersion(_viewModel.AppVersion),
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Bottom,
            Margin = new Thickness(0, 0, 20, 14),
            FontSize = 12,
            Opacity = 0.72,
            Foreground = new SolidColorBrush(Windows.UI.Color.FromArgb(0xC8, 0x9B, 0xB7, 0xD7))
        };
        root.Children.Add(splashVersionText);

        var splashCopyrightText = new TextBlock
        {
            Text = "Copyright: koerby",
            HorizontalAlignment = HorizontalAlignment.Left,
            VerticalAlignment = VerticalAlignment.Bottom,
            Margin = new Thickness(20, 0, 0, 14),
            FontSize = 12,
            Opacity = 0.72,
            Foreground = new SolidColorBrush(Windows.UI.Color.FromArgb(0xC8, 0x9B, 0xB7, 0xD7))
        };
        root.Children.Add(splashCopyrightText);

        var card = new Border
        {
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Width = 540,
            Padding = new Thickness(30, 28, 30, 24),
            CornerRadius = new CornerRadius(24),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(LifecycleVisuals.CardBorder),
            Background = LifecycleVisuals.CreateCardSurfaceBrush(),
            Shadow = new ThemeShadow(),
            Opacity = 0
        };

        var innerFrame = new Border
        {
            CornerRadius = new CornerRadius(18),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(LifecycleVisuals.CardInnerOutline),
            Padding = new Thickness(20, 20, 20, 16),
            Background = LifecycleVisuals.CreateCardInnerBrush()
        };

        var stack = new StackPanel { Spacing = 12, HorizontalAlignment = HorizontalAlignment.Stretch };
        var iconHost = new Grid { Width = 142, Height = 142, HorizontalAlignment = HorizontalAlignment.Center };

        var outerHalo = new Ellipse
        {
            Width = 138,
            Height = 138,
            Fill = new RadialGradientBrush
            {
                Center = new Windows.Foundation.Point(0.5, 0.5),
                GradientOrigin = new Windows.Foundation.Point(0.5, 0.5),
                RadiusX = 0.5,
                RadiusY = 0.5,
                GradientStops =
                {
                    new GradientStop { Color = Windows.UI.Color.FromArgb(0x22, 0x74, 0xC6, 0xFF), Offset = 0.0 },
                    new GradientStop { Color = Windows.UI.Color.FromArgb(0x00, 0x74, 0xC6, 0xFF), Offset = 1.0 }
                }
            },
            Opacity = 0.34
        };
        iconHost.Children.Add(outerHalo);

        var ringWave = new Ellipse
        {
            Width = 132,
            Height = 132,
            StrokeThickness = 1.3,
            Stroke = new SolidColorBrush(Windows.UI.Color.FromArgb(0x96, 0x8E, 0xD7, 0xFF)),
            Opacity = 0.0,
            RenderTransformOrigin = new Windows.Foundation.Point(0.5, 0.5),
            RenderTransform = new ScaleTransform { ScaleX = 0.92, ScaleY = 0.92 }
        };
        iconHost.Children.Add(ringWave);

        var halo = new Ellipse
        {
            Width = 124,
            Height = 124,
            StrokeThickness = 1.2,
            Stroke = new SolidColorBrush(Windows.UI.Color.FromArgb(0x92, 0x75, 0xC9, 0xFF)),
            Opacity = 0.34,
            RenderTransformOrigin = new Windows.Foundation.Point(0.5, 0.5),
            RenderTransform = new ScaleTransform { ScaleX = 1.0, ScaleY = 1.0 }
        };
        iconHost.Children.Add(halo);

        var innerRing = new Ellipse
        {
            Width = 100,
            Height = 100,
            StrokeThickness = 1.0,
            Stroke = new SolidColorBrush(Windows.UI.Color.FromArgb(0x78, 0x8A, 0xD2, 0xFF)),
            Opacity = 0.40,
            RenderTransformOrigin = new Windows.Foundation.Point(0.5, 0.5),
            RenderTransform = new ScaleTransform { ScaleX = 1.0, ScaleY = 1.0 }
        };
        iconHost.Children.Add(innerRing);

        var glow = new Border
        {
            Width = 110,
            Height = 110,
            CornerRadius = new CornerRadius(55),
            Background = new SolidColorBrush(Windows.UI.Color.FromArgb(0x72, 0x68, 0xCE, 0xFF)),
            Opacity = 0.18
        };
        iconHost.Children.Add(glow);

        var coreGlow = new Border
        {
            Width = 88,
            Height = 88,
            CornerRadius = new CornerRadius(44),
            Background = new SolidColorBrush(Windows.UI.Color.FromArgb(0x68, 0x6B, 0xCF, 0xFF)),
            Opacity = 0.28
        };
        iconHost.Children.Add(coreGlow);

        var logoPlate = new Border
        {
            Width = 84,
            Height = 84,
            CornerRadius = new CornerRadius(42),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(Windows.UI.Color.FromArgb(0x68, 0x99, 0xC7, 0xF3)),
            Background = new SolidColorBrush(Windows.UI.Color.FromArgb(0x24, 0x9D, 0xC8, 0xF6))
        };
        iconHost.Children.Add(logoPlate);

        var icon = new Image
        {
            Source = new Microsoft.UI.Xaml.Media.Imaging.BitmapImage(new Uri("ms-appx:///Assets/HyperTool.Icon.Transparent.png")),
            Width = 74,
            Height = 74,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Opacity = 0.0,
            RenderTransformOrigin = new Windows.Foundation.Point(0.5, 0.5),
            RenderTransform = new ScaleTransform { ScaleX = 0.92, ScaleY = 0.92 }
        };
        iconHost.Children.Add(icon);

        stack.Children.Add(iconHost);
        stack.Children.Add(new TextBlock
        {
            Text = "HyperTool",
            FontSize = 34,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            HorizontalAlignment = HorizontalAlignment.Center,
            Foreground = new SolidColorBrush(LifecycleVisuals.TextPrimary)
        });

        _startupStatusText = new TextBlock
        {
            Text = LifecycleVisuals.StartupStatusMessages[0],
            HorizontalAlignment = HorizontalAlignment.Center,
            Foreground = new SolidColorBrush(LifecycleVisuals.TextSecondary),
            Margin = new Thickness(0, 6, 0, 2),
            FontSize = 13,
            Opacity = 0.94
        };
        stack.Children.Add(_startupStatusText);

        var progressHost = new Border
        {
            Height = 10,
            CornerRadius = new CornerRadius(5),
            Margin = new Thickness(4, 6, 4, 2),
            Background = new SolidColorBrush(LifecycleVisuals.ProgressTrack),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(LifecycleVisuals.ProgressTrackEdge)
        };

        var progressLayer = new Grid();
        var progressFill = new Rectangle
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            RadiusX = 5,
            RadiusY = 5,
            Fill = LifecycleVisuals.CreateProgressBrush(),
            RenderTransformOrigin = new Windows.Foundation.Point(0, 0.5),
            RenderTransform = new ScaleTransform { ScaleX = 0.1, ScaleY = 1.0 }
        };
        var progressShimmer = new Rectangle
        {
            HorizontalAlignment = HorizontalAlignment.Left,
            Width = 84,
            RadiusX = 5,
            RadiusY = 5,
            Opacity = 0.24,
            Fill = new LinearGradientBrush
            {
                StartPoint = new Windows.Foundation.Point(0, 0.5),
                EndPoint = new Windows.Foundation.Point(1, 0.5),
                GradientStops =
                {
                    new GradientStop { Color = Windows.UI.Color.FromArgb(0x00, 0xDF, 0xF4, 0xFF), Offset = 0 },
                    new GradientStop { Color = Windows.UI.Color.FromArgb(0xCC, 0xDF, 0xF4, 0xFF), Offset = 0.5 },
                    new GradientStop { Color = Windows.UI.Color.FromArgb(0x00, 0xDF, 0xF4, 0xFF), Offset = 1 }
                }
            },
            RenderTransform = new TranslateTransform { X = -112 }
        };

        progressLayer.Children.Add(progressFill);
        progressLayer.Children.Add(progressShimmer);
        progressHost.Child = progressLayer;
        stack.Children.Add(progressHost);

        innerFrame.Child = stack;
        card.Child = innerFrame;

        root.Children.Add(card);

        RunStartupSplashAnimation(glow, coreGlow, outerHalo, halo, innerRing, ringWave, icon, card, progressFill, progressShimmer, ambientStoryboard, sequenceStoryboard);

        _startupStatusIndex = 0;
        _isStartupStatusTransitionRunning = false;
        _startupStatusTimer = DispatcherQueue.CreateTimer();
        _startupStatusTimer.Interval = TimeSpan.FromMilliseconds(LifecycleVisuals.SplashStatusCycleMs);
        _startupStatusTimer.Tick += async (_, _) =>
        {
            if (_startupStatusText is null || _isStartupStatusTransitionRunning)
            {
                return;
            }

            _isStartupStatusTransitionRunning = true;
            _startupStatusIndex = (_startupStatusIndex + 1) % LifecycleVisuals.StartupStatusMessages.Length;
            await AnimateStartupStatusChangeAsync(LifecycleVisuals.StartupStatusMessages[_startupStatusIndex]);
            _isStartupStatusTransitionRunning = false;
        };
        _startupStatusTimer.Start();

        return root;
    }

    private void BuildSplashNetworkLayer(Grid root, Storyboard ambientStoryboard, Storyboard sequenceStoryboard)
    {
        var canvas = new Canvas { IsHitTestVisible = false };

        var nodes = new[]
        {
            new SplashNodeSpec(new Windows.Foundation.Point(176, 250), 10.2, "Host", true),
            new SplashNodeSpec(new Windows.Foundation.Point(298, 290), 8.6, "Mgmt", false),
            new SplashNodeSpec(new Windows.Foundation.Point(426, 324), 8.8, "Hyper-V", true),
            new SplashNodeSpec(new Windows.Foundation.Point(582, 300), 8.4, null, false),
            new SplashNodeSpec(new Windows.Foundation.Point(742, 278), 9.6, "VM", true),
            new SplashNodeSpec(new Windows.Foundation.Point(872, 316), 8.6, "Client", false),
            new SplashNodeSpec(new Windows.Foundation.Point(980, 352), 8.4, "Target", false),
            new SplashNodeSpec(new Windows.Foundation.Point(352, 388), 7.4, null, false),
            new SplashNodeSpec(new Windows.Foundation.Point(792, 392), 7.4, null, false)
        };

        var links = new (int From, int To)[]
        {
            (0,1), (1,2), (2,3), (3,4), (4,5), (5,6), (2,7), (7,3), (4,8), (8,5)
        };

        foreach (var nodeSpec in nodes)
        {
            var node = new Ellipse
            {
                Width = nodeSpec.Size,
                Height = nodeSpec.Size,
                Fill = new SolidColorBrush(nodeSpec.IsPrimary
                    ? Windows.UI.Color.FromArgb(0xE4, 0x69, 0xC2, 0xFF)
                    : LifecycleVisuals.NodeColor),
                Stroke = new SolidColorBrush(LifecycleVisuals.NodeStroke),
                StrokeThickness = nodeSpec.IsPrimary ? 1.0 : 0.9,
                Opacity = nodeSpec.IsPrimary ? 0.76 : 0.58
            };
            Canvas.SetLeft(node, nodeSpec.Point.X - (nodeSpec.Size / 2));
            Canvas.SetTop(node, nodeSpec.Point.Y - (nodeSpec.Size / 2));
            canvas.Children.Add(node);

            if (!string.IsNullOrWhiteSpace(nodeSpec.Label))
            {
                var label = new TextBlock
                {
                    Text = nodeSpec.Label,
                    FontSize = 11,
                    FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                    Opacity = 0.68,
                    Foreground = new SolidColorBrush(Windows.UI.Color.FromArgb(0xD0, 0xA8, 0xC7, 0xE8))
                };
                Canvas.SetLeft(label, nodeSpec.Point.X - 24);
                Canvas.SetTop(label, nodeSpec.Point.Y + 14);
                canvas.Children.Add(label);
            }

            var nodePulse = new DoubleAnimation
            {
                From = nodeSpec.IsPrimary ? 0.48 : 0.34,
                To = nodeSpec.IsPrimary ? 0.82 : 0.62,
                Duration = TimeSpan.FromMilliseconds(nodeSpec.IsPrimary ? 1900 : 2300),
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                EasingFunction = LifecycleVisuals.CreateEaseInOut(),
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(nodePulse, node);
            Storyboard.SetTargetProperty(nodePulse, "Opacity");
            ambientStoryboard.Children.Add(nodePulse);
        }

        for (var i = 0; i < links.Length; i++)
        {
            var (from, to) = links[i];
            var a = nodes[from].Point;
            var b = nodes[to].Point;

            var line = new Line
            {
                X1 = a.X,
                Y1 = a.Y,
                X2 = b.X,
                Y2 = b.Y,
                StrokeThickness = 0.8,
                Stroke = new SolidColorBrush(LifecycleVisuals.LineColor),
                Opacity = 0.0
            };
            canvas.Children.Add(line);

            var appear = new DoubleAnimation
            {
                From = 0.0,
                To = 0.48,
                Duration = TimeSpan.FromMilliseconds(LifecycleVisuals.SplashPhaseNetworkInMs),
                BeginTime = TimeSpan.FromMilliseconds(560 + (i * 150)),
                FillBehavior = FillBehavior.HoldEnd,
                EasingFunction = LifecycleVisuals.CreateEaseOut(),
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(appear, line);
            Storyboard.SetTargetProperty(appear, "Opacity");
            sequenceStoryboard.Children.Add(appear);

            var pulse = new Ellipse
            {
                Width = 3.2,
                Height = 3.2,
                Fill = new SolidColorBrush(LifecycleVisuals.PulseColor),
                Opacity = 0.0
            };
            Canvas.SetLeft(pulse, a.X - 1.6);
            Canvas.SetTop(pulse, a.Y - 1.6);
            canvas.Children.Add(pulse);

            var pulseFade = new DoubleAnimation
            {
                From = 0.0,
                To = 0.62,
                Duration = TimeSpan.FromMilliseconds(260),
                BeginTime = TimeSpan.FromMilliseconds(LifecycleVisuals.SplashPhasePulseStartMs + (i * 130)),
                FillBehavior = FillBehavior.HoldEnd,
                EasingFunction = LifecycleVisuals.CreateEaseOut(),
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(pulseFade, pulse);
            Storyboard.SetTargetProperty(pulseFade, "Opacity");
            sequenceStoryboard.Children.Add(pulseFade);

            var moveX = new DoubleAnimation
            {
                From = a.X - 1.6,
                To = b.X - 1.6,
                Duration = TimeSpan.FromMilliseconds(2300 + (i * 150)),
                BeginTime = TimeSpan.FromMilliseconds(LifecycleVisuals.SplashPhasePulseStartMs + 180 + (i * 120)),
                RepeatBehavior = RepeatBehavior.Forever,
                EasingFunction = LifecycleVisuals.CreateEaseInOut(),
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(moveX, pulse);
            Storyboard.SetTargetProperty(moveX, "(Canvas.Left)");
            ambientStoryboard.Children.Add(moveX);

            var moveY = new DoubleAnimation
            {
                From = a.Y - 1.6,
                To = b.Y - 1.6,
                Duration = TimeSpan.FromMilliseconds(2300 + (i * 150)),
                BeginTime = TimeSpan.FromMilliseconds(LifecycleVisuals.SplashPhasePulseStartMs + 180 + (i * 120)),
                RepeatBehavior = RepeatBehavior.Forever,
                EasingFunction = LifecycleVisuals.CreateEaseInOut(),
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(moveY, pulse);
            Storyboard.SetTargetProperty(moveY, "(Canvas.Top)");
            ambientStoryboard.Children.Add(moveY);
        }

        for (var i = 0; i < 4; i++)
        {
            var particle = new Ellipse
            {
                Width = i % 2 == 0 ? 2.4 : 2.0,
                Height = i % 2 == 0 ? 2.4 : 2.0,
                Fill = new SolidColorBrush(LifecycleVisuals.ParticleColor),
                Opacity = 0.14,
                RenderTransform = new TranslateTransform()
            };

            Canvas.SetLeft(particle, 248 + (i * 190));
            Canvas.SetTop(particle, 184 + ((i % 2) * 162));
            canvas.Children.Add(particle);

            var driftX = new DoubleAnimation
            {
                From = -3,
                To = 4,
                Duration = TimeSpan.FromMilliseconds(4200 + (i * 320)),
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                EasingFunction = LifecycleVisuals.CreateEaseInOut(),
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(driftX, particle);
            Storyboard.SetTargetProperty(driftX, "(UIElement.RenderTransform).(TranslateTransform.X)");
            ambientStoryboard.Children.Add(driftX);

            var driftY = new DoubleAnimation
            {
                From = -2,
                To = 3,
                Duration = TimeSpan.FromMilliseconds(5200 + (i * 420)),
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                EasingFunction = LifecycleVisuals.CreateEaseInOut(),
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(driftY, particle);
            Storyboard.SetTargetProperty(driftY, "(UIElement.RenderTransform).(TranslateTransform.Y)");
            ambientStoryboard.Children.Add(driftY);
        }

        root.Children.Add(canvas);
    }

    private readonly record struct SplashNodeSpec(Windows.Foundation.Point Point, double Size, string? Label, bool IsPrimary);

    private static Rectangle CreateSplashLightBand(
        double width,
        double height,
        double opacity,
        double rotation,
        double fromX,
        double toX,
        int durationMs,
        int beginMs,
        Storyboard storyboard)
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
                    new GradientStop { Color = Windows.UI.Color.FromArgb(0x00, 0x60, 0xBA, 0xF8), Offset = 0 },
                    new GradientStop { Color = Windows.UI.Color.FromArgb(0xE6, 0x60, 0xBA, 0xF8), Offset = 0.5 },
                    new GradientStop { Color = Windows.UI.Color.FromArgb(0x00, 0x60, 0xBA, 0xF8), Offset = 1 }
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
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(move, band);
        Storyboard.SetTargetProperty(move, "(UIElement.RenderTransform).(CompositeTransform.TranslateX)");
        storyboard.Children.Add(move);

        return band;
    }

    private static void RunStartupSplashAnimation(
        Border glow,
        Border coreGlow,
        Ellipse outerHalo,
        Ellipse halo,
        Ellipse innerRing,
        Ellipse ringWave,
        Image icon,
        Border card,
        Rectangle progressFill,
        Rectangle progressShimmer,
        Storyboard ambientStoryboard,
        Storyboard sequenceStoryboard)
    {
        var cardIntroOpacity = new DoubleAnimation
        {
            From = 0.0,
            To = 1.0,
            Duration = TimeSpan.FromMilliseconds(LifecycleVisuals.SplashPhaseCardInMs),
            BeginTime = TimeSpan.FromMilliseconds(80),
            FillBehavior = FillBehavior.HoldEnd,
            EasingFunction = LifecycleVisuals.CreateEaseOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(cardIntroOpacity, card);
        Storyboard.SetTargetProperty(cardIntroOpacity, "Opacity");
        sequenceStoryboard.Children.Add(cardIntroOpacity);

        var iconFadeIn = new DoubleAnimation
        {
            From = 0.0,
            To = 1.0,
            Duration = TimeSpan.FromMilliseconds(620),
            BeginTime = TimeSpan.FromMilliseconds(240),
            FillBehavior = FillBehavior.HoldEnd,
            EasingFunction = LifecycleVisuals.CreateEaseOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(iconFadeIn, icon);
        Storyboard.SetTargetProperty(iconFadeIn, "Opacity");
        sequenceStoryboard.Children.Add(iconFadeIn);

        var iconScaleInX = new DoubleAnimation
        {
            From = 0.92,
            To = 1.0,
            Duration = TimeSpan.FromMilliseconds(620),
            BeginTime = TimeSpan.FromMilliseconds(240),
            FillBehavior = FillBehavior.HoldEnd,
            EasingFunction = LifecycleVisuals.CreateEaseOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(iconScaleInX, icon);
        Storyboard.SetTargetProperty(iconScaleInX, "(UIElement.RenderTransform).(ScaleTransform.ScaleX)");
        sequenceStoryboard.Children.Add(iconScaleInX);

        var iconScaleInY = new DoubleAnimation
        {
            From = 0.92,
            To = 1.0,
            Duration = TimeSpan.FromMilliseconds(620),
            BeginTime = TimeSpan.FromMilliseconds(240),
            FillBehavior = FillBehavior.HoldEnd,
            EasingFunction = LifecycleVisuals.CreateEaseOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(iconScaleInY, icon);
        Storyboard.SetTargetProperty(iconScaleInY, "(UIElement.RenderTransform).(ScaleTransform.ScaleY)");
        sequenceStoryboard.Children.Add(iconScaleInY);

        var haloScaleX = new DoubleAnimation
        {
            From = 1.0,
            To = 1.06,
            Duration = TimeSpan.FromMilliseconds(1720),
            AutoReverse = true,
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(haloScaleX, halo);
        Storyboard.SetTargetProperty(haloScaleX, "(UIElement.RenderTransform).(ScaleTransform.ScaleX)");
        ambientStoryboard.Children.Add(haloScaleX);

        var haloScaleY = new DoubleAnimation
        {
            From = 1.0,
            To = 1.06,
            Duration = TimeSpan.FromMilliseconds(1720),
            AutoReverse = true,
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(haloScaleY, halo);
        Storyboard.SetTargetProperty(haloScaleY, "(UIElement.RenderTransform).(ScaleTransform.ScaleY)");
        ambientStoryboard.Children.Add(haloScaleY);

        var haloFade = new DoubleAnimation
        {
            From = 0.24,
            To = 0.54,
            Duration = TimeSpan.FromMilliseconds(1720),
            AutoReverse = true,
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(haloFade, halo);
        Storyboard.SetTargetProperty(haloFade, "Opacity");
        ambientStoryboard.Children.Add(haloFade);

        var outerHaloFade = new DoubleAnimation
        {
            From = 0.16,
            To = 0.42,
            Duration = TimeSpan.FromMilliseconds(2400),
            AutoReverse = true,
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(outerHaloFade, outerHalo);
        Storyboard.SetTargetProperty(outerHaloFade, "Opacity");
        ambientStoryboard.Children.Add(outerHaloFade);

        var innerRingFade = new DoubleAnimation
        {
            From = 0.24,
            To = 0.58,
            Duration = TimeSpan.FromMilliseconds(1500),
            AutoReverse = true,
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(innerRingFade, innerRing);
        Storyboard.SetTargetProperty(innerRingFade, "Opacity");
        ambientStoryboard.Children.Add(innerRingFade);

        var ringWaveOpacity = new DoubleAnimation
        {
            From = 0.0,
            To = 0.72,
            Duration = TimeSpan.FromMilliseconds(300),
            BeginTime = TimeSpan.FromMilliseconds(620),
            AutoReverse = true,
            FillBehavior = FillBehavior.Stop,
            EasingFunction = LifecycleVisuals.CreateEaseOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(ringWaveOpacity, ringWave);
        Storyboard.SetTargetProperty(ringWaveOpacity, "Opacity");
        sequenceStoryboard.Children.Add(ringWaveOpacity);

        var ringWaveScaleX = new DoubleAnimation
        {
            From = 0.92,
            To = 1.14,
            Duration = TimeSpan.FromMilliseconds(620),
            BeginTime = TimeSpan.FromMilliseconds(620),
            FillBehavior = FillBehavior.Stop,
            EasingFunction = LifecycleVisuals.CreateEaseOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(ringWaveScaleX, ringWave);
        Storyboard.SetTargetProperty(ringWaveScaleX, "(UIElement.RenderTransform).(ScaleTransform.ScaleX)");
        sequenceStoryboard.Children.Add(ringWaveScaleX);

        var ringWaveScaleY = new DoubleAnimation
        {
            From = 0.92,
            To = 1.14,
            Duration = TimeSpan.FromMilliseconds(620),
            BeginTime = TimeSpan.FromMilliseconds(620),
            FillBehavior = FillBehavior.Stop,
            EasingFunction = LifecycleVisuals.CreateEaseOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(ringWaveScaleY, ringWave);
        Storyboard.SetTargetProperty(ringWaveScaleY, "(UIElement.RenderTransform).(ScaleTransform.ScaleY)");
        sequenceStoryboard.Children.Add(ringWaveScaleY);

        var glowPulse = new DoubleAnimation
        {
            From = 0.14,
            To = 0.30,
            Duration = TimeSpan.FromMilliseconds(1620),
            AutoReverse = true,
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(glowPulse, glow);
        Storyboard.SetTargetProperty(glowPulse, "Opacity");
        ambientStoryboard.Children.Add(glowPulse);

        var coreGlowPulse = new DoubleAnimation
        {
            From = 0.18,
            To = 0.42,
            Duration = TimeSpan.FromMilliseconds(1440),
            AutoReverse = true,
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(coreGlowPulse, coreGlow);
        Storyboard.SetTargetProperty(coreGlowPulse, "Opacity");
        ambientStoryboard.Children.Add(coreGlowPulse);

        var cardPulse = new DoubleAnimation
        {
            From = 0.96,
            To = 1.0,
            Duration = TimeSpan.FromMilliseconds(2200),
            AutoReverse = true,
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(cardPulse, card);
        Storyboard.SetTargetProperty(cardPulse, "Opacity");
        ambientStoryboard.Children.Add(cardPulse);

        var progressSweep = new DoubleAnimationUsingKeyFrames
        {
            BeginTime = TimeSpan.FromMilliseconds(980),
            FillBehavior = FillBehavior.HoldEnd
        };
        progressSweep.KeyFrames.Add(new EasingDoubleKeyFrame { KeyTime = KeyTime.FromTimeSpan(TimeSpan.Zero), Value = 0.10, EasingFunction = LifecycleVisuals.CreateEaseOut() });
        progressSweep.KeyFrames.Add(new EasingDoubleKeyFrame { KeyTime = KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(820)), Value = 0.46, EasingFunction = LifecycleVisuals.CreateEaseInOut() });
        progressSweep.KeyFrames.Add(new EasingDoubleKeyFrame { KeyTime = KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(1560)), Value = 0.74, EasingFunction = LifecycleVisuals.CreateEaseInOut() });
        progressSweep.KeyFrames.Add(new EasingDoubleKeyFrame { KeyTime = KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(2180)), Value = 0.90, EasingFunction = LifecycleVisuals.CreateEaseOut() });
        progressSweep.KeyFrames.Add(new EasingDoubleKeyFrame { KeyTime = KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(2840)), Value = 0.97, EasingFunction = LifecycleVisuals.CreateEaseOut() });
        Storyboard.SetTarget(progressSweep, progressFill);
        Storyboard.SetTargetProperty(progressSweep, "(UIElement.RenderTransform).(ScaleTransform.ScaleX)");
        ambientStoryboard.Children.Add(progressSweep);

        var shimmerMove = new DoubleAnimation
        {
            From = -112,
            To = 386,
            Duration = TimeSpan.FromMilliseconds(2080),
            BeginTime = TimeSpan.FromMilliseconds(1080),
            RepeatBehavior = RepeatBehavior.Forever,
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(shimmerMove, progressShimmer);
        Storyboard.SetTargetProperty(shimmerMove, "(UIElement.RenderTransform).(TranslateTransform.X)");
        ambientStoryboard.Children.Add(shimmerMove);

        sequenceStoryboard.Begin();
        ambientStoryboard.Begin();
    }

    private async Task AnimateStartupStatusChangeAsync(string nextText)
    {
        if (_startupStatusText is null)
        {
            return;
        }

        var fadeOut = new DoubleAnimation
        {
            From = _startupStatusText.Opacity,
            To = 0.22,
            Duration = TimeSpan.FromMilliseconds(170),
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };

        var fadeOutStoryboard = new Storyboard();
        Storyboard.SetTarget(fadeOut, _startupStatusText);
        Storyboard.SetTargetProperty(fadeOut, "Opacity");
        fadeOutStoryboard.Children.Add(fadeOut);
        fadeOutStoryboard.Begin();

        await Task.Delay(180);
        _startupStatusText.Text = nextText;

        var fadeIn = new DoubleAnimation
        {
            From = 0.22,
            To = 0.94,
            Duration = TimeSpan.FromMilliseconds(220),
            EasingFunction = LifecycleVisuals.CreateEaseOut(),
            EnableDependentAnimation = true
        };

        var fadeInStoryboard = new Storyboard();
        Storyboard.SetTarget(fadeIn, _startupStatusText);
        Storyboard.SetTargetProperty(fadeIn, "Opacity");
        fadeInStoryboard.Children.Add(fadeIn);
        fadeInStoryboard.Begin();
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

    private UIElement BuildLayout()
    {
        _navButtons.Clear();
        _vmChipRefreshSignature = string.Empty;

        var root = new Grid
        {
            Background = Application.Current.Resources["PageBackgroundBrush"] as Brush
        };

        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(12) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(12) });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(12) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var headerCard = CreateCard(new Thickness(16, 16, 16, 0), 14, 14);
        var headerGrid = new Grid { RowSpacing = 8 };
        headerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        headerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var titleStack = new StackPanel { Orientation = Orientation.Vertical, Spacing = 2 };
        var titleRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, VerticalAlignment = VerticalAlignment.Center };
        titleRow.Children.Add(new Image
        {
            Source = new Microsoft.UI.Xaml.Media.Imaging.BitmapImage(new Uri("ms-appx:///Assets/HyperTool.ico")),
            Width = 28,
            Height = 28
        });
        titleRow.Children.Add(new TextBlock
        {
            Text = "HyperTool",
            FontSize = 24,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
        });
        titleStack.Children.Add(titleRow);
        titleStack.Children.Add(new TextBlock { Text = "dein nützlicher Hyper V Helfer", Opacity = 0.8, Margin = new Thickness(0, 0, 0, 6) });

        _vmChipScrollViewer.Content = _vmChipPanel;
        _vmChipScrollViewer.ViewChanged += (_, _) => UpdateVmChipNavigationButtons();
        _vmChipScrollViewer.SizeChanged += (_, _) => UpdateVmChipNavigationButtons();
        _vmChipPanel.SizeChanged += (_, _) => UpdateVmChipNavigationButtons();

        _vmChipsLeftButton.Content = new FontIcon { Glyph = "\uE76B", FontSize = 13 };
        _vmChipsLeftButton.Width = 34;
        _vmChipsLeftButton.Height = 34;
        _vmChipsLeftButton.CornerRadius = new CornerRadius(8);
        _vmChipsLeftButton.BorderThickness = new Thickness(1);
        _vmChipsLeftButton.BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush;
        _vmChipsLeftButton.Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush;
        _vmChipsLeftButton.Visibility = Visibility.Collapsed;
        _vmChipsLeftButton.Click += (_, _) => ScrollVmChipsBy(-280);

        _vmChipsRightButton.Content = new FontIcon { Glyph = "\uE76C", FontSize = 13 };
        _vmChipsRightButton.Width = 34;
        _vmChipsRightButton.Height = 34;
        _vmChipsRightButton.CornerRadius = new CornerRadius(8);
        _vmChipsRightButton.BorderThickness = new Thickness(1);
        _vmChipsRightButton.BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush;
        _vmChipsRightButton.Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush;
        _vmChipsRightButton.Visibility = Visibility.Collapsed;
        _vmChipsRightButton.Click += (_, _) => ScrollVmChipsBy(280);

        var chipNavigationGrid = new Grid { ColumnSpacing = 8 };
        chipNavigationGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        chipNavigationGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        chipNavigationGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var chipScrollerHost = new Grid();
        chipScrollerHost.Children.Add(_vmChipScrollViewer);
        chipScrollerHost.Children.Add(_vmChipLeftFadeOverlay);
        chipScrollerHost.Children.Add(_vmChipRightFadeOverlay);

        chipNavigationGrid.Children.Add(_vmChipsLeftButton);
        Grid.SetColumn(chipScrollerHost, 1);
        chipNavigationGrid.Children.Add(chipScrollerHost);
        Grid.SetColumn(_vmChipsRightButton, 2);
        chipNavigationGrid.Children.Add(_vmChipsRightButton);

        Grid.SetRow(chipNavigationGrid, 1);
        Grid.SetColumn(chipNavigationGrid, 0);
        Grid.SetColumnSpan(chipNavigationGrid, 2);
        chipNavigationGrid.Margin = new Thickness(0, 2, 0, 0);
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
        var logoImage = new Image
        {
            Source = new Microsoft.UI.Xaml.Media.Imaging.BitmapImage(new Uri("ms-appx:///Assets/Logo.png")),
            Width = 40,
            Height = 40,
            RenderTransform = _logoRotateTransform,
            RenderTransformOrigin = new Windows.Foundation.Point(0.5, 0.5)
        };
        logoBorder.Child = logoImage;
        logoBorder.Tapped += (_, _) => RunLogoEasterEgg();
        titleActions.Children.Add(logoBorder);

        Grid.SetRow(titleActions, 0);
        Grid.SetColumn(titleActions, 1);
        headerGrid.Children.Add(titleActions);
        headerGrid.Children.Add(chipNavigationGrid);
        headerCard.Child = headerGrid;
        Grid.SetRow(headerCard, 0);
        root.Children.Add(headerCard);

        _configurationNoticeText.Foreground = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, 0xF7, 0xD0, 0x7B));
        _configurationNoticeBorder.Margin = new Thickness(16, 0, 16, 0);
        _configurationNoticeBorder.Padding = new Thickness(12);
        _configurationNoticeBorder.CornerRadius = new CornerRadius(12);
        _configurationNoticeBorder.BorderThickness = new Thickness(1);
        _configurationNoticeBorder.BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush;
        _configurationNoticeBorder.Background = new SolidColorBrush(Windows.UI.Color.FromArgb(0x26, 0xF7, 0xD0, 0x7B));
        _configurationNoticeBorder.Child = _configurationNoticeText;
        Grid.SetRow(_configurationNoticeBorder, 2);
        root.Children.Add(_configurationNoticeBorder);

        var contentBorder = CreateCard(new Thickness(16, 0, 16, 0), 12, 14);
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
        sidebarStack.Children.Add(CreateNavButton("▶", "VM", 0));
        var usbNavButton = CreateNavButton("🔌", "USB-Share", 1);
        _usbNavButton = usbNavButton;
        sidebarStack.Children.Add(usbNavButton);
        var sharedFolderNavButton = CreateNavButton("📁", "Shared Folder", 2);
        _sharedFolderNavButton = sharedFolderNavButton;
        sidebarStack.Children.Add(sharedFolderNavButton);
        sidebarStack.Children.Add(CreateNavButton("📷", "Snapshots", 3));
        sidebarStack.Children.Add(CreateNavButton("⚙", "Einstellungen", 4));
        sidebarStack.Children.Add(CreateNavButton("ℹ", "Info", 5));
        sidebar.Child = sidebarStack;
        mainGrid.Children.Add(sidebar);

        var contentGrid = new Grid();
        contentGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        contentGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(10) });
        contentGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var selectedVmRow = new Border
        {
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            Padding = new Thickness(12)
        };
        var selectedVmGrid = new Grid { ColumnSpacing = 12 };
        selectedVmGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(90) });
        selectedVmGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        selectedVmGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(70) });
        selectedVmGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });
        selectedVmGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(90) });
        selectedVmGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        selectedVmGrid.Children.Add(new TextBlock { Text = "Selected VM", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
        var selectedVmName = new TextBlock();
        selectedVmName.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.SelectedVmDisplayName)) });
        Grid.SetColumn(selectedVmName, 1);
        selectedVmGrid.Children.Add(selectedVmName);
        var stateLabel = new TextBlock { Text = "State", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold };
        Grid.SetColumn(stateLabel, 2);
        selectedVmGrid.Children.Add(stateLabel);
        _selectedVmStateText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.SelectedVmState)) });
        Grid.SetColumn(_selectedVmStateText, 3);
        selectedVmGrid.Children.Add(_selectedVmStateText);
        var networkLabel = new TextBlock { Text = "Network", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold };
        Grid.SetColumn(networkLabel, 4);
        selectedVmGrid.Children.Add(networkLabel);
        var networkText = new TextBlock();
        networkText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.SelectedVmAdapterSwitchDisplay)) });
        Grid.SetColumn(networkText, 5);
        selectedVmGrid.Children.Add(networkText);
        selectedVmRow.Child = selectedVmGrid;
        contentGrid.Children.Add(selectedVmRow);

        _vmPage = BuildVmPage();
        _infoPage = BuildInfoPage();
        UpdateSelectedVmStateTextStyle();
        UpdatePageContent();
        Grid.SetRow(_pageContent, 2);
        contentGrid.Children.Add(_pageContent);
        Grid.SetColumn(contentGrid, 2);
        mainGrid.Children.Add(contentGrid);
        contentBorder.Child = mainGrid;
        Grid.SetRow(contentBorder, 4);
        root.Children.Add(contentBorder);

        var bottomCard = CreateCard(new Thickness(16), 14, 12);
        var bottom = new Grid { RowSpacing = 8 };
        bottom.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        bottom.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        bottom.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        bottom.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var topRow = new Grid { ColumnSpacing = 10 };
        topRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        topRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        topRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        topRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        topRow.Children.Add(new TextBlock { Text = "Notifications", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, VerticalAlignment = VerticalAlignment.Center });
        _busyRing.VerticalAlignment = VerticalAlignment.Center;
        _busyRing.IsActive = false;
        Grid.SetColumn(_busyRing, 1);
        topRow.Children.Add(_busyRing);
        _busyText.VerticalAlignment = VerticalAlignment.Center;
        _busyText.Foreground = Application.Current.Resources["TextMutedBrush"] as Brush;
        Grid.SetColumn(_busyText, 2);
        topRow.Children.Add(_busyText);

        _busyPercentBadge.CornerRadius = new CornerRadius(10);
        _busyPercentBadge.Padding = new Thickness(8, 2, 8, 2);
        _busyPercentBadge.Background = Application.Current.Resources["AccentSoftBrush"] as Brush;
        _busyPercentBadge.BorderThickness = new Thickness(1);
        _busyPercentBadge.BorderBrush = Application.Current.Resources["AccentBrush"] as Brush;
        _busyPercentText.FontWeight = Microsoft.UI.Text.FontWeights.SemiBold;
        _busyPercentText.FontSize = 12;
        _busyPercentBadge.Child = _busyPercentText;
        Grid.SetColumn(_busyPercentBadge, 3);
        topRow.Children.Add(_busyPercentBadge);
        bottom.Children.Add(topRow);

        _busyProgress.Margin = new Thickness(0, 2, 0, 0);
        _busyProgress.Foreground = Application.Current.Resources["AccentStrongBrush"] as Brush ?? Application.Current.Resources["AccentBrush"] as Brush;
        _busyProgress.Background = Application.Current.Resources["AccentSoftBrush"] as Brush;
        Grid.SetRow(_busyProgress, 1);
        bottom.Children.Add(_busyProgress);

        var summaryGrid = new Grid { ColumnSpacing = 8 };
        summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        _notificationSummaryBorder.Padding = new Thickness(10);
        _notificationSummaryBorder.CornerRadius = new CornerRadius(8);
        _notificationSummaryBorder.BorderThickness = new Thickness(1);
        _notificationSummaryBorder.BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush;
        _notificationSummaryBorder.Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush;
        _statusText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.LastNotificationText)) });
        _statusText.TextTrimming = TextTrimming.CharacterEllipsis;
        _notificationSummaryBorder.Child = _statusText;
        summaryGrid.Children.Add(_notificationSummaryBorder);

        var summaryButtons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        summaryButtons.Children.Add(CreateIconButton("📄", "Log-Ordner öffnen", _viewModel.OpenLogFileCommand, compact: true));
        _toggleLogButton.Command = _viewModel.ToggleLogCommand;
        _toggleLogButton.CornerRadius = new CornerRadius(8);
        _toggleLogButton.Padding = new Thickness(8, 2, 8, 2);
        _toggleLogButton.BorderThickness = new Thickness(1);
        _toggleLogButton.BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush;
        _toggleLogButton.Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush;
        summaryButtons.Children.Add(_toggleLogButton);
        Grid.SetColumn(summaryButtons, 1);
        summaryGrid.Children.Add(summaryButtons);
        Grid.SetRow(summaryGrid, 2);
        bottom.Children.Add(summaryGrid);

        _notificationExpandedGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        _notificationExpandedGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(8) });
        _notificationExpandedGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        var expandedButtons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        expandedButtons.Children.Add(CreateIconButton("⧉", "Copy", _viewModel.CopyNotificationsCommand, compact: true));
        expandedButtons.Children.Add(CreateIconButton("⌫", "Clear", _viewModel.ClearNotificationsCommand, compact: true));
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
        _notificationsListView.ItemsSource = _viewModel.Notifications;
        _notificationsListView.DisplayMemberPath = nameof(UiNotification.Message);
        _notificationsListView.MaxHeight = 220;
        logListBorder.Child = _notificationsListView;
        Grid.SetRow(logListBorder, 2);
        _notificationExpandedGrid.Children.Add(logListBorder);
        Grid.SetRow(_notificationExpandedGrid, 3);
        bottom.Children.Add(_notificationExpandedGrid);

        bottomCard.Child = bottom;
        Grid.SetRow(bottomCard, 6);
        root.Children.Add(bottomCard);

        UpdateNoticeVisibility();
        UpdateBusyAndNotificationPanel();
        UpdateNavSelection();

        return root;
    }

    private UIElement BuildVmPage()
    {
        var grid = new Grid { RowSpacing = 10, ColumnSpacing = 8 };
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var actionCard = new Border
        {
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush,
            Padding = new Thickness(10)
        };
        var actionStack = new StackPanel { Spacing = 8 };

        var actionRow1 = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        actionRow1.Children.Add(CreateIconButton("▶", "Start VM", _viewModel.StartSelectedVmCommand));
        actionRow1.Children.Add(CreateIconButton("■", "Stop VM", _viewModel.StopSelectedVmCommand));
        actionRow1.Children.Add(CreateIconButton("⏻", "Hard Off", _viewModel.TurnOffSelectedVmCommand));
        actionRow1.Children.Add(CreateIconButton("↻", "Restart", _viewModel.RestartSelectedVmCommand));
        actionRow1.Children.Add(CreateIconButton("🖥", "Open Console", _viewModel.OpenConsoleCommand));
        actionStack.Children.Add(actionRow1);

        var actionRow2 = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        actionRow2.Children.Add(CreateIconButton("🛠", "Konsole neu aufbauen", _viewModel.ReopenConsoleWithSessionEditCommand));
        actionRow2.Children.Add(CreateIconButton("⟳", "Refresh", _viewModel.RefreshVmStatusCommand));
        actionRow2.Children.Add(CreateIconButton("💾", "Export", _viewModel.ExportSelectedVmCommand));
        actionRow2.Children.Add(CreateIconButton("📥", "Import", _viewModel.ImportVmCommand));
        actionStack.Children.Add(actionRow2);

        actionCard.Child = actionStack;
        grid.Children.Add(actionCard);

        Grid.SetRow(actionRow2, 1);

        var networkSection = new Border
        {
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            Padding = new Thickness(10)
        };
        var networkStack = new StackPanel { Spacing = 8 };

        networkStack.Children.Add(new TextBlock
        {
            Text = "Netzwerk & Adapter",
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
        });

        var switchActions = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        switchActions.Children.Add(CreateIconButton("⟳", "Switches neu laden", _viewModel.RefreshSwitchesCommand));
        networkStack.Children.Add(switchActions);

        networkStack.Children.Add(new TextBlock
        {
            Text = "Switch direkt pro Adapter wählen (aktiver Switch ist hervorgehoben)",
            Opacity = 0.85,
            Margin = new Thickness(0, 2, 0, 0)
        });

        networkStack.Children.Add(_vmAdapterCardsPanel);

        var footerRow = new Grid { ColumnSpacing = 10 };
        footerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        footerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        var hostText = new TextBlock
        {
            Text = $"vmconnect host: {_viewModel.VmConnectComputerName}",
            Opacity = 0.8,
            VerticalAlignment = VerticalAlignment.Center
        };
        footerRow.Children.Add(hostText);
        var hostButtonBottom = CreateIconButton("🌐", "Host Network", onClick: async (_, _) => await OpenHostNetworkWindowAsync(), compact: true);
        Grid.SetColumn(hostButtonBottom, 1);
        footerRow.Children.Add(hostButtonBottom);
        networkStack.Children.Add(footerRow);

        networkSection.Child = networkStack;
        Grid.SetRow(networkSection, 2);
        grid.Children.Add(networkSection);

        return new ScrollViewer { Content = grid };
    }

    private UIElement BuildSnapshotsPage()
    {
        var panel = new Grid { RowSpacing = 10 };
        panel.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var listBorder = new Border
        {
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PageBackgroundBrush"] as Brush,
            Padding = new Thickness(8)
        };

        _checkpointTreeView.SelectionMode = TreeViewSelectionMode.Single;
        _checkpointTreeView.SelectionChanged += (_, args) =>
        {
            if (_isUpdatingCheckpointTreeSelection)
            {
                return;
            }

            var selectedObject = args.AddedItems.FirstOrDefault();
            var selectedItem = selectedObject is TreeViewNode selectedNode
                               && _checkpointItemsByNode.TryGetValue(selectedNode, out var nodeItem)
                ? nodeItem
                : null;

            if (selectedItem is not null)
            {
                _viewModel.SelectedCheckpointNode = selectedItem;
            }
        };

        RefreshCheckpointTreeView();
        listBorder.Child = _checkpointTreeView;
        panel.Children.Add(listBorder);

        var createGrid = new Grid { ColumnSpacing = 8 };
        createGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(130) });
        createGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        createGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        createGrid.Children.Add(new TextBlock { Text = "Checkpoint Name", VerticalAlignment = VerticalAlignment.Center, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });

        var checkpointNameBox = new TextBox { Text = _viewModel.NewCheckpointName, PlaceholderText = "Checkpoint Name", CornerRadius = new CornerRadius(8) };
        checkpointNameBox.TextChanged += (_, _) => _viewModel.NewCheckpointName = checkpointNameBox.Text;
        Grid.SetColumn(checkpointNameBox, 1);
        createGrid.Children.Add(checkpointNameBox);

        var createButton = CreateIconButton("📸", "Create");
        createButton.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(_viewModel.NewCheckpointName))
            {
                _viewModel.NewCheckpointName = $"Checkpoint {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
            }

            if (_viewModel.CreateCheckpointCommand.CanExecute(null))
            {
                _viewModel.CreateCheckpointCommand.Execute(null);
            }
        };
        Grid.SetColumn(createButton, 2);
        createGrid.Children.Add(createButton);
        Grid.SetRow(createGrid, 1);
        panel.Children.Add(createGrid);

        var actionGrid = new Grid { ColumnSpacing = 8 };
        actionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(130) });
        actionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        actionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        actionGrid.Children.Add(new TextBlock { Text = "Beschreibung", VerticalAlignment = VerticalAlignment.Center, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
        var checkpointDescriptionBox = new TextBox { Text = _viewModel.NewCheckpointDescription, PlaceholderText = "Beschreibung", CornerRadius = new CornerRadius(8) };
        checkpointDescriptionBox.TextChanged += (_, _) => _viewModel.NewCheckpointDescription = checkpointDescriptionBox.Text;
        Grid.SetColumn(checkpointDescriptionBox, 1);
        actionGrid.Children.Add(checkpointDescriptionBox);

        var restoreButton = CreateIconButton("✅", "Restore", _viewModel.ApplyCheckpointCommand);
        Grid.SetColumn(restoreButton, 2);
        actionGrid.Children.Add(restoreButton);

        var deleteButton = CreateIconButton("🗑", "Delete", _viewModel.DeleteCheckpointCommand);
        Grid.SetColumn(deleteButton, 3);
        actionGrid.Children.Add(deleteButton);

        var refreshButton = CreateIconButton("⟳", "Refresh", _viewModel.LoadCheckpointsCommand);
        Grid.SetColumn(refreshButton, 4);
        actionGrid.Children.Add(refreshButton);

        Grid.SetRow(actionGrid, 2);
        panel.Children.Add(actionGrid);
        return panel;
    }

    private UIElement BuildConfigPage()
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

        var headingStack = new StackPanel { Spacing = 4 };
        headingStack.Children.Add(new TextBlock
        {
            Text = "Konfiguration",
            FontSize = 20,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
        });
        headingStack.Children.Add(new TextBlock
        {
            Text = "Wichtige Einstellungen übersichtlich und schnell erreichbar.",
            Opacity = 0.9
        });
        headingGrid.Children.Add(headingStack);

        var configHeaderActions = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 8,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Center
        };
        configHeaderActions.Children.Add(CreateIconButton("💾", "Speichern", _viewModel.SaveConfigCommand));
        configHeaderActions.Children.Add(CreateIconButton("⟳", "Neu laden", _viewModel.ReloadConfigCommand));
        configHeaderActions.Children.Add(CreateIconButton(ToolRestartIcon, ToolRestartLabel, onClick: async (_, _) => await RestartHostToolAsync()));
        Grid.SetColumn(configHeaderActions, 1);
        headingGrid.Children.Add(configHeaderActions);

        headingCard.Child = headingGrid;
        root.Children.Add(headingCard);

        var vmSection = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(14)
        };
        var vmStack = new StackPanel { Spacing = 10 };
        vmStack.Children.Add(new TextBlock { Text = "VM-Konfiguration", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, FontSize = 16 });

        var vmGrid = new Grid { ColumnSpacing = 12 };
        vmGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        vmGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var vmOverviewCard = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush,
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(10)
        };
        var vmOverviewStack = new StackPanel { Spacing = 8 };

        var selectedVmText = new TextBlock { TextWrapping = TextWrapping.Wrap };
        selectedVmText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath("SelectedVmForConfig.DisplayLabel") });
        vmOverviewStack.Children.Add(new TextBlock { Text = "Ausgewählte VM", Opacity = 0.9, Foreground = Application.Current.Resources["TextMutedBrush"] as Brush });
        vmOverviewStack.Children.Add(selectedVmText);

        var trayAdapterRow = new Grid { ColumnSpacing = 8, VerticalAlignment = VerticalAlignment.Center };
        trayAdapterRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });
        trayAdapterRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        trayAdapterRow.Children.Add(new TextBlock
        {
            Text = "Tray Adapter",
            VerticalAlignment = VerticalAlignment.Center,
            Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
        });
        var trayAdapterCombo = CreateStyledComboBox();
        trayAdapterCombo.MinHeight = 38;
        trayAdapterCombo.HorizontalAlignment = HorizontalAlignment.Stretch;
        trayAdapterCombo.ItemsSource = _viewModel.AvailableVmTrayAdapterOptions;
        trayAdapterCombo.DisplayMemberPath = nameof(VmTrayAdapterOption.DisplayName);
        trayAdapterCombo.SelectedItem = _viewModel.SelectedVmTrayAdapterOption;
        trayAdapterCombo.SelectionChanged += (_, _) => _viewModel.SelectedVmTrayAdapterOption = trayAdapterCombo.SelectedItem as VmTrayAdapterOption;
        Grid.SetColumn(trayAdapterCombo, 1);
        trayAdapterRow.Children.Add(trayAdapterCombo);
        vmOverviewStack.Children.Add(trayAdapterRow);

        vmOverviewStack.Children.Add(CreateCheckBox(
            "Für diese VM immer mit Sitzungsbearbeitung öffnen",
            () => _viewModel.SelectedVmOpenConsoleWithSessionEdit,
            value => _viewModel.SelectedVmOpenConsoleWithSessionEdit = value));

        vmOverviewCard.Child = vmOverviewStack;
        vmGrid.Children.Add(vmOverviewCard);

        var adapterCard = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush,
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(10)
        };
        var adapterStack = new StackPanel { Spacing = 8 };
        adapterStack.Children.Add(new TextBlock
        {
            Text = "Adapter umbenennen",
            Opacity = 0.95,
            Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
        });

        var renameAdapterCombo = CreateStyledComboBox();
        renameAdapterCombo.MinHeight = 38;
        renameAdapterCombo.HorizontalAlignment = HorizontalAlignment.Stretch;
        renameAdapterCombo.ItemsSource = _viewModel.AvailableVmAdaptersForRename;
        renameAdapterCombo.DisplayMemberPath = nameof(HyperVVmNetworkAdapterInfo.DisplayName);
        renameAdapterCombo.SelectedItem = _viewModel.SelectedVmAdapterForRename;
        renameAdapterCombo.SelectionChanged += (_, _) => _viewModel.SelectedVmAdapterForRename = renameAdapterCombo.SelectedItem as HyperVVmNetworkAdapterInfo;
        adapterStack.Children.Add(renameAdapterCombo);

        var renameGrid = new Grid { ColumnSpacing = 8, VerticalAlignment = VerticalAlignment.Center };
        renameGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        renameGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        var renameTextBox = CreateStyledTextBox(_viewModel.NewVmAdapterName, "Neuer Name");
        renameTextBox.TextChanged += (_, _) => _viewModel.NewVmAdapterName = renameTextBox.Text;
        renameGrid.Children.Add(renameTextBox);
        var renameButton = CreateIconButton("✎", "Umbenennen", _viewModel.RenameVmAdapterCommand);
        Grid.SetColumn(renameButton, 1);
        renameGrid.Children.Add(renameButton);
        adapterStack.Children.Add(renameGrid);

        var validationText = new TextBlock
        {
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 2, 0, 0),
            Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
        };
        validationText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.VmAdapterRenameValidationMessage)) });
        adapterStack.Children.Add(validationText);

        adapterCard.Child = adapterStack;
        Grid.SetColumn(adapterCard, 1);
        vmGrid.Children.Add(adapterCard);
        vmStack.Children.Add(vmGrid);

        var importOptionsCard = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush,
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(10)
        };

        var importOptionsStack = new StackPanel { Spacing = 8 };
        importOptionsStack.Children.Add(new TextBlock
        {
            Text = "VM-Import Optionen",
            Opacity = 0.95,
            Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
        });

        var importModeGrid = new Grid { ColumnSpacing = 8, VerticalAlignment = VerticalAlignment.Center };
        importModeGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
        importModeGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        importModeGrid.Children.Add(new TextBlock
        {
            Text = "Importmodus",
            VerticalAlignment = VerticalAlignment.Center,
            Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
        });
        var importModeCombo = CreateStyledComboBox();
        importModeCombo.MinHeight = 38;
        importModeCombo.ItemsSource = _viewModel.VmImportModeOptions;
        importModeCombo.DisplayMemberPath = nameof(VmImportModeOption.Label);
        importModeCombo.SelectedItem = _viewModel.SelectedVmImportModeOption;
        importModeCombo.SelectionChanged += (_, _) => _viewModel.SelectedVmImportModeOption = importModeCombo.SelectedItem as VmImportModeOption;
        Grid.SetColumn(importModeCombo, 1);
        importModeGrid.Children.Add(importModeCombo);
        importOptionsStack.Children.Add(importModeGrid);

        var importNameGrid = new Grid { ColumnSpacing = 8, VerticalAlignment = VerticalAlignment.Center };
        importNameGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
        importNameGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        importNameGrid.Children.Add(new TextBlock
        {
            Text = "VM-Name (optional)",
            VerticalAlignment = VerticalAlignment.Center,
            Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
        });
        var importVmNameTextBox = CreateStyledTextBox(_viewModel.ImportVmRequestedName, "z. B. Dev-VM");
        importVmNameTextBox.TextChanged += (_, _) => _viewModel.ImportVmRequestedName = importVmNameTextBox.Text;
        Grid.SetColumn(importVmNameTextBox, 1);
        importNameGrid.Children.Add(importVmNameTextBox);
        importOptionsStack.Children.Add(importNameGrid);

        var importFolderGrid = new Grid { ColumnSpacing = 8, VerticalAlignment = VerticalAlignment.Center };
        importFolderGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
        importFolderGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        importFolderGrid.Children.Add(new TextBlock
        {
            Text = "Ordnername (optional)",
            VerticalAlignment = VerticalAlignment.Center,
            Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
        });
        var importFolderNameTextBox = CreateStyledTextBox(_viewModel.ImportVmRequestedFolderName, "z. B. Dev-VM-Import");
        importFolderNameTextBox.TextChanged += (_, _) => _viewModel.ImportVmRequestedFolderName = importFolderNameTextBox.Text;
        Grid.SetColumn(importFolderNameTextBox, 1);
        importFolderGrid.Children.Add(importFolderNameTextBox);
        importOptionsStack.Children.Add(importFolderGrid);

        importOptionsCard.Child = importOptionsStack;
        vmStack.Children.Add(importOptionsCard);

        var vmButtonsGrid = new Grid { ColumnSpacing = 8, Margin = new Thickness(0, 2, 0, 0) };
        vmButtonsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        var importButton = CreateIconButton("📥", "VM importieren", _viewModel.ImportVmCommand);
        importButton.HorizontalAlignment = HorizontalAlignment.Stretch;
        vmButtonsGrid.Children.Add(importButton);

        vmStack.Children.Add(vmButtonsGrid);
        vmSection.Child = vmStack;
        root.Children.Add(vmSection);

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
        quickTogglesGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var trayMenuCheck = CreateCheckBox("Tasktray-Menü aktiv", () => _viewModel.UiEnableTrayMenu, value => _viewModel.UiEnableTrayMenu = value);
        trayMenuCheck.Margin = new Thickness(0);
        Grid.SetColumn(trayMenuCheck, 0);
        Grid.SetRow(trayMenuCheck, 0);
        quickTogglesGrid.Children.Add(trayMenuCheck);

        var openConsoleAfterStartCheck = CreateCheckBox("Beim VM-Start Konsole automatisch öffnen", () => _viewModel.UiOpenConsoleAfterVmStart, value => _viewModel.UiOpenConsoleAfterVmStart = value);
        openConsoleAfterStartCheck.Margin = new Thickness(0);
        Grid.SetColumn(openConsoleAfterStartCheck, 1);
        Grid.SetRow(openConsoleAfterStartCheck, 0);
        quickTogglesGrid.Children.Add(openConsoleAfterStartCheck);

        var startMinCheck = CreateCheckBox("Beim Start minimiert", () => _viewModel.UiStartMinimized, value => _viewModel.UiStartMinimized = value);
        startMinCheck.Margin = new Thickness(0);
        Grid.SetColumn(startMinCheck, 1);
        Grid.SetRow(startMinCheck, 1);
        quickTogglesGrid.Children.Add(startMinCheck);

        var startWithWindowsCheck = CreateCheckBox("Mit Windows starten", () => _viewModel.UiStartWithWindows, value => _viewModel.UiStartWithWindows = value);
        startWithWindowsCheck.Margin = new Thickness(0);
        Grid.SetColumn(startWithWindowsCheck, 0);
        Grid.SetRow(startWithWindowsCheck, 1);
        quickTogglesGrid.Children.Add(startWithWindowsCheck);

        var themeInlineRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 8,
            VerticalAlignment = VerticalAlignment.Center
        };
        themeInlineRow.Children.Add(new TextBlock
        {
            Text = "Dark Mode",
            VerticalAlignment = VerticalAlignment.Center,
            Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
        });

        var themeToggle = new ToggleSwitch
        {
            IsOn = string.Equals(_viewModel.UiTheme, "Dark", StringComparison.OrdinalIgnoreCase),
            OnContent = "An",
            OffContent = "Aus",
            HorizontalAlignment = HorizontalAlignment.Left,
            MinWidth = 86
        };
        themeToggle.Toggled += (_, _) =>
        {
            _viewModel.UiTheme = themeToggle.IsOn ? "Dark" : "Light";
        };
        themeInlineRow.Children.Add(themeToggle);
        Grid.SetColumn(themeInlineRow, 0);
        Grid.SetRow(themeInlineRow, 2);
        quickTogglesGrid.Children.Add(themeInlineRow);

        var updateCheck = CreateCheckBox("Beim Start auf Updates prüfen", () => _viewModel.UpdateCheckOnStartup, value => _viewModel.UpdateCheckOnStartup = value);
        updateCheck.HorizontalAlignment = HorizontalAlignment.Left;
        updateCheck.Margin = new Thickness(0);
        Grid.SetColumn(updateCheck, 1);
        Grid.SetRow(updateCheck, 2);
        quickTogglesGrid.Children.Add(updateCheck);

        systemStack.Children.Add(quickTogglesGrid);

        var hostRow = new Grid { ColumnSpacing = 8, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0) };
        hostRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
        hostRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        hostRow.Children.Add(new TextBlock
        {
            Text = "VMConnect Host",
            VerticalAlignment = VerticalAlignment.Center,
            Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
        });
        var hostTextBox = CreateStyledTextBox(_viewModel.VmConnectComputerName, "z. B. KAI-PC");
        hostTextBox.MinWidth = 420;
        hostTextBox.MaxWidth = 620;
        hostTextBox.HorizontalAlignment = HorizontalAlignment.Left;
        hostTextBox.TextChanged += (_, _) => _viewModel.VmConnectComputerName = hostTextBox.Text;
        Grid.SetColumn(hostTextBox, 1);
        hostRow.Children.Add(hostTextBox);
        systemStack.Children.Add(hostRow);

        systemSection.Child = systemStack;
        root.Children.Add(systemSection);
        return new ScrollViewer { Content = root };
    }

    private void RefreshCheckpointTreeView()
    {
        _checkpointTreeView.RootNodes.Clear();
        _checkpointNodesById.Clear();
        _checkpointItemsByNode.Clear();

        foreach (var root in _viewModel.AvailableCheckpointTree)
        {
            _checkpointTreeView.RootNodes.Add(CreateCheckpointTreeNode(root));
        }

        SelectCheckpointNodeInTree(_viewModel.SelectedCheckpointNode);
    }

    private TreeViewNode CreateCheckpointTreeNode(HyperVCheckpointTreeItem item)
    {
        var node = new TreeViewNode
        {
            Content = BuildCheckpointNodeText(item),
            IsExpanded = true,
            HasUnrealizedChildren = false
        };

        _checkpointItemsByNode[node] = item;

        if (!string.IsNullOrWhiteSpace(item.Checkpoint.Id)
            && !_checkpointNodesById.ContainsKey(item.Checkpoint.Id))
        {
            _checkpointNodesById[item.Checkpoint.Id] = node;
        }

        foreach (var child in item.Children)
        {
            node.Children.Add(CreateCheckpointTreeNode(child));
        }

        return node;
    }

    private static string BuildCheckpointNodeText(HyperVCheckpointTreeItem item)
    {
        var title = string.IsNullOrWhiteSpace(item.Name) ? "(Ohne Namen)" : item.Name;
        var markers = new List<string>();

        if (item.IsCurrent)
        {
            markers.Add("Aktuell");
        }

        if (item.IsLatest)
        {
            markers.Add("Neueste");
        }

        var createdText = item.Created == default ? "-" : item.Created.ToString("dd.MM.yyyy - HH:mm:ss");

        if (markers.Count == 0)
        {
            return $"{title}\n🕒 {createdText}";
        }

        var badgeLine = string.Join("   ", markers.Select(marker =>
            string.Equals(marker, "Aktuell", StringComparison.OrdinalIgnoreCase)
                ? "● Aktuell"
                : "◆ Neueste"));
        return $"{title}\n{badgeLine}\n🕒 {createdText}";
    }

    private void SelectCheckpointNodeInTree(HyperVCheckpointTreeItem? item)
    {
        _isUpdatingCheckpointTreeSelection = true;
        try
        {
            if (item is null || string.IsNullOrWhiteSpace(item.Checkpoint.Id))
            {
                _checkpointTreeView.SelectedNode = null;
                return;
            }

            if (_checkpointNodesById.TryGetValue(item.Checkpoint.Id, out var node))
            {
                _checkpointTreeView.SelectedNode = node;
            }
            else
            {
                _checkpointTreeView.SelectedNode = null;
            }
        }
        finally
        {
            _isUpdatingCheckpointTreeSelection = false;
        }
    }

    private UIElement BuildInfoPage()
    {
        var panel = new StackPanel { Spacing = 10 };

        var titleWrap = new Grid();
        titleWrap.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        titleWrap.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var heading = new TextBlock { Text = "Info", FontSize = 20, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold };
        titleWrap.Children.Add(heading);

        var versionWrap = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, VerticalAlignment = VerticalAlignment.Bottom };
        versionWrap.Children.Add(new TextBlock { Text = "Version:", Opacity = 0.9, VerticalAlignment = VerticalAlignment.Bottom });
        var versionText = new TextBlock { Opacity = 0.9 };
        versionText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.AppVersion)) });
        versionWrap.Children.Add(versionText);
        Grid.SetColumn(versionWrap, 1);
        titleWrap.Children.Add(versionWrap);

        panel.Children.Add(titleWrap);

        var infoStatusRow = new Grid { ColumnSpacing = 8 };
        infoStatusRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        infoStatusRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var updateWrap = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, VerticalAlignment = VerticalAlignment.Center };
        updateWrap.Children.Add(new TextBlock { Text = "Update-Status:", Opacity = 0.9 });
        var updateText = new TextBlock { TextWrapping = TextWrapping.Wrap, Opacity = 0.9 };
        updateText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.UpdateStatus)) });
        updateWrap.Children.Add(updateText);

        var copyrightText = new TextBlock { Text = "Copyright: koerby", Opacity = 0.9, HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Center };

        Grid.SetColumn(updateWrap, 0);
        infoStatusRow.Children.Add(updateWrap);
        Grid.SetColumn(copyrightText, 1);
        infoStatusRow.Children.Add(copyrightText);
        panel.Children.Add(infoStatusRow);

        var infoCard = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PageBackgroundBrush"] as Brush,
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(10)
        };

        var infoStack = new StackPanel { Spacing = 4 };
        infoStack.Children.Add(new TextBlock { Text = "HyperTool Projekt", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
        infoStack.Children.Add(new TextBlock { Text = "HyperTool wird über GitHub Releases verteilt. Hier findest du Version, Update-Status und Release-Links.", TextWrapping = TextWrapping.Wrap, Opacity = 0.85 });

        var linksGrid = new Grid { ColumnSpacing = 8, RowSpacing = 4 };
        linksGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
        linksGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        linksGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        linksGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        linksGrid.Children.Add(new TextBlock { Text = "GitHub Owner", Opacity = 0.9 });
        var ownerText = new TextBlock();
        ownerText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.GithubOwner)) });
        Grid.SetColumn(ownerText, 1);
        linksGrid.Children.Add(ownerText);
        var repoLabel = new TextBlock { Text = "GitHub Repo", Opacity = 0.9 };
        Grid.SetRow(repoLabel, 1);
        linksGrid.Children.Add(repoLabel);
        var repoText = new TextBlock();
        repoText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.GithubRepo)) });
        Grid.SetColumn(repoText, 1);
        Grid.SetRow(repoText, 1);
        linksGrid.Children.Add(repoText);
        infoStack.Children.Add(linksGrid);
        infoCard.Child = infoStack;
        panel.Children.Add(infoCard);

        var usbipdCard = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PageBackgroundBrush"] as Brush,
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(10)
        };

        var usbipdInfoStack = new StackPanel { Spacing = 4 };
        usbipdInfoStack.Children.Add(new TextBlock { Text = "Externe USB/IP Quelle", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
        usbipdInfoStack.Children.Add(new TextBlock { Text = "Quelle: dorssel/usbipd-win", Opacity = 0.9 });
        usbipdInfoStack.Children.Add(new TextBlock { Text = "Nutzung in HyperTool: externer CLI-/Dienst-Stack für USB-Funktionen in der Host-App.", TextWrapping = TextWrapping.Wrap, Opacity = 0.85 });
        usbipdInfoStack.Children.Add(new TextBlock { Text = "Lizenz/Eigentümer: siehe Original-Repository von dorssel.", TextWrapping = TextWrapping.Wrap, Opacity = 0.85 });
        usbipdCard.Child = usbipdInfoStack;
        panel.Children.Add(usbipdCard);

        var diagnosticsCard = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PageBackgroundBrush"] as Brush,
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(10)
        };

        var diagnosticsStack = new StackPanel { Spacing = 4 };
        diagnosticsStack.Children.Add(new TextBlock { Text = "Transport Diagnose", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });

        var hyperVSocketRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        hyperVSocketRow.Children.Add(new TextBlock { Text = "Hyper-V Socket aktiv:", Opacity = 0.9 });
        var hyperVSocketText = new TextBlock { Opacity = 0.9 };
        hyperVSocketText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.UsbDiagnosticsHyperVSocketText)) });
        hyperVSocketRow.Children.Add(hyperVSocketText);
        diagnosticsStack.Children.Add(hyperVSocketRow);

        var registryRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        registryRow.Children.Add(new TextBlock { Text = "Registry-Service vorhanden:", Opacity = 0.9 });
        var registryText = new TextBlock { Opacity = 0.9 };
        registryText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.UsbDiagnosticsRegistryServiceText)) });
        registryRow.Children.Add(registryText);
        diagnosticsStack.Children.Add(registryRow);

        diagnosticsCard.Child = diagnosticsStack;
        panel.Children.Add(diagnosticsCard);

        var buttonRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        buttonRow.Children.Add(CreateIconButton("🛰", "Update prüfen", _viewModel.CheckForUpdatesCommand));
        buttonRow.Children.Add(CreateIconButton("⬇", "Update installieren", _viewModel.InstallUpdateCommand));
        buttonRow.Children.Add(CreateIconButton("🌐", "Changelog / Release", _viewModel.OpenReleasePageCommand));
        buttonRow.Children.Add(CreateIconButton("🔗", "usbipd-win Quelle", onClick: (_, _) =>
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "https://github.com/dorssel/usbipd-win",
                UseShellExecute = true
            });
        }));
        panel.Children.Add(buttonRow);

        return new ScrollViewer { Content = panel };
    }

    private UIElement BuildUsbPage()
    {
        var root = new Grid { RowSpacing = 8 };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var actionsCard = new Border
        {
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush,
            Padding = new Thickness(8)
        };

        var actionsStack = new StackPanel { Spacing = 4 };

        var headerRow = new Grid { ColumnSpacing = 10 };
        headerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var headerLeft = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 14,
            VerticalAlignment = VerticalAlignment.Center
        };

        _usbHostFeatureEnabledToggleSwitch.Toggled += async (_, _) => await OnHostUsbFeatureToggleChangedAsync(_usbHostFeatureEnabledToggleSwitch.IsOn);
        _usbHostFeatureEnabledToggleSwitch.IsOn = _viewModel.HostUsbSharingEnabled;
        _usbHostFeatureEnabledToggleSwitch.MinWidth = 54;
        _usbHostFeatureEnabledToggleSwitch.VerticalAlignment = VerticalAlignment.Center;
        headerLeft.Children.Add(_usbHostFeatureEnabledToggleSwitch);

        headerLeft.Children.Add(new TextBlock
        {
            Text = "USB-Share (USB/IPD als Dienst im Hintergrund)",
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            VerticalAlignment = VerticalAlignment.Center
        });
        headerRow.Children.Add(headerLeft);

        _usbHostFeatureStatusChip.Child = _usbHostFeatureStatusChipText;
        Grid.SetColumn(_usbHostFeatureStatusChip, 1);
        headerRow.Children.Add(_usbHostFeatureStatusChip);
        actionsStack.Children.Add(headerRow);

        _usbFeatureControlsPanel = new StackPanel { Spacing = 4 };

        _usbRuntimeInstallButton = CreateIconButton("⬇", "Installation usbip-win", onClick: async (_, _) => await InstallHostUsbRuntimeAsync());
        _usbRuntimeRestartButton = CreateIconButton(ToolRestartIcon, ToolRestartLabel, onClick: async (_, _) => await RestartHostToolAsync());
        var usbRuntimeActionsRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        usbRuntimeActionsRow.Children.Add(_usbRuntimeInstallButton);
        usbRuntimeActionsRow.Children.Add(_usbRuntimeRestartButton);
        actionsStack.Children.Add(usbRuntimeActionsRow);
        _usbRuntimeHintText.Foreground = Application.Current.Resources["TextMutedBrush"] as Brush;
        _usbRuntimeHintText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.UsbRuntimeHintText)) });

        var actionRow = new Grid { ColumnSpacing = 8, VerticalAlignment = VerticalAlignment.Center };
        actionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        actionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        _usbRefreshButton = CreateIconButton("⟳", "Refresh", _viewModel.RefreshUsbDevicesCommand);
        _usbShareButton = CreateIconButton("🔓", "Share", _viewModel.BindUsbDeviceCommand);
        _usbUnshareButton = CreateIconButton("🔒", "Unshare", _viewModel.UnbindUsbDeviceCommand);

        Grid.SetColumn(_usbRefreshButton, 0);
        actionRow.Children.Add(_usbRefreshButton);
        Grid.SetColumn(_usbShareButton, 1);
        actionRow.Children.Add(_usbShareButton);
        Grid.SetColumn(_usbUnshareButton, 2);
        actionRow.Children.Add(_usbUnshareButton);

        _usbFeatureControlsPanel.Children.Add(actionRow);

        _usbAutoShareCheckBox.Content = "Auto-Share für ausgewähltes Gerät";
        _usbAutoShareCheckBox.Margin = new Thickness(6, 0, 0, 0);
        _usbAutoShareCheckBox.VerticalAlignment = VerticalAlignment.Center;
        _usbAutoShareCheckBox.HorizontalAlignment = HorizontalAlignment.Left;
        _usbAutoShareCheckBox.IsChecked = _viewModel.SelectedUsbDeviceAutoShareEnabled;
        _usbAutoShareCheckBox.IsEnabled = _viewModel.SelectedUsbDevice is not null && _viewModel.UsbRuntimeAvailable;
        _usbAutoShareCheckBox.Checked += (_, _) => _viewModel.SelectedUsbDeviceAutoShareEnabled = true;
        _usbAutoShareCheckBox.Unchecked += (_, _) => _viewModel.SelectedUsbDeviceAutoShareEnabled = false;
        Grid.SetColumn(_usbAutoShareCheckBox, 3);
        actionRow.Children.Add(_usbAutoShareCheckBox);

        _usbAutoDetachOnDisconnectCheckBox.Content = "Automatisches Detach nach Disconnect";
        _usbAutoDetachOnDisconnectCheckBox.Margin = new Thickness(6, 0, 0, 0);
        _usbAutoDetachOnDisconnectCheckBox.VerticalAlignment = VerticalAlignment.Center;
        _usbAutoDetachOnDisconnectCheckBox.HorizontalAlignment = HorizontalAlignment.Left;
        _usbAutoDetachOnDisconnectCheckBox.IsChecked = _viewModel.UsbAutoDetachOnClientDisconnect;
        _usbAutoDetachOnDisconnectCheckBox.Checked += (_, _) => _viewModel.UsbAutoDetachOnClientDisconnect = true;
        _usbAutoDetachOnDisconnectCheckBox.Unchecked += (_, _) => _viewModel.UsbAutoDetachOnClientDisconnect = false;
        Grid.SetColumn(_usbAutoDetachOnDisconnectCheckBox, 4);
        actionRow.Children.Add(_usbAutoDetachOnDisconnectCheckBox);

        _usbUnshareOnExitCheckBox.Content = "Beim Beenden Share aufheben";
        _usbUnshareOnExitCheckBox.Margin = new Thickness(6, 0, 0, 0);
        _usbUnshareOnExitCheckBox.VerticalAlignment = VerticalAlignment.Center;
        _usbUnshareOnExitCheckBox.HorizontalAlignment = HorizontalAlignment.Left;
        _usbUnshareOnExitCheckBox.IsChecked = _viewModel.UsbUnshareOnExit;
        _usbUnshareOnExitCheckBox.Checked += (_, _) => _viewModel.UsbUnshareOnExit = true;
        _usbUnshareOnExitCheckBox.Unchecked += (_, _) => _viewModel.UsbUnshareOnExit = false;
        Grid.SetColumn(_usbUnshareOnExitCheckBox, 5);
        actionRow.Children.Add(_usbUnshareOnExitCheckBox);

        _usbFeatureControlsPanel.Children.Add(_usbRuntimeHintText);

        _usbRemoteFxPolicyHintText.Foreground = Application.Current.Resources["TextMutedBrush"] as Brush;
        _usbRemoteFxPolicyHintText.Text = "Hinweis: Für USB-Share muss 'RemoteFX USB Device Redirection' in den Richtlinien deaktiviert sein.";
        _usbRemoteFxPolicyHintText.Visibility = Visibility.Visible;
        _usbFeatureControlsPanel.Children.Add(_usbRemoteFxPolicyHintText);
        actionsStack.Children.Add(_usbFeatureControlsPanel);

        UpdateUsbRuntimeStatusUi();
        UpdateHostFeatureAvailabilityUi();

        actionsCard.Child = actionsStack;
        Grid.SetRow(actionsCard, 0);
        root.Children.Add(actionsCard);

        var listCard = new Border
        {
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            Padding = new Thickness(8)
        };

        _usbDevicesListView.ItemsSource = _viewModel.UsbDevices;
        _usbDevicesListView.ItemTemplate = CreateUsbDeviceListItemTemplate();
        _usbDevicesListView.SelectionMode = ListViewSelectionMode.Single;
        _usbDevicesListView.IsItemClickEnabled = true;
        _usbDevicesListView.ItemClick += (_, args) =>
        {
            if (args.ClickedItem is UsbIpDeviceInfo usbDevice)
            {
                _viewModel.SelectedUsbDevice = usbDevice;
            }
        };
        _usbDevicesListView.SelectionChanged += (_, _) =>
        {
            _viewModel.SelectedUsbDevice = _usbDevicesListView.SelectedItem as UsbIpDeviceInfo;
        };

        var listLayout = new Grid { RowSpacing = 6 };
        listLayout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        listLayout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var headerGrid = new Grid { ColumnSpacing = 10, Margin = new Thickness(6, 2, 8, 2) };
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(220) });

        headerGrid.Children.Add(new TextBlock
        {
            Text = "Device",
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            Opacity = 0.9
        });

        var connectedByHeader = new TextBlock
        {
            Text = "Connected by",
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            Opacity = 0.9,
            TextAlignment = TextAlignment.Left
        };
        Grid.SetColumn(connectedByHeader, 1);
        headerGrid.Children.Add(connectedByHeader);

        listLayout.Children.Add(headerGrid);

        _usbDisabledOverlayText.Foreground = Application.Current.Resources["TextMutedBrush"] as Brush;
        var listBodyGrid = new Grid();
        listBodyGrid.Children.Add(_usbDevicesListView);
        listBodyGrid.Children.Add(_usbDisabledOverlayText);

        Grid.SetRow(listBodyGrid, 1);
        listLayout.Children.Add(listBodyGrid);

        listCard.Child = listLayout;
        Grid.SetRow(listCard, 1);
        root.Children.Add(listCard);

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

        var editorStack = new StackPanel { Spacing = 8 };

        var headerRow = new Grid { ColumnSpacing = 10 };
        headerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var headerLeft = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 14,
            VerticalAlignment = VerticalAlignment.Center
        };

        _sharedFolderHostFeatureEnabledToggleSwitch.Toggled += async (_, _) => await OnHostSharedFolderFeatureToggleChangedAsync(_sharedFolderHostFeatureEnabledToggleSwitch.IsOn);
        _sharedFolderHostFeatureEnabledToggleSwitch.IsOn = _viewModel.HostSharedFoldersEnabled;
        _sharedFolderHostFeatureEnabledToggleSwitch.MinWidth = 54;
        _sharedFolderHostFeatureEnabledToggleSwitch.VerticalAlignment = VerticalAlignment.Center;
        headerLeft.Children.Add(_sharedFolderHostFeatureEnabledToggleSwitch);

        var titleText = new TextBlock
        {
            Text = "Shared Folder (Netzlaufwerk im Guest-System per WinFsp)",
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            VerticalAlignment = VerticalAlignment.Center
        };
        headerLeft.Children.Add(titleText);
        headerRow.Children.Add(headerLeft);

        var statusStack = new StackPanel { Spacing = 4, HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Center };
        _sharedFolderHostFeatureStatusChip.Child = _sharedFolderHostFeatureStatusChipText;
        statusStack.Children.Add(_sharedFolderHostFeatureStatusChip);

        Grid.SetColumn(statusStack, 1);
        headerRow.Children.Add(statusStack);

        editorStack.Children.Add(headerRow);

        _sharedFolderPathTextBox.PlaceholderText = "Lokaler Ordnerpfad (z. B. C:\\VMShare\\Tools)";
        _sharedFolderShareNameTextBox.PlaceholderText = "Share-Kennung (z. B. HyperToolTools)";

        var fieldsRow = new Grid { ColumnSpacing = 8 };
        fieldsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(220) });
        fieldsRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        _sharedFolderShareNameTextBox.MinWidth = 180;
        fieldsRow.Children.Add(_sharedFolderShareNameTextBox);

        var pathRow = new Grid { ColumnSpacing = 8 };
        pathRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        pathRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        _sharedFolderPathTextBox.MinWidth = 260;
        pathRow.Children.Add(_sharedFolderPathTextBox);
        var browseFolderButton = CreateIconButton("📂", "Ordner wählen", onClick: async (_, _) => await BrowseSharedFolderPathAsync());
        Grid.SetColumn(browseFolderButton, 1);
        pathRow.Children.Add(browseFolderButton);

        Grid.SetColumn(pathRow, 1);
        fieldsRow.Children.Add(pathRow);
        editorStack.Children.Add(fieldsRow);

        var actionRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        actionRow.Children.Add(_sharedFolderReadOnlyCheckBox);
        _sharedFolderNewButton = CreateIconButton("✚", "Neu", onClick: (_, _) => ResetSharedFolderEditor());
        _sharedFolderSaveButton = CreateIconButton("💾", "Speichern", onClick: async (_, _) => await SaveSharedFolderEntryAsync());
        _sharedFolderDeleteButton = CreateIconButton("🗑", "Entfernen", onClick: async (_, _) => await DeleteSharedFolderEntryAsync());
        _sharedFolderFileServiceInstallButton = CreateIconButton("⬇", "hypertool-file nachinstallieren", onClick: async (_, _) => await InstallHostFileServiceRuntimeAsync());
        actionRow.Children.Add(_sharedFolderNewButton);
        actionRow.Children.Add(_sharedFolderSaveButton);
        actionRow.Children.Add(_sharedFolderDeleteButton);
        actionRow.Children.Add(_sharedFolderFileServiceInstallButton);
        editorStack.Children.Add(actionRow);

        var fileServiceAvailable = HyperVSocketUsbHostTunnel.IsServiceRegistered(HyperVSocketUsbTunnelDefaults.FileServiceId);
        if (!fileServiceAvailable)
        {
            editorStack.Children.Add(new TextBlock
            {
                Text = "HyperTool File Service ist nicht registriert. Shared-Folder können erst nach Nachinstallation genutzt werden.",
                Opacity = 0.9,
                TextWrapping = TextWrapping.Wrap,
                Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
            });
        }

        if (_sharedFolderFileServiceInstallButton is not null)
        {
            _sharedFolderFileServiceInstallButton.Visibility = fileServiceAvailable ? Visibility.Collapsed : Visibility.Visible;
        }

        editorCard.Child = editorStack;
        Grid.SetRow(editorCard, 0);
        root.Children.Add(editorCard);

        UpdateSharedFolderFileServiceStatus(socketActive: false, lastActivityUtc: null);

        var listCard = new Border
        {
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            Padding = new Thickness(8)
        };

        _sharedFoldersListView.ItemsSource = _viewModel.HostSharedFolders;
        _sharedFoldersListView.ItemTemplate = CreateSharedFolderListItemTemplate();
        _sharedFoldersListView.SelectionMode = ListViewSelectionMode.Single;
        _sharedFoldersListView.IsItemClickEnabled = true;
        AttachSharedFolderCollectionHandlers();
        _sharedFoldersListView.ItemClick += (_, args) =>
        {
            if (args.ClickedItem is HostSharedFolderDefinition definition)
            {
                _sharedFoldersListView.SelectedItem = definition;
                LoadSharedFolderEditor(definition);
            }
        };
        _sharedFoldersListView.SelectionChanged += (_, _) =>
        {
            if (_sharedFoldersListView.SelectedItem is HostSharedFolderDefinition definition)
            {
                LoadSharedFolderEditor(definition);
            }
        };

        _sharedFolderDisabledOverlayText.Foreground = Application.Current.Resources["TextMutedBrush"] as Brush;
        var listBodyGrid = new Grid();
        listBodyGrid.Children.Add(_sharedFoldersListView);
        listBodyGrid.Children.Add(_sharedFolderDisabledOverlayText);

        var fileServiceOverlay = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Bottom,
            Margin = new Thickness(8, 8, 10, 6),
            Opacity = 0.78,
            IsHitTestVisible = false
        };
        fileServiceOverlay.Children.Add(_sharedFolderFileServiceStatusDot);
        fileServiceOverlay.Children.Add(_sharedFolderFileServiceStatusText);
        listBodyGrid.Children.Add(fileServiceOverlay);

        listCard.Child = listBodyGrid;
        Grid.SetRow(listCard, 1);
        root.Children.Add(listCard);

        if (_viewModel.HostSharedFolders.Count > 0)
        {
            _sharedFoldersListView.SelectedItem = _viewModel.HostSharedFolders[0];
            LoadSharedFolderEditor(_viewModel.HostSharedFolders[0]);
        }
        else
        {
            ResetSharedFolderEditor();
        }

        UpdateHostFeatureAvailabilityUi();

        return root;
    }

    public void UpdateSharedFolderFileServiceStatus(bool socketActive, DateTimeOffset? lastActivityUtc)
    {
        var lastActivityText = lastActivityUtc.HasValue
            ? lastActivityUtc.Value.ToLocalTime().ToString("HH:mm:ss")
            : "-";

        if (!socketActive)
        {
            _sharedFolderFileServiceStatusDot.Fill = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, 0xE8, 0x4A, 0x5F));
            _sharedFolderFileServiceStatusText.Text = "File Service inaktiv";
            return;
        }

        if (!lastActivityUtc.HasValue)
        {
            _sharedFolderFileServiceStatusDot.Fill = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, 0xF6, 0xC3, 0x44));
            _sharedFolderFileServiceStatusText.Text = "File Service bereit";
            return;
        }

        _sharedFolderFileServiceStatusDot.Fill = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, 0x32, 0xD7, 0x4B));
        _sharedFolderFileServiceStatusText.Text = $"File Service aktiv ({lastActivityText})";
    }

    private async Task InstallHostFileServiceRuntimeAsync()
    {
        if (_sharedFolderFileServiceInstallButton is not null)
        {
            _sharedFolderFileServiceInstallButton.IsEnabled = false;
        }

        try
        {
            var exePath = Environment.ProcessPath;
            if (string.IsNullOrWhiteSpace(exePath) || !File.Exists(exePath))
            {
                _sharedFolderStatusText.Text = "Nachinstallation fehlgeschlagen: Host-Executable nicht gefunden.";
                return;
            }

            _sharedFolderStatusText.Text = "HyperTool File Service wird nachinstalliert...";

            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = "--register-sharedfolder-socket",
                UseShellExecute = true,
                Verb = "runas"
            });

            if (process is null)
            {
                _sharedFolderStatusText.Text = "Nachinstallation fehlgeschlagen: Helper konnte nicht gestartet werden.";
                return;
            }

            await process.WaitForExitAsync();

            var installed = HyperVSocketUsbHostTunnel.IsServiceRegistered(HyperVSocketUsbTunnelDefaults.FileServiceId);
            if (installed)
            {
                _sharedFolderStatusText.Text = "HyperTool File Service erfolgreich installiert. Shared-Folder sind bereit.";
                if (_sharedFolderFileServiceInstallButton is not null)
                {
                    _sharedFolderFileServiceInstallButton.Visibility = Visibility.Collapsed;
                }
            }
            else
            {
                _sharedFolderStatusText.Text = "Nachinstallation beendet, Registry-Eintrag fehlt weiterhin. Bitte als Administrator erneut versuchen.";
            }
        }
        catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
        {
            _sharedFolderStatusText.Text = "Nachinstallation abgebrochen (UAC).";
        }
        catch (Exception ex)
        {
            _sharedFolderStatusText.Text = $"Nachinstallation fehlgeschlagen: {ex.Message}";
        }
        finally
        {
            if (_sharedFolderFileServiceInstallButton is not null)
            {
                _sharedFolderFileServiceInstallButton.IsEnabled = true;
            }
        }
    }

    private static CheckBox CreateCheckBox(string text, Func<bool> getter, Action<bool> setter)
    {
        var checkBox = new CheckBox
        {
            IsChecked = getter(),
            Padding = new Thickness(2, 2, 2, 2),
            MinHeight = 28,
            VerticalAlignment = VerticalAlignment.Center,
            Content = new TextBlock
            {
                Text = text,
                Margin = new Thickness(6, 0, 0, 0),
                Foreground = Application.Current.Resources["TextPrimaryBrush"] as Brush,
                FontWeight = Microsoft.UI.Text.FontWeights.Medium,
                TextWrapping = TextWrapping.Wrap
            }
        };
        checkBox.Checked += (_, _) => setter(true);
        checkBox.Unchecked += (_, _) => setter(false);
        return checkBox;
    }

    private static DataTemplate CreateUsbDeviceListItemTemplate()
    {
                const string templateXaml = """
<DataTemplate xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'>
    <Grid ColumnSpacing='10' Margin='4,2,4,2'>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width='*'/>
            <ColumnDefinition Width='220'/>
        </Grid.ColumnDefinitions>
        <TextBlock Text='{Binding DeviceDisplayName}' TextTrimming='CharacterEllipsis' TextWrapping='NoWrap' VerticalAlignment='Center'/>
        <TextBlock Grid.Column='1' Text='{Binding ConnectedByDisplay}' Opacity='0.9' TextTrimming='CharacterEllipsis' TextWrapping='NoWrap' VerticalAlignment='Center'/>
    </Grid>
</DataTemplate>
""";

                return (DataTemplate)XamlReader.Load(templateXaml);
    }

    private static DataTemplate CreateSharedFolderListItemTemplate()
    {
                const string templateXaml = """
<DataTemplate xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'>
    <Grid ColumnSpacing='10' Margin='4,2,4,2'>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width='34'/>
            <ColumnDefinition Width='220'/>
            <ColumnDefinition Width='*'/>
            <ColumnDefinition Width='170'/>
        </Grid.ColumnDefinitions>
        <CheckBox IsChecked='{Binding Enabled, Mode=TwoWay}' VerticalAlignment='Center' HorizontalAlignment='Left'/>
        <TextBlock Grid.Column='1' Text='{Binding Label}' FontWeight='SemiBold' TextTrimming='CharacterEllipsis' TextWrapping='NoWrap' VerticalAlignment='Center'/>
        <TextBlock Grid.Column='2' Text='{Binding LocalPath}' Opacity='0.78' TextTrimming='CharacterEllipsis' TextWrapping='NoWrap' VerticalAlignment='Center'/>
        <TextBlock Grid.Column='3' Text='{Binding AccessModeLabel}' Opacity='0.9' TextTrimming='CharacterEllipsis' TextWrapping='NoWrap' VerticalAlignment='Center'/>
    </Grid>
</DataTemplate>
""";

                return (DataTemplate)XamlReader.Load(templateXaml);
    }

    private void ResetSharedFolderEditor()
    {
        _sharedFolderEditingId = string.Empty;
        _sharedFolderPathTextBox.Text = string.Empty;
        _sharedFolderShareNameTextBox.Text = string.Empty;
        _sharedFolderReadOnlyCheckBox.IsChecked = false;
        _sharedFolderStatusText.Text = "Neuer Shared-Folder Eintrag.";
    }

    private void LoadSharedFolderEditor(HostSharedFolderDefinition definition)
    {
        _sharedFolderEditingId = definition.Id;
        _sharedFolderPathTextBox.Text = definition.LocalPath;
        _sharedFolderShareNameTextBox.Text = definition.ShareName;
        _sharedFolderReadOnlyCheckBox.IsChecked = definition.ReadOnly;
        _sharedFolderStatusText.Text = $"Eintrag '{definition.ShareName}' ausgewählt.";
    }

    private HostSharedFolderDefinition? BuildSharedFolderFromEditor()
    {
        var shareName = (_sharedFolderShareNameTextBox.Text ?? string.Empty).Trim();
        var localPath = (_sharedFolderPathTextBox.Text ?? string.Empty).Trim();

        if (string.IsNullOrWhiteSpace(shareName))
        {
            _sharedFolderStatusText.Text = "Share-Kennung ist erforderlich.";
            return null;
        }

        if (string.IsNullOrWhiteSpace(localPath))
        {
            _sharedFolderStatusText.Text = "Lokaler Ordnerpfad ist erforderlich.";
            return null;
        }

        return new HostSharedFolderDefinition
        {
            Id = string.IsNullOrWhiteSpace(_sharedFolderEditingId)
                ? Guid.NewGuid().ToString("N")
                : _sharedFolderEditingId,
            Label = shareName,
            LocalPath = localPath,
            ShareName = shareName,
            Enabled = ResolveExistingSharedFolderEnabledState(),
            ReadOnly = _sharedFolderReadOnlyCheckBox.IsChecked == true
        };
    }

    private bool ResolveExistingSharedFolderEnabledState()
    {
        if (string.IsNullOrWhiteSpace(_sharedFolderEditingId))
        {
            return true;
        }

        var existing = _viewModel.HostSharedFolders.FirstOrDefault(item => string.Equals(item.Id, _sharedFolderEditingId, StringComparison.OrdinalIgnoreCase));
        return existing?.Enabled ?? true;
    }

    private void AttachSharedFolderCollectionHandlers()
    {
        if (_sharedFolderCollectionHandlersAttached)
        {
            return;
        }

        _viewModel.HostSharedFolders.CollectionChanged += OnHostSharedFoldersCollectionChanged;
        foreach (var definition in _viewModel.HostSharedFolders)
        {
            definition.PropertyChanged += OnHostSharedFolderDefinitionPropertyChanged;
        }

        _sharedFolderCollectionHandlersAttached = true;
    }

    private void OnHostSharedFoldersCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.OldItems is not null)
        {
            foreach (var item in e.OldItems.OfType<HostSharedFolderDefinition>())
            {
                item.PropertyChanged -= OnHostSharedFolderDefinitionPropertyChanged;
            }
        }

        if (e.NewItems is not null)
        {
            foreach (var item in e.NewItems.OfType<HostSharedFolderDefinition>())
            {
                item.PropertyChanged += OnHostSharedFolderDefinitionPropertyChanged;
            }
        }
    }

    private async void OnHostSharedFolderDefinitionPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (!string.Equals(e.PropertyName, nameof(HostSharedFolderDefinition.Enabled), StringComparison.Ordinal))
        {
            return;
        }

        if (sender is not HostSharedFolderDefinition definition)
        {
            return;
        }

        await ApplySharedFolderEnabledToggleAsync(definition);
    }

    private async Task ApplySharedFolderEnabledToggleAsync(HostSharedFolderDefinition definition)
    {
        await _sharedFolderEnabledToggleGate.WaitAsync();
        try
        {
            if (_viewModel.SaveConfigCommand.CanExecute(null))
            {
                await _viewModel.SaveConfigCommand.ExecuteAsync(null);
            }

            if (definition.Enabled)
            {
                await _hostSharedFolderService.EnsureShareAsync(definition, CancellationToken.None);
                _sharedFolderStatusText.Text = $"Eintrag '{definition.ShareName}' im Katalog aktiviert.";
            }
            else
            {
                _sharedFolderStatusText.Text = $"Eintrag '{definition.ShareName}' im Katalog deaktiviert.";
            }

            _sharedFolderLastError = "-";
        }
        catch (Exception ex)
        {
            _sharedFolderLastError = ex.Message;
            _sharedFolderStatusText.Text = $"Aktiv/Deaktiv fehlgeschlagen: {ex.Message}";
            Log.ForContext("eventId", "sharedfolders.host.toggle.failed")
                .Warning(ex,
                "Shared-folder toggle failed. ShareName={ShareName}; Enabled={Enabled}",
                definition.ShareName,
                definition.Enabled);
        }
        finally
        {
            _sharedFolderEnabledToggleGate.Release();
        }
    }

    private async Task SaveSharedFolderEntryAsync()
    {
        var definition = BuildSharedFolderFromEditor();
        if (definition is null)
        {
            return;
        }

        _viewModel.UpsertHostSharedFolderDefinition(definition);
        _sharedFoldersListView.SelectedItem = _viewModel.HostSharedFolders.FirstOrDefault(item => string.Equals(item.Id, definition.Id, StringComparison.OrdinalIgnoreCase));

        if (_viewModel.SaveConfigCommand.CanExecute(null))
        {
            await _viewModel.SaveConfigCommand.ExecuteAsync(null);
        }

        try
        {
            if (definition.Enabled)
            {
                await _hostSharedFolderService.EnsureShareAsync(definition, CancellationToken.None);
                _sharedFolderStatusText.Text = $"Eintrag gespeichert und für HyperTool File Service aktiviert ('{definition.ShareName}').";
            }
            else
            {
                _sharedFolderStatusText.Text = $"Eintrag gespeichert. Share-Kennung '{definition.ShareName}' ist deaktiviert.";
            }

            _sharedFolderLastError = "-";
        }
        catch (Exception ex)
        {
            _sharedFolderLastError = ex.Message;
            _sharedFolderStatusText.Text = $"Speichern fehlgeschlagen: {ex.Message}";
            Log.ForContext("eventId", "sharedfolders.host.save.failed")
                .Warning(ex,
                "Shared-folder save/apply failed. ShareName={ShareName}; LocalPath={LocalPath}; Enabled={Enabled}; ReadOnly={ReadOnly}",
                definition.ShareName,
                definition.LocalPath,
                definition.Enabled,
                definition.ReadOnly);
        }
    }

    private async Task DeleteSharedFolderEntryAsync()
    {
        if (_sharedFoldersListView.SelectedItem is not HostSharedFolderDefinition selected)
        {
            _sharedFolderStatusText.Text = "Kein Shared-Folder Eintrag ausgewählt.";
            return;
        }

        var shareName = (selected.ShareName ?? string.Empty).Trim();

        try
        {
            _viewModel.RemoveHostSharedFolderDefinition(selected.Id);

            if (_viewModel.SaveConfigCommand.CanExecute(null))
            {
                await _viewModel.SaveConfigCommand.ExecuteAsync(null);
            }

            ResetSharedFolderEditor();
            _sharedFolderStatusText.Text = $"Eintrag '{selected.ShareName}' gelöscht.";
            _sharedFolderLastError = "-";
        }
        catch (Exception ex)
        {
            _sharedFolderLastError = ex.Message;
            _sharedFolderStatusText.Text = $"Entfernen fehlgeschlagen: {ex.Message}";
            Log.ForContext("eventId", "sharedfolders.host.delete.failed")
                .Warning(ex,
                "Shared-folder delete failed. ShareName={ShareName}; EntryId={EntryId}",
                selected.ShareName,
                selected.Id);
        }
    }

    private async Task RunHostSharedFolderSelfTestAsync()
    {
        var target = BuildSharedFolderFromEditor();
        if (target is null)
        {
            return;
        }

        try
        {
            var configEntryExists = _viewModel.HostSharedFolders
                .Any(item => string.Equals(item.Id, target.Id, StringComparison.OrdinalIgnoreCase)
                             || string.Equals(item.ShareName, target.ShareName, StringComparison.OrdinalIgnoreCase));
            var pathExists = Directory.Exists(target.LocalPath);
            var hyperVSocketActive = string.Equals(_viewModel.UsbDiagnosticsHyperVSocketText, "Ja", StringComparison.OrdinalIgnoreCase);

            var registryServiceOk = HyperVSocketUsbHostTunnel.IsServiceRegistered(HyperVSocketUsbTunnelDefaults.ServiceId)
                                    && HyperVSocketUsbHostTunnel.IsServiceRegistered(HyperVSocketUsbTunnelDefaults.DiagnosticsServiceId)
                                    && HyperVSocketUsbHostTunnel.IsServiceRegistered(HyperVSocketUsbTunnelDefaults.SharedFolderCatalogServiceId)
                                    && HyperVSocketUsbHostTunnel.IsServiceRegistered(HyperVSocketUsbTunnelDefaults.HostIdentityServiceId)
                                    && HyperVSocketUsbHostTunnel.IsServiceRegistered(HyperVSocketUsbTunnelDefaults.FileServiceId);

            _sharedFolderStatusText.Text = $"Self-Test (Host) · Konfig-Eintrag: {(configEntryExists ? "Ja" : "Nein")} · Pfad vorhanden: {(pathExists ? "Ja" : "Nein")} · Hyper-V Socket: {(hyperVSocketActive ? "Ja" : "Nein")} · Registry-Service: {(registryServiceOk ? "Ja" : "Nein")} · Letzter Fehler: {_sharedFolderLastError}";
        }
        catch (Exception ex)
        {
            _sharedFolderLastError = ex.Message;
            _sharedFolderStatusText.Text = $"Self-Test fehlgeschlagen: {ex.Message}";
            Log.ForContext("eventId", "sharedfolders.host.selftest.failed")
                .Warning(ex, "Shared-folder host self-test failed. ShareName={ShareName}; LocalPath={LocalPath}", target.ShareName, target.LocalPath);
        }
    }

    private Task BrowseSharedFolderPathAsync()
    {
        try
        {
            var folderPath = new UiInteropService().PickFolderPath("Shared-Folder Ordner auswählen");
            if (string.IsNullOrWhiteSpace(folderPath))
            {
                return Task.CompletedTask;
            }

            _sharedFolderPathTextBox.Text = folderPath;
            _sharedFolderStatusText.Text = $"Ordner ausgewählt: {folderPath}";
        }
        catch (Exception ex)
        {
            _sharedFolderStatusText.Text = $"Ordnerauswahl fehlgeschlagen: {ex.Message}";
        }

        return Task.CompletedTask;
    }

    private static ComboBox CreateStyledComboBox()
    {
        return new ComboBox
        {
            CornerRadius = new CornerRadius(8),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            Foreground = Application.Current.Resources["TextPrimaryBrush"] as Brush,
            MinHeight = 34,
            VerticalAlignment = VerticalAlignment.Center
        };
    }

    private static TextBox CreateStyledTextBox(string? value, string? placeholder)
    {
        return new TextBox
        {
            Text = value ?? string.Empty,
            PlaceholderText = placeholder ?? string.Empty,
            CornerRadius = new CornerRadius(8),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            Foreground = Application.Current.Resources["TextPrimaryBrush"] as Brush,
            MinHeight = 34,
            VerticalAlignment = VerticalAlignment.Center
        };
    }

    private static Button CreateIconButton(
        string icon,
        string label,
        ICommand? command = null,
        RoutedEventHandler? onClick = null,
        bool compact = false)
    {
        var iconSize = compact ? 18d : 20d;
        var iconHost = new Grid
        {
            Width = iconSize,
            Height = iconSize,
            VerticalAlignment = VerticalAlignment.Center
        };
        iconHost.Children.Add(new Viewbox
        {
            Stretch = Stretch.Uniform,
            Margin = new Thickness(0),
            Child = new TextBlock
            {
                Text = icon,
                TextAlignment = TextAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = compact ? 16 : 18,
                FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                Foreground = Application.Current.Resources["TextPrimaryBrush"] as Brush
            }
        });

        var content = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            VerticalAlignment = VerticalAlignment.Center
        };
        content.Children.Add(iconHost);
        content.Children.Add(new Border
        {
            MinHeight = iconSize,
            VerticalAlignment = VerticalAlignment.Center,
            Child = new TextBlock
            {
                Text = label,
                VerticalAlignment = VerticalAlignment.Center,
                TextAlignment = TextAlignment.Center,
                FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                TextTrimming = TextTrimming.CharacterEllipsis,
                LineHeight = compact ? 16 : 18,
                LineStackingStrategy = LineStackingStrategy.BlockLineHeight
            }
        });

        var button = new Button
        {
            Content = content,
            Padding = compact ? new Thickness(8, 4, 8, 4) : new Thickness(10, 7, 10, 7),
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush
        };

        if (command is not null)
        {
            button.Command = command;
        }

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
            Color = topBrush?.Color ?? Windows.UI.Color.FromArgb(0xFA, 0xFF, 0xFF, 0xFF),
            Offset = 0.0
        });
        gradientBrush.GradientStops.Add(new GradientStop
        {
            Color = bottomBrush?.Color ?? Windows.UI.Color.FromArgb(0xF0, 0xF2, 0xF8, 0xFF),
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

    private void UpdateNoticeVisibility()
    {
        _configurationNoticeText.Text = _viewModel.ConfigurationNotice ?? string.Empty;
        _configurationNoticeBorder.Visibility = _viewModel.HasConfigurationNotice ? Visibility.Visible : Visibility.Collapsed;
    }

    private void UpdateBusyAndNotificationPanel()
    {
        _toggleLogButton.Content = _viewModel.LogToggleText;
        _busyRing.Visibility = _viewModel.IsBusy ? Visibility.Visible : Visibility.Collapsed;
        _busyRing.IsActive = _viewModel.IsBusy;
        _busyText.Visibility = _viewModel.IsBusy ? Visibility.Visible : Visibility.Collapsed;
        _busyText.Text = _viewModel.BusyText;
        _busyProgress.Visibility = _viewModel.HasBusyProgress ? Visibility.Visible : Visibility.Collapsed;
        _busyProgress.Value = _viewModel.BusyProgressPercent;
        _busyPercentBadge.Visibility = _viewModel.HasBusyProgress ? Visibility.Visible : Visibility.Collapsed;
        _busyPercentText.Text = $"{Math.Clamp(_viewModel.BusyProgressPercent, 0, 100)}%";
        _notificationExpandedGrid.Visibility = _viewModel.IsLogExpanded ? Visibility.Visible : Visibility.Collapsed;
        _notificationSummaryBorder.Visibility = _viewModel.IsLogExpanded ? Visibility.Collapsed : Visibility.Visible;
    }

    private void UpdateNavSelection()
    {
        for (var index = 0; index < _navButtons.Count; index++)
        {
            var isSelected = index == _viewModel.SelectedMenuIndex;
            _navButtons[index].Background = isSelected
                ? Application.Current.Resources["AccentSoftBrush"] as Brush
                : Application.Current.Resources["SurfaceSoftBrush"] as Brush;
            _navButtons[index].BorderBrush = isSelected
                ? Application.Current.Resources["AccentBrush"] as Brush
                : Application.Current.Resources["PanelBorderBrush"] as Brush;
        }
    }

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
        button.Click += (_, _) => _viewModel.SelectedMenuIndex = index;
        _navButtons.Add(button);
        return button;
    }

    private void RequestVmChipsRefresh()
    {
        if (!_isMainLayoutLoaded)
        {
            return;
        }

        _vmChipRefreshDebounceTimer?.Stop();
        _vmChipRefreshDebounceTimer?.Start();
    }

    private string BuildVmChipRefreshSignature()
    {
        var selectedVmName = _viewModel.SelectedVm?.Name ?? string.Empty;
        var snapshot = _viewModel.AvailableVms
            .Select(vm => $"{vm.Name}|{vm.DisplayLabel}|{vm.RuntimeState}|{vm.RuntimeSwitchName}")
            .ToArray();

        return $"{selectedVmName}::{string.Join(";;", snapshot)}";
    }

    private void RefreshVmChips()
    {
        var signature = BuildVmChipRefreshSignature();
        if (string.Equals(signature, _vmChipRefreshSignature, StringComparison.Ordinal))
        {
            UpdateVmChipNavigationButtons();
            return;
        }

        _vmChipRefreshSignature = signature;
        _vmChipPanel.Children.Clear();

        foreach (var vm in _viewModel.AvailableVms)
        {
            _vmChipPanel.Children.Add(CreateVmChip(vm));
        }

        UpdateVmChipNavigationButtons();
    }

    private void ScrollVmChipsBy(double delta)
    {
        var currentOffset = _vmChipScrollViewer.HorizontalOffset;
        var targetOffset = Math.Max(0, currentOffset + delta);
        _vmChipScrollViewer.ChangeView(targetOffset, null, null);
        UpdateVmChipNavigationButtons();
    }

    private void UpdateVmChipNavigationButtons()
    {
        var hasOverflow = _vmChipScrollViewer.ScrollableWidth > 1;

        _vmChipsLeftButton.Visibility = hasOverflow ? Visibility.Visible : Visibility.Collapsed;
        _vmChipsRightButton.Visibility = hasOverflow ? Visibility.Visible : Visibility.Collapsed;

        if (!hasOverflow)
        {
            _vmChipLeftFadeOverlay.Visibility = Visibility.Collapsed;
            _vmChipRightFadeOverlay.Visibility = Visibility.Collapsed;
            return;
        }

        var leftEnabled = _vmChipScrollViewer.HorizontalOffset > 1;
        var rightEnabled = _vmChipScrollViewer.HorizontalOffset < (_vmChipScrollViewer.ScrollableWidth - 1);

        _vmChipsLeftButton.IsEnabled = leftEnabled;
        _vmChipsRightButton.IsEnabled = rightEnabled;

        var fadeBase = IsDarkMode()
            ? Windows.UI.Color.FromArgb(0xFF, 0x13, 0x1C, 0x2F)
            : Windows.UI.Color.FromArgb(0xFF, 0xF3, 0xF6, 0xFB);

        _vmChipLeftFadeOverlay.Background = new LinearGradientBrush
        {
            StartPoint = new Windows.Foundation.Point(0, 0.5),
            EndPoint = new Windows.Foundation.Point(1, 0.5),
            GradientStops =
            {
                new GradientStop { Color = fadeBase, Offset = 0 },
                new GradientStop { Color = Windows.UI.Color.FromArgb(0x00, fadeBase.R, fadeBase.G, fadeBase.B), Offset = 1 }
            }
        };

        _vmChipRightFadeOverlay.Background = new LinearGradientBrush
        {
            StartPoint = new Windows.Foundation.Point(0, 0.5),
            EndPoint = new Windows.Foundation.Point(1, 0.5),
            GradientStops =
            {
                new GradientStop { Color = Windows.UI.Color.FromArgb(0x00, fadeBase.R, fadeBase.G, fadeBase.B), Offset = 0 },
                new GradientStop { Color = fadeBase, Offset = 1 }
            }
        };

        _vmChipLeftFadeOverlay.Visibility = leftEnabled ? Visibility.Visible : Visibility.Collapsed;
        _vmChipRightFadeOverlay.Visibility = rightEnabled ? Visibility.Visible : Visibility.Collapsed;
    }

    private void RefreshVmAdapterCards()
    {
        _vmAdapterCardsPanel.Children.Clear();

        if (_viewModel.AvailableVmNetworkAdapters.Count == 0)
        {
            _vmAdapterCardsPanel.Children.Add(new TextBlock
            {
                Text = "Keine Netzwerkkarten für die ausgewählte VM gefunden.",
                Opacity = 0.85
            });
            return;
        }

        foreach (var adapter in _viewModel.AvailableVmNetworkAdapters)
        {
            var card = new Border
            {
                CornerRadius = new CornerRadius(10),
                BorderThickness = new Thickness(1),
                BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
                Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
                Padding = new Thickness(10)
            };

            var cardStack = new StackPanel { Spacing = 8 };

            var topRow = new Grid { ColumnSpacing = 8 };
            topRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            topRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            topRow.Children.Add(new TextBlock
            {
                Text = string.IsNullOrWhiteSpace(adapter.DisplayName) ? adapter.Name : adapter.DisplayName,
                FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center
            });

            var disconnectButton = CreateIconButton("⛔", "Disconnect", compact: true, onClick: async (_, _) =>
            {
                await _viewModel.DisconnectAdapterByNameCommand.ExecuteAsync(adapter.Name);
                RefreshVmAdapterCards();
            });
            Grid.SetColumn(disconnectButton, 1);
            topRow.Children.Add(disconnectButton);
            cardStack.Children.Add(topRow);

            cardStack.Children.Add(new TextBlock
            {
                Text = $"Aktiv: {adapter.SwitchName}",
                Opacity = 0.8
            });

            var switchButtons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 8
            };

            foreach (var vmSwitch in _viewModel.AvailableSwitches)
            {
                var isActive = string.Equals(adapter.SwitchName, vmSwitch.Name, StringComparison.OrdinalIgnoreCase);
                var switchButton = CreateIconButton("↔", isActive ? $"{vmSwitch.Name} (Aktiv)" : vmSwitch.Name, compact: true, onClick: async (_, _) =>
                {
                    await _viewModel.ConnectAdapterToSwitchByKeyCommand.ExecuteAsync($"{adapter.Name}|||{vmSwitch.Name}");
                    RefreshVmAdapterCards();
                });

                if (isActive)
                {
                    switchButton.Background = Application.Current.Resources["AccentSoftBrush"] as Brush;
                    switchButton.BorderBrush = Application.Current.Resources["AccentBrush"] as Brush;
                }

                switchButtons.Children.Add(switchButton);
            }

            cardStack.Children.Add(new ScrollViewer
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Content = switchButtons
            });
            card.Child = cardStack;
            _vmAdapterCardsPanel.Children.Add(card);
        }
    }

    private bool IsDarkMode()
        => string.Equals(_viewModel.UiTheme, "Dark", StringComparison.OrdinalIgnoreCase);

    private static bool IsVmRunningState(string? runtimeState)
        => !string.IsNullOrWhiteSpace(runtimeState)
           && (runtimeState.Contains("running", StringComparison.OrdinalIgnoreCase)
               || runtimeState.Contains("läuft", StringComparison.OrdinalIgnoreCase)
               || runtimeState.Contains("ausgeführt", StringComparison.OrdinalIgnoreCase));

    private static bool IsVmOffState(string? runtimeState)
        => !string.IsNullOrWhiteSpace(runtimeState)
           && (runtimeState.Contains("off", StringComparison.OrdinalIgnoreCase)
               || runtimeState.Contains("aus", StringComparison.OrdinalIgnoreCase));

    private void UpdateSelectedVmStateTextStyle()
    {
        static SolidColorBrush Brush(byte a, byte r, byte g, byte b) => new(Color.FromArgb(a, r, g, b));

        var state = _viewModel.SelectedVmState;
        var isDark = IsDarkMode();

        if (IsVmRunningState(state))
        {
            _selectedVmStateText.Foreground = isDark
                ? Brush(0xFF, 0xD9, 0xF6, 0xE8)
                : Brush(0xFF, 0x0E, 0x4F, 0x31);
            return;
        }

        if (IsVmOffState(state))
        {
            _selectedVmStateText.Foreground = isDark
                ? Brush(0xFF, 0xFF, 0xDC, 0xE5)
                : Brush(0xFF, 0x77, 0x1D, 0x33);
            return;
        }

        _selectedVmStateText.Foreground = Application.Current.Resources["TextMutedBrush"] as Brush ?? Brush(0xFF, 0xA6, 0xB9, 0xD8);
    }

    private Button CreateVmChip(VmDefinition vm)
    {
        var chip = new Button
        {
            Padding = new Thickness(11, 8, 12, 8),
            MinWidth = 0,
            HorizontalAlignment = HorizontalAlignment.Left,
            CornerRadius = new CornerRadius(12),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush
        };

        var isDark = IsDarkMode();

        if (IsVmRunningState(vm.RuntimeState))
        {
            chip.Background = isDark
                ? new SolidColorBrush(Windows.UI.Color.FromArgb(0x38, 0x2E, 0x64, 0x4A))
                : new SolidColorBrush(Windows.UI.Color.FromArgb(0x30, 0x73, 0xB7, 0x8E));
            chip.BorderBrush = isDark
                ? new SolidColorBrush(Windows.UI.Color.FromArgb(0xBA, 0x4D, 0xA6, 0x79))
                : new SolidColorBrush(Windows.UI.Color.FromArgb(0xA6, 0x5E, 0xAF, 0x87));
        }
        else if (IsVmOffState(vm.RuntimeState))
        {
            chip.Background = isDark
                ? new SolidColorBrush(Windows.UI.Color.FromArgb(0x34, 0x67, 0x35, 0x3F))
                : new SolidColorBrush(Windows.UI.Color.FromArgb(0x2C, 0xD4, 0x7E, 0x89));
            chip.BorderBrush = isDark
                ? new SolidColorBrush(Windows.UI.Color.FromArgb(0xB8, 0xB8, 0x5D, 0x6B))
                : new SolidColorBrush(Windows.UI.Color.FromArgb(0xA8, 0xC2, 0x69, 0x79));
        }

        if (_viewModel.SelectedVm is not null
            && string.Equals(_viewModel.SelectedVm.Name, vm.Name, StringComparison.OrdinalIgnoreCase))
        {
            chip.BorderThickness = new Thickness(2);
            chip.BorderBrush = Application.Current.Resources["AccentBrush"] as Brush;
        }

        var iconBadge = new Grid
        {
            Width = 31,
            Height = 31,
            VerticalAlignment = VerticalAlignment.Center
        };

        iconBadge.Children.Add(new Border
        {
            Width = 31,
            Height = 31,
            CornerRadius = new CornerRadius(9),
            VerticalAlignment = VerticalAlignment.Center,
            BorderThickness = new Thickness(1),
            BorderBrush = chip.BorderBrush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            Child = new TextBlock
            {
                Text = "🖥",
                FontSize = 17,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                TextAlignment = TextAlignment.Center
            }
        });

        var isDefaultVm = string.Equals(vm.Name, _viewModel.DefaultVmName, StringComparison.OrdinalIgnoreCase);

        if (isDefaultVm)
        {
            var defaultBadge = new Border
            {
                Width = 13,
                Height = 13,
                CornerRadius = new CornerRadius(6.5),
                Background = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, 0xF0, 0xC7, 0x56)),
                BorderThickness = new Thickness(1),
                BorderBrush = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, 0xB8, 0x82, 0x22)),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Bottom,
                Margin = new Thickness(0, 0, -3, -3),
                Child = new TextBlock
                {
                    Text = "★",
                    FontSize = 8.5,
                    FontWeight = Microsoft.UI.Text.FontWeights.Bold,
                    Foreground = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, 0x6F, 0x4A, 0x00)),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    TextAlignment = TextAlignment.Center
                }
            };

            iconBadge.Children.Add(defaultBadge);
        }

        var text = new TextBlock
        {
            Text = vm.DisplayLabel,
            TextTrimming = TextTrimming.CharacterEllipsis,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            MaxWidth = 350
        };

        var subText = new TextBlock
        {
            Text = $"{vm.RuntimeState} · {vm.RuntimeSwitchName}",
            TextTrimming = TextTrimming.CharacterEllipsis,
            Opacity = 0.82,
            FontSize = 12,
            MaxWidth = 350
        };

        var textStack = new StackPanel { Spacing = 1 };
        textStack.Children.Add(text);
        textStack.Children.Add(subText);

        var chipContent = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 10,
            VerticalAlignment = VerticalAlignment.Center
        };
        chipContent.Children.Add(iconBadge);
        chipContent.Children.Add(textStack);

        chip.Content = chipContent;

        chip.Click += (_, _) => _viewModel.SelectVmFromChipCommand.Execute(vm);

        var flyout = new MenuFlyout();
        flyout.Items.Add(CreateVmMenuItem("Start", () => _viewModel.StartVmByNameCommand.ExecuteAsync(vm.Name)));
        flyout.Items.Add(CreateVmMenuItem("Stop", () => _viewModel.StopVmByNameCommand.ExecuteAsync(vm.Name)));
        flyout.Items.Add(CreateVmMenuItem("Hard Off", () => _viewModel.TurnOffVmByNameCommand.ExecuteAsync(vm.Name)));
        flyout.Items.Add(CreateVmMenuItem("Restart", () => _viewModel.RestartVmByNameCommand.ExecuteAsync(vm.Name)));
        flyout.Items.Add(new MenuFlyoutSeparator());
        flyout.Items.Add(CreateVmMenuItem("Als Default-VM setzen", () => SetDefaultVmFromChipAsync(vm)));
        flyout.Items.Add(CreateVmMenuItem("Schnellstart-Verknüpfung erstellen", () => CreateVmQuickstartForVmAsync(vm)));
        flyout.Items.Add(new MenuFlyoutSeparator());
        flyout.Items.Add(CreateVmMenuItem("Open Console", () => _viewModel.OpenConsoleByNameCommand.ExecuteAsync(vm.Name)));
        flyout.Items.Add(CreateVmMenuItem("Snapshot", () => _viewModel.CreateSnapshotByNameCommand.ExecuteAsync(vm.Name)));
        chip.ContextFlyout = flyout;

        return chip;
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
            var soundPath = System.IO.Path.Combine(AppContext.BaseDirectory, "Assets", "logo-spin.wav");
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

    private void UpdatePageContent()
    {
        try
        {
            var content = _viewModel.SelectedMenuIndex switch
            {
                1 => _usbPage ??= BuildUsbPage(),
                2 => _sharedFoldersPage ??= BuildSharedFoldersPage(),
                3 => _snapshotsPage ??= BuildSnapshotsPage(),
                4 => _configPage ??= BuildConfigPage(),
                5 => _infoPage ??= BuildInfoPage(),
                _ => _vmPage ??= BuildVmPage()
            };

            _pageContent.Content = content;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to switch to page index {PageIndex}", _viewModel.SelectedMenuIndex);

            var fallback = new StackPanel { Spacing = 8 };
            fallback.Children.Add(new TextBlock
            {
                Text = "Diese Ansicht konnte nicht geladen werden.",
                FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
            });
            fallback.Children.Add(new TextBlock
            {
                Text = "Bitte auf VM oder Info wechseln und Log prüfen."
            });

            _pageContent.Content = fallback;
        }
    }

    private void OnVmItemClick(object sender, ItemClickEventArgs e)
    {
        if (e.ClickedItem is VmDefinition vm)
        {
            _viewModel.SelectedVm = vm;
        }
    }

    private void OnVmListRightTapped(object sender, RightTappedRoutedEventArgs e)
    {
        if (e.OriginalSource is not FrameworkElement element)
        {
            return;
        }

        var vm = element.DataContext as VmDefinition;
        if (vm is null)
        {
            return;
        }

        var flyout = new MenuFlyout();
        flyout.Items.Add(CreateVmMenuItem("Start", () => _viewModel.StartVmByNameCommand.ExecuteAsync(vm.Name)));
        flyout.Items.Add(CreateVmMenuItem("Stop", () => _viewModel.StopVmByNameCommand.ExecuteAsync(vm.Name)));
        flyout.Items.Add(CreateVmMenuItem("Hard Off", () => _viewModel.TurnOffVmByNameCommand.ExecuteAsync(vm.Name)));
        flyout.Items.Add(CreateVmMenuItem("Restart", () => _viewModel.RestartVmByNameCommand.ExecuteAsync(vm.Name)));
        flyout.Items.Add(new MenuFlyoutSeparator());
        flyout.Items.Add(CreateVmMenuItem("Als Default-VM setzen", () => SetDefaultVmFromChipAsync(vm)));
        flyout.Items.Add(new MenuFlyoutSeparator());
        flyout.Items.Add(CreateVmMenuItem("Schnellstart-Verknüpfung erstellen", () => CreateVmQuickstartForVmAsync(vm)));
        flyout.Items.Add(new MenuFlyoutSeparator());
        flyout.Items.Add(CreateVmMenuItem("Open Console", () => _viewModel.OpenConsoleByNameCommand.ExecuteAsync(vm.Name)));
        flyout.Items.Add(CreateVmMenuItem("Snapshot", () => _viewModel.CreateSnapshotByNameCommand.ExecuteAsync(vm.Name)));
        flyout.ShowAt(element);
        e.Handled = true;
    }

    private Task SetDefaultVmFromChipAsync(VmDefinition vm)
    {
        if (vm is null)
        {
            return Task.CompletedTask;
        }

        _viewModel.SelectedVm = vm;
        _viewModel.SelectedVmForConfig = vm;
        _viewModel.SetDefaultVmCommand.Execute(null);
        RequestVmChipsRefresh();
        return Task.CompletedTask;
    }

    private async Task CreateVmQuickstartForVmAsync(VmDefinition vm)
    {
        if (vm is null)
        {
            return;
        }

        _viewModel.SelectedVm = vm;
        _viewModel.SelectedVmForConfig = vm;
        await CreateVmQuickstartShortcutAsync();
    }

    private static MenuFlyoutItem CreateVmMenuItem(string text, Func<Task> action)
    {
        var item = new MenuFlyoutItem { Text = text };
        item.Click += async (_, _) => await action();
        return item;
    }

    private void OpenHelpWindow()
    {
        if (_helpWindow is not null)
        {
            _helpWindow.Activate();
            return;
        }

        var repoUrl = $"https://github.com/{_viewModel.GithubOwner}/{_viewModel.GithubRepo}";
        _helpWindow = new HelpWindow(_viewModel.ConfigPath, repoUrl, _viewModel.UiTheme);
        _helpWindow.Closed += (_, _) => _helpWindow = null;
        _helpWindow.Activate();
    }

    private async Task OpenHostNetworkWindowAsync()
    {
        if (_hostNetworkWindow is not null)
        {
            _hostNetworkWindow.Activate();
            return;
        }

        var adapters = await _viewModel.GetHostNetworkAdaptersWithUplinkAsync();
        if (adapters.Count == 0)
        {
            return;
        }

        _hostNetworkWindow = new HostNetworkWindow(adapters, _viewModel.UiTheme);
        _hostNetworkWindow.Closed += (_, _) => _hostNetworkWindow = null;
        _hostNetworkWindow.Activate();
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

        try
        {
            _hostNetworkWindow?.Close();
        }
        catch
        {
        }
        finally
        {
            _hostNetworkWindow = null;
        }
    }

    private void OnWindowClosed(object sender, WindowEventArgs args)
    {
        CloseAuxiliaryWindows();

        _vmChipRefreshDebounceTimer?.Stop();
        _vmChipRefreshDebounceTimer = null;

        _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
        _viewModel.AvailableVms.CollectionChanged -= OnAvailableVmsCollectionChanged;
        _viewModel.AvailableVmNetworkAdapters.CollectionChanged -= OnVmNetworkAdaptersCollectionChanged;
        _viewModel.AvailableSwitches.CollectionChanged -= OnVmNetworkAdaptersCollectionChanged;
        _viewModel.AvailableCheckpointTree.CollectionChanged -= OnCheckpointTreeCollectionChanged;
    }

    private void OnAvailableVmsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        RunOnUiThread(() =>
        {
            RequestVmChipsRefresh();
        });
    }

    private void OnVmNetworkAdaptersCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        RunOnUiThread(() =>
        {
            RefreshVmAdapterCards();
        });
    }

    private void OnCheckpointTreeCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        RunOnUiThread(() =>
        {
            RefreshCheckpointTreeView();
        });
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        RunOnUiThread(() => HandleViewModelPropertyChangedOnUiThread(e));
    }

    private void HandleViewModelPropertyChangedOnUiThread(PropertyChangedEventArgs e)
    {
        if (string.Equals(e.PropertyName, nameof(MainViewModel.UiTheme), StringComparison.Ordinal))
        {
            if (_isMainLayoutLoaded)
            {
                _ = RestartForThemeChangeAsync();
            }
            else
            {
                _themeService.ApplyTheme(_viewModel.UiTheme);
                ApplyRequestedTheme();
                UpdateTitleBarAppearance();
            }
        }

        if (string.Equals(e.PropertyName, nameof(MainViewModel.SelectedMenuIndex), StringComparison.Ordinal))
        {
            UpdatePageContent();
            UpdateNavSelection();
        }

        if (string.Equals(e.PropertyName, nameof(MainViewModel.SelectedVm), StringComparison.Ordinal)
            || string.Equals(e.PropertyName, nameof(MainViewModel.SelectedVmState), StringComparison.Ordinal)
            || string.Equals(e.PropertyName, nameof(MainViewModel.SelectedVmDisplayName), StringComparison.Ordinal))
        {
            UpdateSelectedVmStateTextStyle();

            if (_isMainLayoutLoaded)
            {
                RequestVmChipsRefresh();
                RefreshVmAdapterCards();
            }
        }

        if (string.Equals(e.PropertyName, nameof(MainViewModel.SelectedVmAdapterSwitchDisplay), StringComparison.Ordinal)
            || string.Equals(e.PropertyName, nameof(MainViewModel.UiTheme), StringComparison.Ordinal))
        {
            UpdateSelectedVmStateTextStyle();

            if (_isMainLayoutLoaded)
            {
                RefreshVmAdapterCards();
            }
        }

        if (string.Equals(e.PropertyName, nameof(MainViewModel.ConfigurationNotice), StringComparison.Ordinal)
            || string.Equals(e.PropertyName, nameof(MainViewModel.HasConfigurationNotice), StringComparison.Ordinal))
        {
            UpdateNoticeVisibility();
        }

        if (string.Equals(e.PropertyName, nameof(MainViewModel.IsBusy), StringComparison.Ordinal)
            || string.Equals(e.PropertyName, nameof(MainViewModel.BusyText), StringComparison.Ordinal)
            || string.Equals(e.PropertyName, nameof(MainViewModel.BusyProgressPercent), StringComparison.Ordinal)
            || string.Equals(e.PropertyName, nameof(MainViewModel.HasBusyProgress), StringComparison.Ordinal)
            || string.Equals(e.PropertyName, nameof(MainViewModel.IsLogExpanded), StringComparison.Ordinal)
            || string.Equals(e.PropertyName, nameof(MainViewModel.LogToggleText), StringComparison.Ordinal)
            || string.Equals(e.PropertyName, nameof(MainViewModel.LastNotificationText), StringComparison.Ordinal))
        {
            UpdateBusyAndNotificationPanel();
        }

        if (string.Equals(e.PropertyName, nameof(MainViewModel.SelectedUsbDevice), StringComparison.Ordinal)
            && _usbDevicesListView.SelectedItem != _viewModel.SelectedUsbDevice)
        {
            _usbDevicesListView.SelectedItem = _viewModel.SelectedUsbDevice;
            _usbAutoShareCheckBox.IsEnabled = _viewModel.SelectedUsbDevice is not null && _viewModel.UsbRuntimeAvailable && _viewModel.HostUsbSharingEnabled;
        }

        if (string.Equals(e.PropertyName, nameof(MainViewModel.UsbRuntimeAvailable), StringComparison.Ordinal))
        {
            _usbAutoShareCheckBox.IsEnabled = _viewModel.SelectedUsbDevice is not null && _viewModel.UsbRuntimeAvailable && _viewModel.HostUsbSharingEnabled;
            UpdateHostFeatureAvailabilityUi();
        }

        if (string.Equals(e.PropertyName, nameof(MainViewModel.HostUsbSharingEnabled), StringComparison.Ordinal)
            || string.Equals(e.PropertyName, nameof(MainViewModel.HostSharedFoldersEnabled), StringComparison.Ordinal))
        {
            UpdateHostFeatureAvailabilityUi();
        }

        if (string.Equals(e.PropertyName, nameof(MainViewModel.UsbAutoDetachOnClientDisconnect), StringComparison.Ordinal))
        {
            _usbAutoDetachOnDisconnectCheckBox.IsChecked = _viewModel.UsbAutoDetachOnClientDisconnect;
        }

        if (string.Equals(e.PropertyName, nameof(MainViewModel.UsbUnshareOnExit), StringComparison.Ordinal))
        {
            _usbUnshareOnExitCheckBox.IsChecked = _viewModel.UsbUnshareOnExit;
        }

        if (string.Equals(e.PropertyName, nameof(MainViewModel.SelectedUsbDeviceAutoShareEnabled), StringComparison.Ordinal))
        {
            _usbAutoShareCheckBox.IsChecked = _viewModel.SelectedUsbDeviceAutoShareEnabled;
        }

        if (string.Equals(e.PropertyName, nameof(MainViewModel.SelectedCheckpointNode), StringComparison.Ordinal))
        {
            SelectCheckpointNodeInTree(_viewModel.SelectedCheckpointNode);
        }
    }

    private void RunOnUiThread(Action action)
    {
        if (DispatcherQueue.HasThreadAccess)
        {
            action();
            return;
        }

        _ = DispatcherQueue.TryEnqueue(() => action());
    }

    private void UpdateUsbRuntimeStatusUi()
    {
        var isFeatureEnabled = _viewModel.HostUsbSharingEnabled;
        var isAvailable = _viewModel.UsbRuntimeAvailable;

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

        _usbRuntimeHintText.Visibility = (isFeatureEnabled && !isAvailable) ? Visibility.Visible : Visibility.Collapsed;

        _usbRemoteFxPolicyHintText.Visibility = Visibility.Visible;
    }

    private void UpdateHostFeatureAvailabilityUi()
    {
        var usbRuntimeAvailable = _viewModel.UsbRuntimeAvailable;
        var usbEnabled = _viewModel.HostUsbSharingEnabled && usbRuntimeAvailable;
        var sharedFoldersEnabled = _viewModel.HostSharedFoldersEnabled;

        _suppressHostFeatureToggleEvents = true;
        try
        {
            _usbHostFeatureEnabledToggleSwitch.IsOn = usbEnabled;
            _sharedFolderHostFeatureEnabledToggleSwitch.IsOn = sharedFoldersEnabled;
        }
        finally
        {
            _suppressHostFeatureToggleEvents = false;
        }

        var usbInteractive = usbEnabled;
        if (_usbFeatureControlsPanel is not null)
        {
            _usbFeatureControlsPanel.IsHitTestVisible = true;
            _usbFeatureControlsPanel.Opacity = 1.0;
        }

        if (_usbRefreshButton is not null)
        {
            _usbRefreshButton.IsEnabled = usbInteractive;
        }

        if (_usbShareButton is not null)
        {
            _usbShareButton.IsEnabled = usbInteractive;
        }

        if (_usbUnshareButton is not null)
        {
            _usbUnshareButton.IsEnabled = usbInteractive;
        }

        _usbAutoDetachOnDisconnectCheckBox.IsEnabled = usbInteractive;
        _usbUnshareOnExitCheckBox.IsEnabled = usbInteractive;

        _usbHostFeatureEnabledToggleSwitch.IsEnabled = usbRuntimeAvailable;

        _usbDisabledOverlayText.Visibility = usbInteractive ? Visibility.Collapsed : Visibility.Visible;

        var usbChipPalette = ResolveFeatureChipPalette(usbEnabled);
        _usbHostFeatureStatusChip.Background = usbChipPalette.chipBackground;
        _usbHostFeatureStatusChip.BorderBrush = usbChipPalette.chipBorder;
        _usbHostFeatureStatusChipText.Text = usbEnabled ? "Aktiv" : "Inaktiv";
        _usbHostFeatureStatusChipText.Foreground = usbChipPalette.textForeground;

        if (!usbEnabled)
        {
            _viewModel.UsbStatusText = "USB Share ist global deaktiviert. Aktivieren, um Geräte wieder zu laden/freizugeben.";
        }

        var sharedFoldersInteractive = sharedFoldersEnabled;

        var sharedFolderChipPalette = ResolveFeatureChipPalette(sharedFoldersEnabled);
        _sharedFolderHostFeatureStatusChip.Background = sharedFolderChipPalette.chipBackground;
        _sharedFolderHostFeatureStatusChip.BorderBrush = sharedFolderChipPalette.chipBorder;
        _sharedFolderHostFeatureStatusChipText.Text = sharedFoldersEnabled ? "Aktiv" : "Inaktiv";
        _sharedFolderHostFeatureStatusChipText.Foreground = sharedFolderChipPalette.textForeground;

        _sharedFolderDisabledOverlayText.Visibility = sharedFoldersInteractive ? Visibility.Collapsed : Visibility.Visible;
        _sharedFoldersListView.Visibility = sharedFoldersInteractive ? Visibility.Visible : Visibility.Collapsed;

        _sharedFolderPathTextBox.IsEnabled = sharedFoldersInteractive;
        _sharedFolderShareNameTextBox.IsEnabled = sharedFoldersInteractive;
        _sharedFolderReadOnlyCheckBox.IsEnabled = sharedFoldersInteractive;

        if (_sharedFolderNewButton is not null)
        {
            _sharedFolderNewButton.IsEnabled = sharedFoldersInteractive;
        }

        if (_sharedFolderSaveButton is not null)
        {
            _sharedFolderSaveButton.IsEnabled = sharedFoldersInteractive;
        }

        if (_sharedFolderDeleteButton is not null)
        {
            _sharedFolderDeleteButton.IsEnabled = sharedFoldersInteractive;
        }

        if (_sharedFoldersListView is not null)
        {
            _sharedFoldersListView.IsEnabled = sharedFoldersInteractive;
        }

        UpdateUsbRuntimeStatusUi();
    }

    private (Brush chipBackground, Brush chipBorder, Brush textForeground) ResolveFeatureChipPalette(bool isActive)
    {
        static SolidColorBrush Brush(byte a, byte r, byte g, byte b) => new(Color.FromArgb(a, r, g, b));
        var isDarkMode = IsDarkMode();

        if (isActive)
        {
            if (isDarkMode)
            {
                return (
                    Brush(0xFF, 0x14, 0x3C, 0x2C),
                    Brush(0xFF, 0x43, 0xB5, 0x81),
                    Brush(0xFF, 0xD9, 0xF6, 0xE8));
            }

            return (
                Brush(0xFF, 0xE8, 0xF8, 0xEF),
                Brush(0xFF, 0x2F, 0x9E, 0x68),
                Brush(0xFF, 0x0E, 0x4F, 0x31));
        }

        return (
            Application.Current.Resources["SurfaceSoftBrush"] as Brush ?? Brush(0xFF, 0x20, 0x2A, 0x48),
            Application.Current.Resources["PanelBorderBrush"] as Brush ?? Brush(0xFF, 0x44, 0x57, 0x7F),
            Application.Current.Resources["TextMutedBrush"] as Brush ?? Brush(0xFF, 0xA6, 0xB9, 0xD8));
    }

    private async Task OnHostUsbFeatureToggleChangedAsync(bool enabled)
    {
        if (_suppressHostFeatureToggleEvents)
        {
            return;
        }

        if (enabled && !_viewModel.UsbRuntimeAvailable)
        {
            _viewModel.PublishNotification("USB-Share kann erst aktiviert werden, wenn usbipd-win installiert ist.", "Warning");
            UpdateHostFeatureAvailabilityUi();
            return;
        }

        try
        {
            await _viewModel.SetHostUsbSharingEnabledAsync(enabled);

            if (_viewModel.SaveConfigCommand.CanExecute(null))
            {
                await _viewModel.SaveConfigCommand.ExecuteAsync(null);
            }

            await NotifyGuestHostFeatureChangedAsync();

            if (enabled)
            {
                if (_viewModel.RefreshUsbDevicesCommand.CanExecute(null))
                {
                    await _viewModel.RefreshUsbDevicesCommand.ExecuteAsync(null);
                }
            }
        }
        catch (Exception ex)
        {
            _viewModel.PublishNotification($"USB-Share konnte nicht umgeschaltet werden: {ex.Message}", "Error");
        }

        UpdateHostFeatureAvailabilityUi();
    }

    private async Task OnHostSharedFolderFeatureToggleChangedAsync(bool enabled)
    {
        if (_suppressHostFeatureToggleEvents)
        {
            return;
        }

        await _viewModel.SetHostSharedFoldersEnabledAsync(enabled);

        if (_viewModel.SaveConfigCommand.CanExecute(null))
        {
            await _viewModel.SaveConfigCommand.ExecuteAsync(null);
        }

        await NotifyGuestHostFeatureChangedAsync();

        UpdateHostFeatureAvailabilityUi();
    }

    private static async Task NotifyGuestHostFeatureChangedAsync()
    {
        try
        {
            using var client = new NamedPipeClientStream(".", GuestSingleInstancePipeName, PipeDirection.Out, PipeOptions.Asynchronous);
            await client.ConnectAsync(250);
            await using var writer = new StreamWriter(client, Encoding.UTF8, leaveOpen: false);
            await writer.WriteLineAsync("REFRESH_HOST_FEATURES");
            await writer.FlushAsync();
        }
        catch
        {
        }
    }

    private async Task InstallHostUsbRuntimeAsync()
    {
        if (_usbRuntimeInstallButton is not null)
        {
            _usbRuntimeInstallButton.IsEnabled = false;
        }

        try
        {
            _viewModel.PublishNotification("usbipd-win Installer wird vorbereitet...", "Info");

            var installerResult = await new GitHubUpdateService().CheckForUpdateAsync(
                HostUsbRuntimeOwner,
                HostUsbRuntimeRepo,
                "0.0.0",
                CancellationToken.None,
                HostUsbRuntimeAssetHint);

            if (!installerResult.Success || string.IsNullOrWhiteSpace(installerResult.InstallerDownloadUrl))
            {
                _viewModel.PublishNotification("Installer-Asset konnte nicht automatisch ermittelt werden. Release-Seite wird geöffnet.", "Warning");
                Process.Start(new ProcessStartInfo
                {
                    FileName = "https://github.com/dorssel/usbipd-win/releases/latest",
                    UseShellExecute = true
                });
                return;
            }

            var targetDirectory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "HyperTool", "runtime-installers");
            Directory.CreateDirectory(targetDirectory);

            var fileName = ResolveInstallerFileName(
                installerResult.InstallerDownloadUrl,
                installerResult.InstallerFileName,
                "usbipd-win-x64.msi");

            var installerPath = System.IO.Path.Combine(targetDirectory, fileName);

            _viewModel.PublishNotification($"Lade usbipd-win herunter: {fileName}", "Info");
            using (var response = await RuntimeInstallerDownloadClient.GetAsync(installerResult.InstallerDownloadUrl, CancellationToken.None))
            {
                response.EnsureSuccessStatusCode();
                await using var stream = new System.IO.FileStream(installerPath, FileMode.Create, FileAccess.Write, FileShare.None);
                await response.Content.CopyToAsync(stream);
            }

            var extension = System.IO.Path.GetExtension(installerPath);
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

            _viewModel.PublishNotification("usbipd-win Installer gestartet. Nach Abschluss USB-Refresh oder App-Neustart ausführen.", "Success");
            _ = RefreshHostUsbRuntimeAfterInstallAsync();
        }
        catch (Exception ex)
        {
            _viewModel.PublishNotification($"Automatische usbipd-win Installation fehlgeschlagen: {ex.Message}", "Error");
        }
        finally
        {
            if (_usbRuntimeInstallButton is not null)
            {
                _usbRuntimeInstallButton.IsEnabled = true;
            }
        }
    }

    private async Task RefreshHostUsbRuntimeAfterInstallAsync()
    {
        var attempts = 30;
        for (var attempt = 0; attempt < attempts; attempt++)
        {
            try
            {
                await Task.Delay(TimeSpan.FromSeconds(4));
                await _viewModel.RefreshUsbRuntimeAvailabilityAsync();
                UpdateHostFeatureAvailabilityUi();

                if (_viewModel.UsbRuntimeAvailable)
                {
                    _viewModel.PublishNotification("usbipd-win erkannt. Installationsbutton wird ausgeblendet.", "Success");
                    break;
                }
            }
            catch
            {
            }
        }
    }

    private async Task RestartHostToolAsync()
    {
        var overlayShown = false;

        try
        {
            if (_isThemeRestartInProgress)
            {
                return;
            }

            await ShowThemeTransitionOverlayAsync("Layout wird aktualisiert …");
            overlayShown = true;

            if (_themeTransitionStatusText is not null)
            {
                _themeTransitionStatusText.Text = "Tool wird neu gestartet …";
            }

            await Task.Delay(900);

            if (Application.Current is not App app)
            {
                _viewModel.PublishNotification("Neustart nicht möglich: App-Kontext nicht verfügbar.", "Error");
                return;
            }

            _viewModel.PublishNotification("Tool wird neu gestartet …", "Info");
            await app.ReopenMainWindowForThemeChangeAsync();
        }
        catch (Exception ex)
        {
            _viewModel.PublishNotification($"Neustart fehlgeschlagen: {ex.Message}", "Error");
        }
        finally
        {
            if (overlayShown)
            {
                await HideThemeTransitionOverlaySafeAsync();
            }
        }
    }

    private async Task CreateVmQuickstartShortcutAsync()
    {
        var vm = _viewModel.SelectedVmForConfig ?? _viewModel.SelectedVm;
        if (vm is null || string.IsNullOrWhiteSpace(vm.Name))
        {
            _viewModel.PublishNotification("Bitte zuerst eine VM auswählen.", "Info");
            return;
        }

        var vmName = vm.Name.Trim();
        var displayLabel = string.IsNullOrWhiteSpace(vm.DisplayLabel) ? vmName : vm.DisplayLabel.Trim();
        var safeFileName = string.Join("_", displayLabel.Split(System.IO.Path.GetInvalidFileNameChars(), StringSplitOptions.RemoveEmptyEntries)).Trim();
        if (string.IsNullOrWhiteSpace(safeFileName))
        {
            safeFileName = "HyperTool-VM";
        }

        var desktopDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        var shortcutPath = System.IO.Path.Combine(desktopDirectory, $"HyperTool Start {safeFileName}.lnk");

        var ps =
            "$ErrorActionPreference='Stop'; " +
            "$shell=New-Object -ComObject WScript.Shell; " +
            $"$lnk=$shell.CreateShortcut('{shortcutPath.Replace("'", "''")}'); " +
            "$lnk.TargetPath='powershell.exe'; " +
            $"$lnk.Arguments='-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command \"Start-VM -Name ''' + '{vmName.Replace("'", "''")}' + ''' -Confirm:$false\"'; " +
            "$lnk.WorkingDirectory=$env:SystemRoot; " +
            "$lnk.IconLocation=$env:SystemRoot + '\\\\System32\\\\shell32.dll,25'; " +
            $"$lnk.Description='HyperTool Schnellstart fuer VM {displayLabel.Replace("'", "''")}'; " +
            "$lnk.Save();";

        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            UseShellExecute = false,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        psi.ArgumentList.Add("-NoProfile");
        psi.ArgumentList.Add("-NonInteractive");
        psi.ArgumentList.Add("-ExecutionPolicy");
        psi.ArgumentList.Add("Bypass");
        psi.ArgumentList.Add("-Command");
        psi.ArgumentList.Add(ps);

        try
        {
            using var process = Process.Start(psi);
            if (process is null)
            {
                throw new InvalidOperationException("PowerShell konnte nicht gestartet werden.");
            }

            var stdErr = await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            if (process.ExitCode != 0)
            {
                throw new InvalidOperationException(string.IsNullOrWhiteSpace(stdErr)
                    ? $"ExitCode {process.ExitCode}"
                    : stdErr.Trim());
            }

            _viewModel.PublishNotification($"Schnellstart-Verknuepfung erstellt: {shortcutPath}", "Success");
        }
        catch (Exception ex)
        {
            _viewModel.PublishNotification($"Schnellstart-Verknuepfung konnte nicht erstellt werden: {ex.Message}", "Error");
        }
    }

    private static string ResolveInstallerFileName(string downloadUrl, string? fileName, string defaultFileName)
    {
        if (!string.IsNullOrWhiteSpace(fileName))
        {
            return fileName.Trim();
        }

        if (Uri.TryCreate(downloadUrl, UriKind.Absolute, out var uri))
        {
            var inferred = System.IO.Path.GetFileName(uri.LocalPath);
            if (!string.IsNullOrWhiteSpace(inferred))
            {
                return inferred;
            }
        }

        return defaultFileName;
    }

    private void UpdateTitleBarAppearance()
    {
        try
        {
            if (AppWindow?.TitleBar is not AppWindowTitleBar titleBar)
            {
                return;
            }

            if (IsDarkMode())
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

    private void ApplyRequestedTheme()
    {
        if (Content is FrameworkElement root)
        {
            root.RequestedTheme = IsDarkMode()
                ? ElementTheme.Dark
                : ElementTheme.Light;
        }
    }

    private void RebuildMainLayoutForTheme()
    {
        _vmPage = null;
        _snapshotsPage = null;
        _sharedFoldersPage = null;
        _configPage = null;
        _infoPage = null;
        _usbPage = null;
        SetWindowMainContent(BuildLayout());
        ApplyRequestedTheme();
        UpdateTitleBarAppearance();
        UpdatePageContent();
        RequestVmChipsRefresh();
        RefreshVmAdapterCards();
    }

    private async Task RestartForThemeChangeAsync()
    {
        if (_isThemeRestartInProgress)
        {
            return;
        }

        _isThemeRestartInProgress = true;

        try
        {
            _viewModel.PublishNotification("Theme-Wechsel wird angewendet...", "Info");

            await ShowThemeTransitionOverlayAsync("Layout wird aktualisiert …");

            var saveSucceeded = await TrySaveThemeConfigBeforeRestartAsync();
            if (!saveSucceeded)
            {
                _viewModel.PublishNotification("Theme-Einstellung konnte nicht gespeichert werden. Wechsel wird nur für diese Sitzung angewendet.", "Warning");
            }

            if (_themeTransitionStatusText is not null)
            {
                _themeTransitionStatusText.Text = "Design wird neu geladen …";
            }

            await Task.Delay(1000);

            if (Application.Current is App app)
            {
                await app.ReopenMainWindowForThemeChangeAsync();
                return;
            }

            _themeService.ApplyTheme(_viewModel.UiTheme);
            RebuildMainLayoutForTheme();
            _isMainLayoutLoaded = true;
            await Task.Delay(110);
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Theme transition failed; applying fallback rebuild.");

            try
            {
                _themeService.ApplyTheme(_viewModel.UiTheme);
                RebuildMainLayoutForTheme();
                _isMainLayoutLoaded = true;
            }
            catch (Exception fallbackEx)
            {
                Log.Error(fallbackEx, "Fallback rebuild after theme transition failure did not complete.");
            }
        }
        finally
        {
            await HideThemeTransitionOverlaySafeAsync();
            _isThemeRestartInProgress = false;
        }
    }

    private async Task<bool> TrySaveThemeConfigBeforeRestartAsync()
    {
        if (!_viewModel.HasPendingConfigChanges)
        {
            return true;
        }

        const int maxAttempts = 40;
        const int retryDelayMs = 100;

        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            if (_viewModel.SaveConfigCommand.CanExecute(null))
            {
                await _viewModel.SaveConfigCommand.ExecuteAsync(null);
                return !_viewModel.HasPendingConfigChanges;
            }

            await Task.Delay(retryDelayMs);
        }

        return false;
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
                System.IO.Path.Combine(AppContext.BaseDirectory, "Assets", "HyperTool.ico"),
                System.IO.Path.Combine(AppContext.BaseDirectory, "HyperTool.ico")
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
