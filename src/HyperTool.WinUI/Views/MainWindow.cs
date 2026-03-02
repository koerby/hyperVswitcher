using HyperTool.Services;
using HyperTool.Models;
using HyperTool.ViewModels;
using HyperTool.WinUI.Helpers;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Animation;
using Microsoft.UI.Xaml.Shapes;
using Serilog;
using System.Linq;
using System.ComponentModel;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Media;
using System.Windows.Input;
using Windows.Graphics;
using Windows.UI;

namespace HyperTool.WinUI.Views;

public sealed class MainWindow : Window
{
    public const int DefaultWindowWidth = 1400;
    public const int DefaultWindowHeight = 940;

    private readonly IThemeService _themeService;
    private readonly MainViewModel _viewModel;
    private readonly List<Button> _navButtons = [];
    private readonly StackPanel _vmChipPanel = new() { Orientation = Orientation.Horizontal, Spacing = 8 };
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
    private readonly ListView _checkpointListView = new();
    private readonly StackPanel _vmAdapterCardsPanel = new() { Spacing = 10 };
    private readonly RotateTransform _logoRotateTransform = new();
    private UIElement? _vmPage;
    private UIElement? _snapshotsPage;
    private UIElement? _configPage;
    private UIElement? _infoPage;
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
        _themeService.ApplyTheme(_viewModel.UiTheme);
        ApplyRequestedTheme();
        UpdateTitleBarAppearance();

        if (!showStartupSplash)
        {
            RefreshVmChips();
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
        RefreshVmChips();
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
        var headerGrid = new Grid();
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

        var chipScroller = new ScrollViewer
        {
            HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Content = _vmChipPanel
        };
        titleStack.Children.Add(chipScroller);
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

        Grid.SetColumn(titleActions, 1);
        headerGrid.Children.Add(titleActions);
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
        sidebarStack.Children.Add(CreateNavButton("📷", "Snapshots", 1));
        sidebarStack.Children.Add(CreateNavButton("⚙", "Einstellungen", 2));
        sidebarStack.Children.Add(CreateNavButton("ℹ", "Info", 3));
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
        var stateText = new TextBlock();
        stateText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.SelectedVmState)) });
        Grid.SetColumn(stateText, 3);
        selectedVmGrid.Children.Add(stateText);
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
        summaryButtons.Children.Add(CreateIconButton("📄", "Logdatei öffnen", _viewModel.OpenLogFileCommand, compact: true));
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

        _checkpointListView.ItemsSource = _viewModel.AvailableCheckpoints;
        _checkpointListView.SelectionMode = ListViewSelectionMode.Single;
        _checkpointListView.IsItemClickEnabled = true;
        _checkpointListView.DisplayMemberPath = nameof(HyperVCheckpointInfo.Name);
        _checkpointListView.ItemClick += (_, args) =>
        {
            if (args.ClickedItem is HyperVCheckpointInfo checkpoint)
            {
                _viewModel.SelectedCheckpoint = checkpoint;

                var selectedNode = _viewModel.AvailableCheckpointTree
                    .SelectMany(FlattenCheckpointTree)
                    .FirstOrDefault(node => node.Checkpoint.Id == checkpoint.Id);
                _viewModel.SelectedCheckpointNode = selectedNode;
            }
        };
        listBorder.Child = _checkpointListView;
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
        headingCard.Child = headingStack;
        root.Children.Add(headingCard);

        var topBar = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush,
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(14)
        };

        var topGrid = new Grid { ColumnSpacing = 10, VerticalAlignment = VerticalAlignment.Center };
        topGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        topGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        topGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        topGrid.Children.Add(CreateIconButton("💾", "Speichern", _viewModel.SaveConfigCommand));
        var reloadButton = CreateIconButton("⟳", "Neu laden", _viewModel.ReloadConfigCommand);
        Grid.SetColumn(reloadButton, 1);
        topGrid.Children.Add(reloadButton);

        var configPathText = new TextBlock
        {
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Opacity = 0.85,
            MaxWidth = 560
        };
        configPathText.Foreground = Application.Current.Resources["TextPrimaryBrush"] as Brush;
        configPathText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.ConfigPath)) });
        var configPathWrap = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        configPathWrap.Children.Add(new TextBlock
        {
            Text = "Aktive Config:",
            VerticalAlignment = VerticalAlignment.Center,
            Opacity = 0.9,
            Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
        });
        configPathWrap.Children.Add(configPathText);
        Grid.SetColumn(configPathWrap, 2);
        topGrid.Children.Add(configPathWrap);
        topBar.Child = topGrid;
        root.Children.Add(topBar);

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

        var defaultVmText = new TextBlock { TextWrapping = TextWrapping.Wrap };
        defaultVmText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.DefaultVmName)) });
        vmOverviewStack.Children.Add(new TextBlock { Text = "Default VM", Opacity = 0.9, Foreground = Application.Current.Resources["TextMutedBrush"] as Brush, Margin = new Thickness(0, 2, 0, 0) });
        vmOverviewStack.Children.Add(defaultVmText);

        var trayAdapterRow = new Grid { ColumnSpacing = 8, VerticalAlignment = VerticalAlignment.Center };
        trayAdapterRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(132) });
        trayAdapterRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        trayAdapterRow.Children.Add(new TextBlock
        {
            Text = "Tray Adapter",
            VerticalAlignment = VerticalAlignment.Center,
            Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
        });
        var trayAdapterCombo = CreateStyledComboBox();
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

        var vmButtonsGrid = new Grid { ColumnSpacing = 8, Margin = new Thickness(0, 2, 0, 0) };
        vmButtonsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        vmButtonsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        vmButtonsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var setDefaultButton = CreateIconButton("⭐", "Ausgewählte als Default", _viewModel.SetDefaultVmCommand);
        setDefaultButton.HorizontalAlignment = HorizontalAlignment.Stretch;
        vmButtonsGrid.Children.Add(setDefaultButton);

        var exportButton = CreateIconButton("💾", "Ausgewählte VM exportieren", _viewModel.ExportSelectedVmCommand);
        exportButton.HorizontalAlignment = HorizontalAlignment.Stretch;
        Grid.SetColumn(exportButton, 1);
        vmButtonsGrid.Children.Add(exportButton);

        var importButton = CreateIconButton("📥", "VM importieren", _viewModel.ImportVmCommand);
        importButton.HorizontalAlignment = HorizontalAlignment.Stretch;
        Grid.SetColumn(importButton, 2);
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

        var quickTogglesGrid = new Grid { ColumnSpacing = 10, RowSpacing = 6 };
        quickTogglesGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        quickTogglesGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        quickTogglesGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        quickTogglesGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var trayMenuCheck = CreateCheckBox("Tasktray-Menü aktiv", () => _viewModel.UiEnableTrayMenu, value => _viewModel.UiEnableTrayMenu = value);
        Grid.SetColumn(trayMenuCheck, 0);
        Grid.SetRow(trayMenuCheck, 0);
        quickTogglesGrid.Children.Add(trayMenuCheck);

        var startMinCheck = CreateCheckBox("Beim Start minimiert", () => _viewModel.UiStartMinimized, value => _viewModel.UiStartMinimized = value);
        Grid.SetColumn(startMinCheck, 1);
        Grid.SetRow(startMinCheck, 0);
        quickTogglesGrid.Children.Add(startMinCheck);

        var startWithWindowsCheck = CreateCheckBox("Mit Windows starten", () => _viewModel.UiStartWithWindows, value => _viewModel.UiStartWithWindows = value);
        Grid.SetColumn(startWithWindowsCheck, 0);
        Grid.SetRow(startWithWindowsCheck, 1);
        quickTogglesGrid.Children.Add(startWithWindowsCheck);

        var updateCheck = CreateCheckBox("Beim Start auf Updates prüfen", () => _viewModel.UpdateCheckOnStartup, value => _viewModel.UpdateCheckOnStartup = value);
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
        Grid.SetColumn(themeToggle, 1);
        themeRow.Children.Add(themeToggle);

        var themeText = new TextBlock
        {
            Text = string.Equals(_viewModel.UiTheme, "Dark", StringComparison.OrdinalIgnoreCase) ? "Dunkles Theme" : "Helles Theme",
            VerticalAlignment = VerticalAlignment.Center,
            Foreground = Application.Current.Resources["TextMutedBrush"] as Brush,
            Opacity = 0.95
        };
        themeToggle.Toggled += (_, _) =>
        {
            themeText.Text = themeToggle.IsOn ? "Dunkles Theme" : "Helles Theme";
        };
        Grid.SetColumn(themeText, 2);
        themeRow.Children.Add(themeText);
        systemStack.Children.Add(themeRow);

        var hostRow = new Grid { ColumnSpacing = 8, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 2, 0, 0) };
        hostRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
        hostRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        hostRow.Children.Add(new TextBlock
        {
            Text = "VMConnect Host",
            VerticalAlignment = VerticalAlignment.Center,
            Foreground = Application.Current.Resources["TextMutedBrush"] as Brush
        });
        var hostTextBox = CreateStyledTextBox(_viewModel.VmConnectComputerName, "z. B. KAI-PC");
        hostTextBox.TextChanged += (_, _) => _viewModel.VmConnectComputerName = hostTextBox.Text;
        Grid.SetColumn(hostTextBox, 1);
        hostRow.Children.Add(hostTextBox);
        systemStack.Children.Add(hostRow);

        systemSection.Child = systemStack;
        root.Children.Add(systemSection);
        return new ScrollViewer { Content = root };
    }

    private static IEnumerable<HyperVCheckpointTreeItem> FlattenCheckpointTree(HyperVCheckpointTreeItem root)
    {
        yield return root;

        foreach (var child in root.Children.SelectMany(FlattenCheckpointTree))
        {
            yield return child;
        }
    }

    private UIElement BuildInfoPage()
    {
        var panel = new StackPanel { Spacing = 10 };

        var heading = new TextBlock { Text = "Info", FontSize = 20, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold };
        panel.Children.Add(heading);

        var versionWrap = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        versionWrap.Children.Add(new TextBlock { Text = "Version:", Opacity = 0.9 });
        var versionText = new TextBlock { Opacity = 0.9 };
        versionText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.AppVersion)) });
        versionWrap.Children.Add(versionText);
        panel.Children.Add(versionWrap);

        var updateWrap = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        updateWrap.Children.Add(new TextBlock { Text = "Update-Status:", Opacity = 0.9 });
        var updateText = new TextBlock { TextWrapping = TextWrapping.Wrap, Opacity = 0.9 };
        updateText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.UpdateStatus)) });
        updateWrap.Children.Add(updateText);
        panel.Children.Add(updateWrap);

        var infoCard = new Border
        {
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PageBackgroundBrush"] as Brush,
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(10)
        };

        var infoStack = new StackPanel { Spacing = 6 };
        infoStack.Children.Add(new TextBlock { Text = "Projekt", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
        infoStack.Children.Add(new TextBlock { Text = "HyperTool wird über GitHub Releases verteilt. Hier findest du Version, Update-Status und Release-Links.", TextWrapping = TextWrapping.Wrap, Opacity = 0.85 });

        var linksGrid = new Grid { ColumnSpacing = 8, Margin = new Thickness(0, 8, 0, 0) };
        linksGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
        linksGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        linksGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        linksGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        linksGrid.Children.Add(new TextBlock { Text = "GitHub Owner", Opacity = 0.8 });
        var ownerText = new TextBlock();
        ownerText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.GithubOwner)) });
        Grid.SetColumn(ownerText, 1);
        linksGrid.Children.Add(ownerText);
        var repoLabel = new TextBlock { Text = "GitHub Repo", Opacity = 0.8, Margin = new Thickness(0, 8, 0, 0) };
        Grid.SetRow(repoLabel, 1);
        linksGrid.Children.Add(repoLabel);
        var repoText = new TextBlock { Margin = new Thickness(0, 8, 0, 0) };
        repoText.SetBinding(TextBlock.TextProperty, new Binding { Source = _viewModel, Path = new PropertyPath(nameof(MainViewModel.GithubRepo)) });
        Grid.SetColumn(repoText, 1);
        Grid.SetRow(repoText, 1);
        linksGrid.Children.Add(repoText);
        infoStack.Children.Add(linksGrid);
        infoCard.Child = infoStack;
        panel.Children.Add(infoCard);

        var buttonRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        buttonRow.Children.Add(CreateIconButton("🛰", "Update prüfen", _viewModel.CheckForUpdatesCommand));
        buttonRow.Children.Add(CreateIconButton("⬇", "Update installieren", _viewModel.InstallUpdateCommand));
        buttonRow.Children.Add(CreateIconButton("🌐", "Changelog / Release", _viewModel.OpenReleasePageCommand));
        panel.Children.Add(buttonRow);

        return new ScrollViewer { Content = panel };
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
                Foreground = Application.Current.Resources["TextPrimaryBrush"] as Brush,
                FontWeight = Microsoft.UI.Text.FontWeights.Medium,
                TextWrapping = TextWrapping.Wrap
            }
        };
        checkBox.Checked += (_, _) => setter(true);
        checkBox.Unchecked += (_, _) => setter(false);
        return checkBox;
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
            MinHeight = 76,
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush,
            HorizontalAlignment = HorizontalAlignment.Stretch
        };

        var content = new StackPanel { Spacing = 4, HorizontalAlignment = HorizontalAlignment.Center };
        var navIconHost = new Grid
        {
            Width = 30,
            Height = 30,
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
                FontSize = 20,
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

    private void RefreshVmChips()
    {
        _vmChipPanel.Children.Clear();

        foreach (var vm in _viewModel.AvailableVms)
        {
            _vmChipPanel.Children.Add(CreateVmChip(vm));
        }
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

    private Button CreateVmChip(VmDefinition vm)
    {
        var chip = new Button
        {
            Padding = new Thickness(11, 8, 12, 8),
            MinWidth = 240,
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

        var iconBadge = new Border
        {
            Width = 28,
            Height = 28,
            CornerRadius = new CornerRadius(8),
            VerticalAlignment = VerticalAlignment.Center,
            BorderThickness = new Thickness(1),
            BorderBrush = chip.BorderBrush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush,
            Child = new TextBlock
            {
                Text = "🖥",
                FontSize = 15,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                TextAlignment = TextAlignment.Center
            }
        };

        var text = new TextBlock
        {
            Text = $"{vm.DisplayLabel}",
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

    private void UpdatePageContent()
    {
        try
        {
            var content = _viewModel.SelectedMenuIndex switch
            {
                1 => _snapshotsPage ??= BuildSnapshotsPage(),
                2 => _configPage ??= BuildConfigPage(),
                3 => _infoPage ??= BuildInfoPage(),
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
        flyout.Items.Add(CreateVmMenuItem("Open Console", () => _viewModel.OpenConsoleByNameCommand.ExecuteAsync(vm.Name)));
        flyout.Items.Add(CreateVmMenuItem("Snapshot", () => _viewModel.CreateSnapshotByNameCommand.ExecuteAsync(vm.Name)));
        flyout.ShowAt(element);
        e.Handled = true;
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

    private void OnWindowClosed(object sender, WindowEventArgs args)
    {
        _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
        _viewModel.AvailableVms.CollectionChanged -= OnAvailableVmsCollectionChanged;
        _viewModel.AvailableVmNetworkAdapters.CollectionChanged -= OnVmNetworkAdaptersCollectionChanged;
        _viewModel.AvailableSwitches.CollectionChanged -= OnVmNetworkAdaptersCollectionChanged;
    }

    private void OnAvailableVmsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        RunOnUiThread(() =>
        {
            RefreshVmChips();
        });
    }

    private void OnVmNetworkAdaptersCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        RunOnUiThread(() =>
        {
            RefreshVmAdapterCards();
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
            if (_isMainLayoutLoaded)
            {
                RefreshVmChips();
                RefreshVmAdapterCards();
            }
        }

        if (string.Equals(e.PropertyName, nameof(MainViewModel.SelectedVmAdapterSwitchDisplay), StringComparison.Ordinal)
            || string.Equals(e.PropertyName, nameof(MainViewModel.UiTheme), StringComparison.Ordinal))
        {
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
        _configPage = null;
        _infoPage = null;
        SetWindowMainContent(BuildLayout());
        ApplyRequestedTheme();
        UpdateTitleBarAppearance();
        UpdatePageContent();
        RefreshVmChips();
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
