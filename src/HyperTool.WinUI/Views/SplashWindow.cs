using HyperTool.WinUI.Helpers;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Animation;
using Microsoft.UI.Xaml.Media.Imaging;
using Microsoft.UI.Xaml.Shapes;
using System.Diagnostics;
using Windows.Graphics;
using Windows.UI;

namespace HyperTool.WinUI.Views;

public sealed class SplashWindow : Window
{
    private const int SplashWidth = MainWindow.DefaultWindowWidth;
    private const int SplashHeight = MainWindow.DefaultWindowHeight;

    private readonly Grid _root;
    private readonly ProgressBar _progressBar;
    private readonly TextBlock _statusText;
    private readonly Storyboard _ambientStoryboard = new();
    private readonly Stopwatch _lifetime = Stopwatch.StartNew();
    private readonly DispatcherQueueTimer _pseudoTimer;
    private readonly string _windowTitle;
    private readonly string _headline;
    private readonly string _iconUri;

    private readonly string[] _pseudoStatusMessages = LifecycleVisuals.StartupStatusMessages;

    private int _pseudoStatusIndex;
    private DateTime _lastExternalUpdateUtc = DateTime.UtcNow;

    public SplashWindow(string windowTitle = "HyperTool", string headline = "HyperTool", string iconUri = "ms-appx:///Assets/HyperTool.Icon.Transparent.png")
    {
        _windowTitle = string.IsNullOrWhiteSpace(windowTitle) ? "HyperTool" : windowTitle;
        _headline = string.IsNullOrWhiteSpace(headline) ? "HyperTool" : headline;
        _iconUri = string.IsNullOrWhiteSpace(iconUri) ? "ms-appx:///Assets/HyperTool.Icon.Transparent.png" : iconUri;

        Title = _windowTitle;
        ExtendsContentIntoTitleBar = false;
        DwmWindowHelper.ApplyRoundedCorners(this);

        var presenter = AppWindow.Presenter as OverlappedPresenter;
        if (presenter is not null)
        {
            presenter.SetBorderAndTitleBar(false, false);
            presenter.IsResizable = false;
            presenter.IsMinimizable = false;
            presenter.IsMaximizable = false;
            presenter.IsAlwaysOnTop = true;
        }

        try
        {
            AppWindow.Resize(new SizeInt32(SplashWidth, SplashHeight));
            AppWindow.IsShownInSwitchers = false;
            CenterOnCurrentDisplay();
        }
        catch
        {
        }

        _root = new Grid
        {
            Background = LifecycleVisuals.CreateRootBackgroundBrush()
        };

        BuildAmbientBands();
        BuildNetworkLayer();

        var splashVersionText = new TextBlock
        {
            Text = LifecycleVisuals.ResolveDisplayVersion(),
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Bottom,
            Margin = new Thickness(0, 0, 20, 14),
            FontSize = 12,
            Opacity = 0.72,
            Foreground = new SolidColorBrush(Color.FromArgb(0xC8, 0x9B, 0xB7, 0xD7))
        };
        _root.Children.Add(splashVersionText);

        var centerCard = new Border
        {
            Width = 540,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Padding = new Thickness(34),
            CornerRadius = new CornerRadius(26),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(LifecycleVisuals.CardBorder),
            Background = new SolidColorBrush(LifecycleVisuals.CardBackground)
        };

        var stack = new StackPanel { Spacing = 14 };

        var logoHost = new Grid
        {
            Width = 140,
            Height = 140,
            HorizontalAlignment = HorizontalAlignment.Center
        };

        var halo = new Ellipse
        {
            Width = 126,
            Height = 126,
            StrokeThickness = 1.8,
            Stroke = new SolidColorBrush(Color.FromArgb(0xC8, 0x79, 0xCD, 0xFF)),
            Opacity = 0.6,
            RenderTransformOrigin = new Windows.Foundation.Point(0.5, 0.5),
            RenderTransform = new ScaleTransform { ScaleX = 1.0, ScaleY = 1.0 }
        };
        logoHost.Children.Add(halo);

        var coreGlow = new Border
        {
            Width = 112,
            Height = 112,
            CornerRadius = new CornerRadius(56),
            Background = new SolidColorBrush(Color.FromArgb(0x80, 0x69, 0xC8, 0xFF)),
            Opacity = 0.34
        };
        logoHost.Children.Add(coreGlow);

        var logo = new Image
        {
            Source = new BitmapImage(new Uri(_iconUri)),
            Width = 76,
            Height = 76,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Opacity = 0.0,
            RenderTransformOrigin = new Windows.Foundation.Point(0.5, 0.5),
            RenderTransform = new ScaleTransform { ScaleX = 0.92, ScaleY = 0.92 }
        };
        logoHost.Children.Add(logo);

        stack.Children.Add(logoHost);
        stack.Children.Add(new TextBlock
        {
            Text = _headline,
            FontSize = 34,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            HorizontalAlignment = HorizontalAlignment.Center,
            Foreground = new SolidColorBrush(LifecycleVisuals.TextPrimary)
        });

        _statusText = new TextBlock
        {
            Text = "Initialisiere Hyper-V Umgebung …",
            HorizontalAlignment = HorizontalAlignment.Center,
            Foreground = new SolidColorBrush(LifecycleVisuals.TextSecondary),
            Margin = new Thickness(0, 2, 0, 0)
        };
        stack.Children.Add(_statusText);

        _progressBar = new ProgressBar
        {
            Height = 10,
            Minimum = 0,
            Maximum = 100,
            Value = 8,
            CornerRadius = new CornerRadius(5),
            Foreground = LifecycleVisuals.CreateProgressBrush(),
            Background = new SolidColorBrush(LifecycleVisuals.ProgressTrack)
        };
        stack.Children.Add(_progressBar);

        centerCard.Child = stack;
        _root.Children.Add(centerCard);
        Content = _root;

        BuildCenterAnimations(logo, halo, coreGlow, centerCard);

        _ambientStoryboard.Begin();

        _pseudoTimer = DispatcherQueue.CreateTimer();
        _pseudoTimer.Interval = TimeSpan.FromMilliseconds(LifecycleVisuals.SplashStatusCycleMs);
        _pseudoTimer.Tick += OnPseudoTick;
        _pseudoTimer.Start();
    }

    public void SetProgress(double ratio, string status)
    {
        var percent = Math.Clamp(ratio * 100.0, 0.0, 100.0);
        _progressBar.Value = percent;
        _statusText.Text = status;
        _lastExternalUpdateUtc = DateTime.UtcNow;
    }

    public async Task CompleteAndCloseAsync()
    {
        _pseudoTimer.Stop();

        SetProgress(1.0, "Starte HyperTool Oberfläche …");

        var remainingVisibleMs = LifecycleVisuals.SplashMinVisibleMs - _lifetime.ElapsedMilliseconds;
        if (remainingVisibleMs > 0)
        {
            await Task.Delay((int)remainingVisibleMs);
        }

        var fadeRoot = new DoubleAnimation
        {
            From = 1.0,
            To = 0.0,
            Duration = TimeSpan.FromMilliseconds(320),
            EnableDependentAnimation = true
        };

        var fadeStoryboard = new Storyboard();
        Storyboard.SetTarget(fadeRoot, _root);
        Storyboard.SetTargetProperty(fadeRoot, "Opacity");
        fadeStoryboard.Children.Add(fadeRoot);
        fadeStoryboard.Begin();

        await Task.Delay(360);

        try
        {
            Close();
        }
        catch
        {
        }
    }

    private void OnPseudoTick(DispatcherQueueTimer sender, object args)
    {
        if (DateTime.UtcNow - _lastExternalUpdateUtc < TimeSpan.FromMilliseconds(860))
        {
            return;
        }

        _progressBar.Value = Math.Min(92, _progressBar.Value + 0.9);

        _pseudoStatusIndex = (_pseudoStatusIndex + 1) % _pseudoStatusMessages.Length;
        _statusText.Text = _pseudoStatusMessages[_pseudoStatusIndex];
    }

    private void BuildAmbientBands()
    {
        var canvas = new Canvas { IsHitTestVisible = false };

        var band1 = CreateMovingBand(760, 54, 0.16, -18, -900, 1350, 3600, 0);
        var band2 = CreateMovingBand(620, 40, 0.11, -12, -840, 1280, 4100, 620);
        var band3 = CreateMovingBand(500, 34, 0.08, -14, -860, 1260, 4600, 1240);

        Canvas.SetTop(band1, 110);
        Canvas.SetTop(band2, 286);
        Canvas.SetTop(band3, 510);

        canvas.Children.Add(band1);
        canvas.Children.Add(band2);
        canvas.Children.Add(band3);

        _root.Children.Add(canvas);
    }

    private void BuildNetworkLayer()
    {
        var canvas = new Canvas { IsHitTestVisible = false };

        var nodes = new[]
        {
            new NodeSpec(184, 266, 14, true, "Host"),
            new NodeSpec(446, 276, 12, false, "Hyper-V"),
            new NodeSpec(702, 268, 14, true, "VM"),
            new NodeSpec(934, 318, 12, false, "Netz"),
            new NodeSpec(318, 186, 8, false, null),
            new NodeSpec(564, 176, 8, false, null),
            new NodeSpec(820, 194, 8, false, null),
            new NodeSpec(564, 382, 8, false, null),
            new NodeSpec(820, 404, 8, false, null)
        };

        foreach (var node in nodes)
        {
            var circle = new Ellipse
            {
                Width = node.Size,
                Height = node.Size,
                Fill = new SolidColorBrush(node.Highlight
                    ? Color.FromArgb(0xFF, 0x64, 0xC2, 0xFF)
                    : Color.FromArgb(0xD8, 0x6C, 0xB3, 0xF0)),
                Stroke = new SolidColorBrush(Color.FromArgb(0xA8, 0xC5, 0xE7, 0xFF)),
                StrokeThickness = node.Highlight ? 1.4 : 1,
                Opacity = node.Highlight ? 0.95 : 0.75
            };
            Canvas.SetLeft(circle, node.X - (node.Size / 2));
            Canvas.SetTop(circle, node.Y - (node.Size / 2));
            canvas.Children.Add(circle);

            if (!string.IsNullOrWhiteSpace(node.Label))
            {
                var label = new TextBlock
                {
                    Text = node.Label,
                    FontSize = 13,
                    FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                    Foreground = new SolidColorBrush(Color.FromArgb(0xFF, 0xB9, 0xD4, 0xF6)),
                    Opacity = 0.9
                };
                Canvas.SetLeft(label, node.X - 22);
                Canvas.SetTop(label, node.Y + 18);
                canvas.Children.Add(label);
            }

            var nodePulse = new DoubleAnimation
            {
                From = node.Highlight ? 0.55 : 0.42,
                To = node.Highlight ? 0.98 : 0.76,
                Duration = TimeSpan.FromMilliseconds(node.Highlight ? 1200 : 1600),
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(nodePulse, circle);
            Storyboard.SetTargetProperty(nodePulse, "Opacity");
            _ambientStoryboard.Children.Add(nodePulse);
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
                Stroke = new SolidColorBrush(Color.FromArgb(0x92, 0x67, 0xB2, 0xF4)),
                StrokeThickness = 1.1,
                Opacity = 0.0
            };
            canvas.Children.Add(line);

            var lineAppear = new DoubleAnimation
            {
                From = 0.0,
                To = 0.62,
                Duration = TimeSpan.FromMilliseconds(340),
                BeginTime = TimeSpan.FromMilliseconds(170 + (i * 140)),
                FillBehavior = FillBehavior.HoldEnd,
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(lineAppear, line);
            Storyboard.SetTargetProperty(lineAppear, "Opacity");
            _ambientStoryboard.Children.Add(lineAppear);

            var pulse = new Ellipse
            {
                Width = 4,
                Height = 4,
                Fill = new SolidColorBrush(Color.FromArgb(0xEE, 0x8E, 0xDB, 0xFF)),
                Opacity = 0.0
            };
            Canvas.SetLeft(pulse, a.X - 2);
            Canvas.SetTop(pulse, a.Y - 2);
            canvas.Children.Add(pulse);

            var pulseFade = new DoubleAnimation
            {
                From = 0.0,
                To = 0.95,
                Duration = TimeSpan.FromMilliseconds(220),
                BeginTime = TimeSpan.FromMilliseconds(560 + (i * 120)),
                FillBehavior = FillBehavior.HoldEnd,
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(pulseFade, pulse);
            Storyboard.SetTargetProperty(pulseFade, "Opacity");
            _ambientStoryboard.Children.Add(pulseFade);

            var pulseX = new DoubleAnimation
            {
                From = a.X - 2,
                To = b.X - 2,
                Duration = TimeSpan.FromMilliseconds(1360 + (i * 90)),
                BeginTime = TimeSpan.FromMilliseconds(760 + (i * 100)),
                RepeatBehavior = RepeatBehavior.Forever,
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(pulseX, pulse);
            Storyboard.SetTargetProperty(pulseX, "(Canvas.Left)");
            _ambientStoryboard.Children.Add(pulseX);

            var pulseY = new DoubleAnimation
            {
                From = a.Y - 2,
                To = b.Y - 2,
                Duration = TimeSpan.FromMilliseconds(1360 + (i * 90)),
                BeginTime = TimeSpan.FromMilliseconds(760 + (i * 100)),
                RepeatBehavior = RepeatBehavior.Forever,
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(pulseY, pulse);
            Storyboard.SetTargetProperty(pulseY, "(Canvas.Top)");
            _ambientStoryboard.Children.Add(pulseY);
        }

        for (var i = 0; i < 10; i++)
        {
            var spark = new Ellipse
            {
                Width = i % 3 == 0 ? 3.2 : 2.4,
                Height = i % 3 == 0 ? 3.2 : 2.4,
                Fill = new SolidColorBrush(Color.FromArgb(0x9A, 0x85, 0xCC, 0xFF)),
                Opacity = 0.26,
                RenderTransform = new TranslateTransform()
            };

            var x = 190 + (i * 74);
            var y = 138 + ((i % 4) * 70);
            Canvas.SetLeft(spark, x);
            Canvas.SetTop(spark, y);
            canvas.Children.Add(spark);

            var moveX = new DoubleAnimation
            {
                From = -4,
                To = 7,
                Duration = TimeSpan.FromMilliseconds(2600 + (i * 190)),
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(moveX, spark);
            Storyboard.SetTargetProperty(moveX, "(UIElement.RenderTransform).(TranslateTransform.X)");
            _ambientStoryboard.Children.Add(moveX);

            var moveY = new DoubleAnimation
            {
                From = -3,
                To = 6,
                Duration = TimeSpan.FromMilliseconds(2800 + (i * 140)),
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(moveY, spark);
            Storyboard.SetTargetProperty(moveY, "(UIElement.RenderTransform).(TranslateTransform.Y)");
            _ambientStoryboard.Children.Add(moveY);
        }

        _root.Children.Add(canvas);
    }

    private Rectangle CreateMovingBand(
        double width,
        double height,
        double opacity,
        double rotation,
        double fromX,
        double toX,
        int durationMs,
        int beginMs)
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
        _ambientStoryboard.Children.Add(move);

        return band;
    }

    private void BuildCenterAnimations(Image logo, Ellipse halo, Border glow, Border centerCard)
    {
        var story = new Storyboard { RepeatBehavior = RepeatBehavior.Forever };

        var logoFadeIn = new DoubleAnimation
        {
            From = 0.0,
            To = 1.0,
            Duration = TimeSpan.FromMilliseconds(420),
            BeginTime = TimeSpan.FromMilliseconds(90),
            FillBehavior = FillBehavior.HoldEnd,
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(logoFadeIn, logo);
        Storyboard.SetTargetProperty(logoFadeIn, "Opacity");
        _ambientStoryboard.Children.Add(logoFadeIn);

        var logoScaleInX = new DoubleAnimation
        {
            From = 0.92,
            To = 1.0,
            Duration = TimeSpan.FromMilliseconds(420),
            BeginTime = TimeSpan.FromMilliseconds(90),
            FillBehavior = FillBehavior.HoldEnd,
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(logoScaleInX, logo);
        Storyboard.SetTargetProperty(logoScaleInX, "(UIElement.RenderTransform).(ScaleTransform.ScaleX)");
        _ambientStoryboard.Children.Add(logoScaleInX);

        var logoScaleInY = new DoubleAnimation
        {
            From = 0.92,
            To = 1.0,
            Duration = TimeSpan.FromMilliseconds(420),
            BeginTime = TimeSpan.FromMilliseconds(90),
            FillBehavior = FillBehavior.HoldEnd,
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(logoScaleInY, logo);
        Storyboard.SetTargetProperty(logoScaleInY, "(UIElement.RenderTransform).(ScaleTransform.ScaleY)");
        _ambientStoryboard.Children.Add(logoScaleInY);

        var haloScale = new DoubleAnimation
        {
            From = 1.0,
            To = 1.12,
            Duration = TimeSpan.FromMilliseconds(1140),
            AutoReverse = true,
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(haloScale, halo);
        Storyboard.SetTargetProperty(haloScale, "(UIElement.RenderTransform).(ScaleTransform.ScaleX)");
        story.Children.Add(haloScale);

        var haloScaleY = new DoubleAnimation
        {
            From = 1.0,
            To = 1.12,
            Duration = TimeSpan.FromMilliseconds(1140),
            AutoReverse = true,
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(haloScaleY, halo);
        Storyboard.SetTargetProperty(haloScaleY, "(UIElement.RenderTransform).(ScaleTransform.ScaleY)");
        story.Children.Add(haloScaleY);

        var haloFade = new DoubleAnimation
        {
            From = 0.34,
            To = 0.88,
            Duration = TimeSpan.FromMilliseconds(1140),
            AutoReverse = true,
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(haloFade, halo);
        Storyboard.SetTargetProperty(haloFade, "Opacity");
        story.Children.Add(haloFade);

        var glowPulse = new DoubleAnimation
        {
            From = 0.24,
            To = 0.5,
            Duration = TimeSpan.FromMilliseconds(980),
            AutoReverse = true,
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(glowPulse, glow);
        Storyboard.SetTargetProperty(glowPulse, "Opacity");
        story.Children.Add(glowPulse);

        var cardPulse = new DoubleAnimation
        {
            From = 0.93,
            To = 1.0,
            Duration = TimeSpan.FromMilliseconds(1320),
            AutoReverse = true,
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(cardPulse, centerCard);
        Storyboard.SetTargetProperty(cardPulse, "Opacity");
        story.Children.Add(cardPulse);

        story.Begin();
    }

    private void CenterOnCurrentDisplay()
    {
        try
        {
            var displayArea = DisplayArea.GetFromWindowId(AppWindow.Id, DisplayAreaFallback.Primary);
            var work = displayArea.WorkArea;
            var x = work.X + Math.Max(0, (work.Width - SplashWidth) / 2);
            var y = work.Y + Math.Max(0, (work.Height - SplashHeight) / 2);
            AppWindow.Move(new PointInt32(x, y));
        }
        catch
        {
        }
    }

    private readonly record struct NodeSpec(double X, double Y, double Size, bool Highlight, string? Label);
}
