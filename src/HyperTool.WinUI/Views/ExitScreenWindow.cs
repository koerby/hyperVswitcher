using HyperTool.WinUI.Helpers;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Animation;
using Microsoft.UI.Xaml.Media.Imaging;
using Microsoft.UI.Xaml.Shapes;
using System.Linq;
using Windows.Graphics;
using Windows.UI;

namespace HyperTool.WinUI.Views;

public sealed class ExitScreenWindow : Window
{
    private readonly Grid _root;
    private readonly TextBlock _statusText;
    private readonly Ellipse _outerHalo;
    private readonly Border _coreHalo;
    private readonly Rectangle _progressFill;
    private readonly Rectangle _progressShimmer;
    private readonly List<Line> _networkLines = [];
    private readonly List<Ellipse> _impulses = [];
    private readonly List<Ellipse> _particles = [];
    private readonly string _windowTitle;
    private readonly string _headline;
    private readonly string _iconUri;

    public ExitScreenWindow(string windowTitle = "HyperTool", string headline = "HyperTool", string iconUri = "ms-appx:///Assets/HyperTool.Icon.Transparent.png")
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
        }

        try
        {
            AppWindow.Resize(new SizeInt32(MainWindow.DefaultWindowWidth, MainWindow.DefaultWindowHeight));
            AppWindow.IsShownInSwitchers = false;
        }
        catch
        {
        }

        _root = new Grid
        {
            Opacity = 0,
            Background = LifecycleVisuals.CreateRootBackgroundBrush()
        };

        var focusLayerPrimary = new Ellipse
        {
            Width = 820,
            Height = 820,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Fill = LifecycleVisuals.CreateCenterFocusBrush(LifecycleVisuals.BackgroundFocusSecondary)
        };
        _root.Children.Add(focusLayerPrimary);

        var focusLayerSecondary = new Ellipse
        {
            Width = 620,
            Height = 620,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(88, -52, -88, 52),
            Fill = LifecycleVisuals.CreateCenterFocusBrush(LifecycleVisuals.BackgroundFocusTertiary),
            Opacity = 0.54
        };
        _root.Children.Add(focusLayerSecondary);

        BuildNetworkLayer();
        BuildAmbientBands();

        var card = new Border
        {
            Width = 520,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Padding = new Thickness(30, 28, 30, 24),
            CornerRadius = new CornerRadius(24),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(LifecycleVisuals.CardBorder),
            Background = LifecycleVisuals.CreateCardSurfaceBrush(),
            Shadow = new ThemeShadow()
        };

        var innerFrame = new Border
        {
            CornerRadius = new CornerRadius(20),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(LifecycleVisuals.CardInnerOutline),
            Background = LifecycleVisuals.CreateCardInnerBrush(),
            Padding = new Thickness(20, 18, 20, 16)
        };

        var stack = new StackPanel { Spacing = 12 };

        var logoHost = new Grid { Width = 122, Height = 122, HorizontalAlignment = HorizontalAlignment.Center };
        _outerHalo = new Ellipse
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
        };
        logoHost.Children.Add(_outerHalo);

        var middleRing = new Ellipse
        {
            Width = 108,
            Height = 108,
            StrokeThickness = 1.1,
            Stroke = new SolidColorBrush(Color.FromArgb(0x78, 0x88, 0xD1, 0xFF)),
            Opacity = 0.36,
            RenderTransformOrigin = new Windows.Foundation.Point(0.5, 0.5),
            RenderTransform = new ScaleTransform { ScaleX = 1.0, ScaleY = 1.0 }
        };
        logoHost.Children.Add(middleRing);

        _coreHalo = new Border
        {
            Width = 98,
            Height = 98,
            CornerRadius = new CornerRadius(49),
            Background = new SolidColorBrush(Color.FromArgb(0x64, 0x66, 0xC3, 0xFF)),
            Opacity = 0.28
        };
        logoHost.Children.Add(_coreHalo);

        logoHost.Children.Add(new Image
        {
            Source = new BitmapImage(new Uri(_iconUri)),
            Width = 68,
            Height = 68,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        });

        stack.Children.Add(logoHost);
        stack.Children.Add(new TextBlock
        {
            Text = _headline,
            FontSize = 30,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            HorizontalAlignment = HorizontalAlignment.Center,
            Foreground = new SolidColorBrush(LifecycleVisuals.TextPrimary)
        });

        _statusText = new TextBlock
        {
            Text = LifecycleVisuals.ShutdownStatusMessages[0],
            HorizontalAlignment = HorizontalAlignment.Center,
            Foreground = new SolidColorBrush(LifecycleVisuals.TextSecondary),
            FontSize = 13,
            Margin = new Thickness(0, 6, 0, 2),
            Opacity = 0.94
        };
        stack.Children.Add(_statusText);

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
        _progressFill = new Rectangle
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            RadiusX = 4,
            RadiusY = 4,
            Fill = LifecycleVisuals.CreateProgressBrush(),
            RenderTransformOrigin = new Windows.Foundation.Point(0, 0.5),
            RenderTransform = new ScaleTransform { ScaleX = 1.0, ScaleY = 1.0 }
        };

        _progressShimmer = new Rectangle
        {
            Width = 84,
            HorizontalAlignment = HorizontalAlignment.Left,
            RadiusX = 4,
            RadiusY = 4,
            Opacity = 0.20,
            Fill = new LinearGradientBrush
            {
                StartPoint = new Windows.Foundation.Point(0, 0.5),
                EndPoint = new Windows.Foundation.Point(1, 0.5),
                GradientStops =
                {
                    new GradientStop { Color = Color.FromArgb(0x00, 0xDF, 0xF4, 0xFF), Offset = 0 },
                    new GradientStop { Color = Color.FromArgb(0xC5, 0xDF, 0xF4, 0xFF), Offset = 0.5 },
                    new GradientStop { Color = Color.FromArgb(0x00, 0xDF, 0xF4, 0xFF), Offset = 1 }
                }
            },
            RenderTransform = new TranslateTransform { X = -80 }
        };

        progressLayer.Children.Add(_progressFill);
        progressLayer.Children.Add(_progressShimmer);
        progressHost.Child = progressLayer;
        stack.Children.Add(progressHost);

        innerFrame.Child = stack;
        card.Child = innerFrame;
        _root.Children.Add(card);

        var vignette = new Rectangle
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment = VerticalAlignment.Stretch,
            Fill = LifecycleVisuals.CreateVignetteBrush(0x64)
        };
        _root.Children.Add(vignette);

        Content = _root;

        StartAmbientShutdownAnimation();
        FadeInRoot();
    }

    public void ConfigureBounds(PointInt32 position, SizeInt32 size)
    {
        try
        {
            var width = Math.Max(size.Width, MainWindow.DefaultWindowWidth);
            var height = Math.Max(size.Height, MainWindow.DefaultWindowHeight);
            AppWindow.Resize(new SizeInt32(width, height));
            AppWindow.Move(position);
        }
        catch
        {
        }
    }

    public async Task PlayAndCloseAsync()
    {
        await Task.Delay(220);
        await AnimateStatusTextAsync(LifecycleVisuals.ShutdownStatusMessages[1]);

        await Task.Delay(300);
        await AnimateLayerOpacityAsync(_networkLines.Cast<UIElement>(), 0.28, 440);
        await AnimateLayerOpacityAsync(_impulses.Cast<UIElement>(), 0.14, 420);

        await AnimateStatusTextAsync(LifecycleVisuals.ShutdownStatusMessages[2]);
        await Task.Delay(280);

        await AnimateLayerOpacityAsync(_impulses.Cast<UIElement>(), 0.04, 460);
        await AnimateStatusTextAsync(LifecycleVisuals.ShutdownStatusMessages[3]);

        await Task.Delay(320);
        await AnimateLayerOpacityAsync(_networkLines.Cast<UIElement>(), 0.04, 620);
        await AnimateLayerOpacityAsync(_particles.Cast<UIElement>(), 0.02, 520);

        var calmStoryboard = new Storyboard();

        var haloFade = new DoubleAnimation
        {
            From = _outerHalo.Opacity,
            To = 0.08,
            Duration = TimeSpan.FromMilliseconds(720),
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(haloFade, _outerHalo);
        Storyboard.SetTargetProperty(haloFade, "Opacity");
        calmStoryboard.Children.Add(haloFade);

        var coreFade = new DoubleAnimation
        {
            From = _coreHalo.Opacity,
            To = 0.04,
            Duration = TimeSpan.FromMilliseconds(720),
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(coreFade, _coreHalo);
        Storyboard.SetTargetProperty(coreFade, "Opacity");
        calmStoryboard.Children.Add(coreFade);

        var lineShrink = new DoubleAnimation
        {
            From = 1.0,
            To = 0.03,
            Duration = TimeSpan.FromMilliseconds(860),
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(lineShrink, _progressFill);
        Storyboard.SetTargetProperty(lineShrink, "(UIElement.RenderTransform).(ScaleTransform.ScaleX)");
        calmStoryboard.Children.Add(lineShrink);

        var shimmerFade = new DoubleAnimation
        {
            From = _progressShimmer.Opacity,
            To = 0.0,
            Duration = TimeSpan.FromMilliseconds(520),
            EasingFunction = LifecycleVisuals.CreateEaseOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(shimmerFade, _progressShimmer);
        Storyboard.SetTargetProperty(shimmerFade, "Opacity");
        calmStoryboard.Children.Add(shimmerFade);

        calmStoryboard.Begin();
        await Task.Delay(700);

        await AnimateStatusTextAsync(LifecycleVisuals.ShutdownStatusMessages[4]);
        await Task.Delay(220);

        var fadeOut = new DoubleAnimation
        {
            From = 1.0,
            To = 0.0,
            Duration = TimeSpan.FromMilliseconds(520),
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };

        var fadeStory = new Storyboard();
        Storyboard.SetTarget(fadeOut, _root);
        Storyboard.SetTargetProperty(fadeOut, "Opacity");
        fadeStory.Children.Add(fadeOut);
        fadeStory.Begin();

        await Task.Delay(560);

        try
        {
            Close();
        }
        catch
        {
        }
    }

    private void BuildAmbientBands()
    {
        var canvas = new Canvas { IsHitTestVisible = false };

        var band1 = CreateMovingBand(660, 44, 0.06, -14, -800, 1180, 7000, 0);
        var band2 = CreateMovingBand(520, 32, 0.04, -10, -760, 1140, 7600, 1200);

        Canvas.SetTop(band1, 152);
        Canvas.SetTop(band2, 476);

        canvas.Children.Add(band1);
        canvas.Children.Add(band2);
        _root.Children.Add(canvas);
    }

    private void BuildNetworkLayer()
    {
        var canvas = new Canvas { IsHitTestVisible = false };

        var nodes = new[]
        {
            new ExitNodeSpec(new Windows.Foundation.Point(236, 260), 9.8, "Host", true),
            new ExitNodeSpec(new Windows.Foundation.Point(364, 298), 8.4, "Mgmt", false),
            new ExitNodeSpec(new Windows.Foundation.Point(504, 316), 8.6, "Hyper-V", true),
            new ExitNodeSpec(new Windows.Foundation.Point(642, 294), 8.4, null, false),
            new ExitNodeSpec(new Windows.Foundation.Point(780, 274), 9.4, "VM", true),
            new ExitNodeSpec(new Windows.Foundation.Point(914, 312), 8.2, "Client", false),
            new ExitNodeSpec(new Windows.Foundation.Point(1006, 346), 8.0, "Target", false),
            new ExitNodeSpec(new Windows.Foundation.Point(422, 388), 7.2, null, false),
            new ExitNodeSpec(new Windows.Foundation.Point(838, 392), 7.2, null, false)
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
                    ? Color.FromArgb(0xD6, 0x67, 0xBF, 0xF8)
                    : LifecycleVisuals.NodeColor),
                Stroke = new SolidColorBrush(LifecycleVisuals.NodeStroke),
                StrokeThickness = nodeSpec.IsPrimary ? 1.0 : 0.9,
                Opacity = nodeSpec.IsPrimary ? 0.66 : 0.52
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
                    Opacity = 0.58,
                    Foreground = new SolidColorBrush(Color.FromArgb(0xC2, 0xA6, 0xC4, 0xE3))
                };
                Canvas.SetLeft(label, nodeSpec.Point.X - 24);
                Canvas.SetTop(label, nodeSpec.Point.Y + 13);
                canvas.Children.Add(label);
            }
        }

        foreach (var (from, to) in links)
        {
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
                Opacity = 0.34
            };

            _networkLines.Add(line);
            canvas.Children.Add(line);

            var pulse = new Ellipse
            {
                Width = 3.0,
                Height = 3.0,
                Fill = new SolidColorBrush(LifecycleVisuals.PulseColor),
                Opacity = 0.28
            };
            Canvas.SetLeft(pulse, a.X - 1.5);
            Canvas.SetTop(pulse, a.Y - 1.5);
            _impulses.Add(pulse);
            canvas.Children.Add(pulse);
        }

        for (var i = 0; i < 4; i++)
        {
            var particle = new Ellipse
            {
                Width = i % 2 == 0 ? 2.4 : 2.0,
                Height = i % 2 == 0 ? 2.4 : 2.0,
                Fill = new SolidColorBrush(LifecycleVisuals.ParticleColor),
                Opacity = 0.12,
                RenderTransform = new TranslateTransform()
            };

            Canvas.SetLeft(particle, 266 + (i * 184));
            Canvas.SetTop(particle, 216 + ((i % 2) * 148));
            _particles.Add(particle);
            canvas.Children.Add(particle);
        }

        _root.Children.Add(canvas);
    }

    private readonly record struct ExitNodeSpec(Windows.Foundation.Point Point, double Size, string? Label, bool IsPrimary);

    private void StartAmbientShutdownAnimation()
    {
        var story = new Storyboard { RepeatBehavior = RepeatBehavior.Forever };

        var haloPulse = new DoubleAnimation
        {
            From = 0.16,
            To = 0.36,
            Duration = TimeSpan.FromMilliseconds(2200),
            AutoReverse = true,
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(haloPulse, _outerHalo);
        Storyboard.SetTargetProperty(haloPulse, "Opacity");
        story.Children.Add(haloPulse);

        var corePulse = new DoubleAnimation
        {
            From = 0.12,
            To = 0.30,
            Duration = TimeSpan.FromMilliseconds(2000),
            AutoReverse = true,
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(corePulse, _coreHalo);
        Storyboard.SetTargetProperty(corePulse, "Opacity");
        story.Children.Add(corePulse);

        for (var i = 0; i < _impulses.Count; i++)
        {
            var pulse = _impulses[i];
            var originX = Canvas.GetLeft(pulse);

            var moveX = new DoubleAnimation
            {
                From = originX,
                To = originX + 160,
                Duration = TimeSpan.FromMilliseconds(2700 + (i * 240)),
                BeginTime = TimeSpan.FromMilliseconds(i * 160),
                RepeatBehavior = RepeatBehavior.Forever,
                EasingFunction = LifecycleVisuals.CreateEaseInOut(),
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(moveX, pulse);
            Storyboard.SetTargetProperty(moveX, "(Canvas.Left)");
            story.Children.Add(moveX);
        }

        for (var i = 0; i < _particles.Count; i++)
        {
            var particle = _particles[i];
            var driftX = new DoubleAnimation
            {
                From = -2,
                To = 2,
                Duration = TimeSpan.FromMilliseconds(5600 + (i * 420)),
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                EasingFunction = LifecycleVisuals.CreateEaseInOut(),
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(driftX, particle);
            Storyboard.SetTargetProperty(driftX, "(UIElement.RenderTransform).(TranslateTransform.X)");
            story.Children.Add(driftX);

            var driftY = new DoubleAnimation
            {
                From = -2,
                To = 2,
                Duration = TimeSpan.FromMilliseconds(6200 + (i * 460)),
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                EasingFunction = LifecycleVisuals.CreateEaseInOut(),
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(driftY, particle);
            Storyboard.SetTargetProperty(driftY, "(UIElement.RenderTransform).(TranslateTransform.Y)");
            story.Children.Add(driftY);
        }

        var shimmerMove = new DoubleAnimation
        {
            From = -80,
            To = 360,
            Duration = TimeSpan.FromMilliseconds(2500),
            RepeatBehavior = RepeatBehavior.Forever,
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };
        Storyboard.SetTarget(shimmerMove, _progressShimmer);
        Storyboard.SetTargetProperty(shimmerMove, "(UIElement.RenderTransform).(TranslateTransform.X)");
        story.Children.Add(shimmerMove);

        story.Begin();
    }

    private static Rectangle CreateMovingBand(
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
                    new GradientStop { Color = Color.FromArgb(0x00, 0x60, 0xBA, 0xF8), Offset = 0 },
                    new GradientStop { Color = Color.FromArgb(0xCF, 0x60, 0xBA, 0xF8), Offset = 0.5 },
                    new GradientStop { Color = Color.FromArgb(0x00, 0x60, 0xBA, 0xF8), Offset = 1 }
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

        var story = new Storyboard();
        Storyboard.SetTarget(move, band);
        Storyboard.SetTargetProperty(move, "(UIElement.RenderTransform).(CompositeTransform.TranslateX)");
        story.Children.Add(move);
        story.Begin();

        return band;
    }

    private void FadeInRoot()
    {
        var introFade = new DoubleAnimation
        {
            From = 0.0,
            To = 1.0,
            Duration = TimeSpan.FromMilliseconds(280),
            EasingFunction = LifecycleVisuals.CreateEaseOut(),
            EnableDependentAnimation = true
        };

        var introStoryboard = new Storyboard();
        Storyboard.SetTarget(introFade, _root);
        Storyboard.SetTargetProperty(introFade, "Opacity");
        introStoryboard.Children.Add(introFade);
        introStoryboard.Begin();
    }

    private async Task AnimateStatusTextAsync(string text)
    {
        var fadeOut = new DoubleAnimation
        {
            From = _statusText.Opacity,
            To = 0.22,
            Duration = TimeSpan.FromMilliseconds(170),
            EasingFunction = LifecycleVisuals.CreateEaseInOut(),
            EnableDependentAnimation = true
        };

        var fadeOutStory = new Storyboard();
        Storyboard.SetTarget(fadeOut, _statusText);
        Storyboard.SetTargetProperty(fadeOut, "Opacity");
        fadeOutStory.Children.Add(fadeOut);
        fadeOutStory.Begin();

        await Task.Delay(180);
        _statusText.Text = text;

        var fadeIn = new DoubleAnimation
        {
            From = 0.22,
            To = 0.94,
            Duration = TimeSpan.FromMilliseconds(220),
            EasingFunction = LifecycleVisuals.CreateEaseOut(),
            EnableDependentAnimation = true
        };

        var fadeInStory = new Storyboard();
        Storyboard.SetTarget(fadeIn, _statusText);
        Storyboard.SetTargetProperty(fadeIn, "Opacity");
        fadeInStory.Children.Add(fadeIn);
        fadeInStory.Begin();
    }

    private static async Task AnimateLayerOpacityAsync(IEnumerable<UIElement> elements, double toOpacity, int durationMs)
    {
        var storyboard = new Storyboard();

        foreach (var element in elements)
        {
            var fade = new DoubleAnimation
            {
                To = toOpacity,
                Duration = TimeSpan.FromMilliseconds(durationMs),
                EasingFunction = LifecycleVisuals.CreateEaseInOut(),
                EnableDependentAnimation = true
            };
            Storyboard.SetTarget(fade, element);
            Storyboard.SetTargetProperty(fade, "Opacity");
            storyboard.Children.Add(fade);
        }

        storyboard.Begin();
        await Task.Delay(durationMs + 20);
    }
}
