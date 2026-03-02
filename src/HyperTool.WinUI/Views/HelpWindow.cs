using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Windowing;
using HyperTool.WinUI.Helpers;
using System.Diagnostics;
using System.IO;
using Windows.Graphics;
using Windows.UI;

namespace HyperTool.WinUI.Views;

public sealed class HelpWindow : Window
{
    private readonly string _configPath;
    private readonly string _repoUrl;
    private readonly bool _isDarkMode;

    public HelpWindow(string configPath, string repoUrl, string uiTheme)
    {
        _configPath = configPath;
        _repoUrl = repoUrl;
        _isDarkMode = string.Equals(uiTheme, "Dark", StringComparison.OrdinalIgnoreCase);

        Title = "HyperTool Hilfe";
        DwmWindowHelper.ApplyRoundedCorners(this);
        AppWindow.Resize(new SizeInt32(860, 720));
        TryApplyWindowIcon();

        Content = BuildLayout();
        ApplyRequestedTheme();
        UpdateTitleBarAppearance();
    }

    private UIElement BuildLayout()
    {
        var host = new Grid
        {
            Background = Application.Current.Resources["PageBackgroundBrush"] as Brush
        };

        var root = new Grid
        {
            Margin = new Thickness(18),
            Background = Application.Current.Resources["PageBackgroundBrush"] as Brush
        };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(10) });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(10) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var headerCard = CreateCard(16);
        var headerRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10, VerticalAlignment = VerticalAlignment.Center };
        headerRow.Children.Add(new Image
        {
            Source = new Microsoft.UI.Xaml.Media.Imaging.BitmapImage(new Uri("ms-appx:///Assets/HyperTool.ico")),
            Width = 24,
            Height = 24
        });
        var titleStack = new StackPanel { Spacing = 2 };
        titleStack.Children.Add(new TextBlock { Text = "HyperTool Hilfe", FontSize = 24, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
        titleStack.Children.Add(new TextBlock { Text = "Kurzübersicht über Funktionen und Einstellungen", Opacity = 0.8 });
        headerRow.Children.Add(titleStack);
        headerCard.Child = headerRow;
        root.Children.Add(headerCard);

        var bodyCard = CreateCard(12);
        var bodyStack = new StackPanel { Spacing = 12 };
        bodyStack.Children.Add(CreateSection("VM Auswahl", "Oben im Header werden alle Hyper-V VMs angezeigt. Klick auf einen Chip wählt die aktive Arbeits-VM."));
        bodyStack.Children.Add(CreateSection("One Click Aktionen", "Im VM-Bereich kannst du VM starten, stoppen, hart ausschalten, neu starten und die VM-Konsole öffnen."));
        bodyStack.Children.Add(CreateSection("Netzwerk", "Pro Netzwerkkarte kannst du den Switch direkt per One-Click umstellen. Host-Netzwerk öffnet die detaillierte Adapter-Ansicht."));
        bodyStack.Children.Add(CreateSection("Snapshots", "Checkpoints erstellen, anwenden und löschen."));
        bodyStack.Children.Add(CreateSection("Config / Info", "Einstellungen speichern, Updates prüfen und Konfiguration verwalten."));
        bodyCard.Child = new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            Content = bodyStack
        };

        Grid.SetRow(bodyCard, 2);
        root.Children.Add(bodyCard);

        var actions = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Spacing = 10
        };

        actions.Children.Add(CreateActionButton("📄", "Logs öffnen", (_, _) => OpenLogs()));
        actions.Children.Add(CreateActionButton("⚙", "Config öffnen", (_, _) => OpenConfig()));
        actions.Children.Add(CreateActionButton("🌐", "GitHub Repo", (_, _) => OpenRepo()));
        actions.Children.Add(CreateActionButton("✓", "OK", (_, _) => Close()));

        Grid.SetRow(actions, 4);
        root.Children.Add(actions);

        host.Children.Add(root);
        return host;
    }

    private static Border CreateCard(double padding)
    {
        return new Border
        {
            Padding = new Thickness(padding),
            CornerRadius = new CornerRadius(12),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["PanelBackgroundBrush"] as Brush
        };
    }

    private static Border CreateSection(string title, string text)
    {
        var section = new Border
        {
            Padding = new Thickness(12, 10, 12, 10),
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush
        };

        var stack = new StackPanel { Spacing = 4 };
        stack.Children.Add(new TextBlock { Text = title, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
        stack.Children.Add(new TextBlock { Text = text, TextWrapping = TextWrapping.Wrap, Opacity = 0.86 });
        section.Child = stack;
        return section;
    }

    private static Button CreateActionButton(string icon, string text, RoutedEventHandler onClick)
    {
        var iconHost = new Grid
        {
            Width = 20,
            Height = 20,
            VerticalAlignment = VerticalAlignment.Center
        };
        iconHost.Children.Add(new Viewbox
        {
            Stretch = Stretch.Uniform,
            Child = new TextBlock
            {
                Text = icon,
                FontSize = 17,
                FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                TextAlignment = TextAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            }
        });

        var button = new Button
        {
            Padding = new Thickness(10, 6, 10, 6),
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush,
            Content = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 6,
                Children =
                {
                    iconHost,
                    new TextBlock { Text = text, VerticalAlignment = VerticalAlignment.Center, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold }
                }
            }
        };

        button.Click += onClick;
        return button;
    }

    private static void OpenLogs()
    {
        var logsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "HyperTool",
            "logs");

        Directory.CreateDirectory(logsPath);
        Process.Start(new ProcessStartInfo
        {
            FileName = logsPath,
            UseShellExecute = true
        });
    }

    private void OpenConfig()
    {
        var directoryPath = Path.GetDirectoryName(_configPath);
        if (!string.IsNullOrWhiteSpace(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
        }

        if (File.Exists(_configPath))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"/select,\"{_configPath}\"",
                UseShellExecute = true
            });
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = directoryPath ?? Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            UseShellExecute = true
        });
    }

    private void OpenRepo()
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = _repoUrl,
            UseShellExecute = true
        });
    }

    private void ApplyRequestedTheme()
    {
        if (Content is FrameworkElement root)
        {
            root.RequestedTheme = _isDarkMode ? ElementTheme.Dark : ElementTheme.Light;
        }
    }

    private void UpdateTitleBarAppearance()
    {
        try
        {
            if (AppWindow?.TitleBar is not AppWindowTitleBar titleBar)
            {
                return;
            }

            if (_isDarkMode)
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
                Path.Combine(AppContext.BaseDirectory, "Assets", "HyperTool.ico"),
                Path.Combine(AppContext.BaseDirectory, "HyperTool.ico")
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
