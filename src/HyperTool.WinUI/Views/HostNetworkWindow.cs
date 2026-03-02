using HyperTool.Models;
using HyperTool.WinUI.Helpers;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Windowing;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Windows.Graphics;
using Windows.UI;

namespace HyperTool.WinUI.Views;

public sealed class HostNetworkWindow : Window
{
    private readonly IReadOnlyList<HostNetworkAdapterInfo> _adapters;
  private readonly bool _isDarkMode;

  public HostNetworkWindow(IReadOnlyList<HostNetworkAdapterInfo> adapters, string uiTheme)
    {
        _adapters = adapters
            .Where(adapter => adapter is not null)
            .ToList();
        _isDarkMode = string.Equals(uiTheme, "Dark", StringComparison.OrdinalIgnoreCase);

        Title = "Host Network";
        DwmWindowHelper.ApplyRoundedCorners(this);
        AppWindow.Resize(new SizeInt32(1060, 640));
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
            Margin = new Thickness(16),
            RowSpacing = 10,
            Background = Application.Current.Resources["PageBackgroundBrush"] as Brush
        };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var headCard = CreateCard(14);
        var head = new StackPanel { Spacing = 2 };
        head.Children.Add(new TextBlock { Text = "Host Network", FontSize = 22, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
        head.Children.Add(new TextBlock { Text = "Übersicht aller Host-Adapter inklusive IP, Gateway und DNS.", Opacity = 0.82 });
        head.Children.Add(new TextBlock { Text = $"{_adapters.Count} Adapter gefunden", Opacity = 0.82 });
        headCard.Child = head;
        root.Children.Add(headCard);

        var listStack = new StackPanel { Spacing = 10 };
        foreach (var adapter in _adapters)
        {
            listStack.Children.Add(CreateAdapterCard(adapter));
        }

        var bodyCard = CreateCard(12);
        bodyCard.Child = new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            Content = listStack
        };
        Grid.SetRow(bodyCard, 1);
        root.Children.Add(bodyCard);

        var close = new Button
        {
            Content = "Schließen",
            HorizontalAlignment = HorizontalAlignment.Right,
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush
        };
        close.Click += (_, _) => Close();
        Grid.SetRow(close, 2);
        root.Children.Add(close);

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

    private static Border CreateAdapterCard(HostNetworkAdapterInfo adapter)
    {
        var card = new Border
        {
            Padding = new Thickness(12, 10, 12, 10),
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
            Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush
        };

        var cardStack = new StackPanel { Spacing = 8 };

        var topRow = new Grid { ColumnSpacing = 10 };
        topRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        topRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var adapterTitle = new TextBlock
        {
            Text = string.IsNullOrWhiteSpace(adapter.AdapterName) ? "Unbekannter Adapter" : adapter.AdapterName,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            TextTrimming = TextTrimming.CharacterEllipsis
        };
        topRow.Children.Add(adapterTitle);

        var badgeRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            HorizontalAlignment = HorizontalAlignment.Right
        };

      if (adapter.IsDefaultSwitch)
      {
        badgeRow.Children.Add(CreateBadge("Default Switch"));
      }

      if (adapter.HasGateway)
      {
        badgeRow.Children.Add(CreateBadge("Gateway"));
      }

      if (badgeRow.Children.Count > 0)
      {
        Grid.SetColumn(badgeRow, 1);
        topRow.Children.Add(badgeRow);
      }

        cardStack.Children.Add(topRow);

        var grid = new Grid { ColumnSpacing = 10, RowSpacing = 6 };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(190) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(190) });

        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var descriptionCell = CreateKeyValue("Beschreibung", adapter.InterfaceDescription);
        grid.Children.Add(descriptionCell);
        Grid.SetColumn(descriptionCell, 0);
        Grid.SetRow(descriptionCell, 0);

        var ipCell = CreateKeyValue("IP", adapter.IpAddresses);
        grid.Children.Add(ipCell);
        Grid.SetColumn(ipCell, 1);
        Grid.SetRow(ipCell, 0);

        var dnsCell = CreateKeyValue("DNS", adapter.DnsServers);
        grid.Children.Add(dnsCell);
        Grid.SetColumn(dnsCell, 2);
        Grid.SetRow(dnsCell, 0);

        var subnetCell = CreateKeyValue("Subnetz", adapter.Subnets);
        grid.Children.Add(subnetCell);
        Grid.SetColumn(subnetCell, 0);
        Grid.SetRow(subnetCell, 1);

        var gatewayCell = CreateKeyValue("Gateway", adapter.Gateway);
        grid.Children.Add(gatewayCell);
        Grid.SetColumn(gatewayCell, 1);
        Grid.SetRow(gatewayCell, 1);

        cardStack.Children.Add(grid);
        card.Child = cardStack;
        return card;
    }

    private static Border CreateBadge(string text)
    {
        return new Border
        {
            Padding = new Thickness(10, 2, 10, 2),
            CornerRadius = new CornerRadius(999),
            BorderThickness = new Thickness(1),
            BorderBrush = Application.Current.Resources["AccentBrush"] as Brush,
            Background = Application.Current.Resources["AccentSoftBrush"] as Brush,
            Child = new TextBlock
            {
                Text = text,
                FontSize = 11,
                FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                Foreground = Application.Current.Resources["AccentTextBrush"] as Brush
            }
        };
    }

    private static FrameworkElement CreateKeyValue(string key, string? value)
    {
      var stack = new StackPanel { Spacing = 2 };
      stack.Children.Add(new TextBlock { Text = key, FontSize = 12, Opacity = 0.78 });
      stack.Children.Add(new TextBlock
      {
        Text = string.IsNullOrWhiteSpace(value) ? "-" : value,
        FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
        TextWrapping = TextWrapping.Wrap
      });
      return stack;
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
