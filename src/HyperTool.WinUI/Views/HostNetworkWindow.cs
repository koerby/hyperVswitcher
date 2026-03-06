using HyperTool.Models;
using HyperTool.WinUI.Helpers;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Windowing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Windows.Graphics;
using Windows.UI;

namespace HyperTool.WinUI.Views;

public sealed class HostNetworkWindow : Window
{
  private enum AdapterChipKind
  {
    DefaultSwitch,
    Gateway,
    NetworkProfileTrusted,
    NetworkProfilePublic
  }

    private IReadOnlyList<HostNetworkAdapterInfo> _adapters;
  private readonly Func<string?, string, Task<bool>> _setProfileCategoryAsync;
  private readonly Func<Task<IReadOnlyList<HostNetworkAdapterInfo>>> _reloadAdaptersAsync;
  private readonly List<Control> _profileEditControls = [];
  private StackPanel? _adapterListStack;
  private TextBlock? _adapterCountText;
  private bool _profileUpdateInProgress;
  private readonly bool _isDarkMode;

  public HostNetworkWindow(
      IReadOnlyList<HostNetworkAdapterInfo> adapters,
      string uiTheme,
      Func<string?, string, Task<bool>>? setProfileCategoryAsync = null,
      Func<Task<IReadOnlyList<HostNetworkAdapterInfo>>>? reloadAdaptersAsync = null)
    {
        _adapters = NormalizeAdapters(adapters);
        _setProfileCategoryAsync = setProfileCategoryAsync ?? ((_, _) => Task.FromResult(false));
        _reloadAdaptersAsync = reloadAdaptersAsync ?? (() => Task.FromResult(_adapters));
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
        _adapterCountText = new TextBlock { Text = $"{_adapters.Count} Adapter gefunden", Opacity = 0.82 };
        head.Children.Add(_adapterCountText);
        headCard.Child = head;
        root.Children.Add(headCard);

        _adapterListStack = new StackPanel { Spacing = 10 };
        RebuildAdapterCards();

        var bodyCard = CreateCard(12);
        bodyCard.Child = new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
          Content = _adapterListStack
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

    private Border CreateAdapterCard(HostNetworkAdapterInfo adapter)
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
          HorizontalAlignment = HorizontalAlignment.Right,
          VerticalAlignment = VerticalAlignment.Center
        };

        if (adapter.IsDefaultSwitch)
        {
          badgeRow.Children.Add(CreateHighlightBadge("Default Switch", Symbol.Switch, AdapterChipKind.DefaultSwitch));
        }

        if (adapter.HasGateway)
        {
          badgeRow.Children.Add(CreateHighlightBadge("Gateway", Symbol.World, AdapterChipKind.Gateway));
        }

        var networkCategory = NormalizeNetworkProfileCategory(adapter.NetworkProfileCategory);
        var profileBadge = CreateNetworkProfileBadge(networkCategory);
        if (profileBadge is not null)
        {
          badgeRow.Children.Add(profileBadge);
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
        cardStack.Children.Add(CreateNetworkProfileEditor(adapter));
        card.Child = cardStack;
        return card;
    }

    private FrameworkElement CreateNetworkProfileEditor(HostNetworkAdapterInfo adapter)
    {
      var row = new StackPanel
      {
        Orientation = Orientation.Horizontal,
        Spacing = 6,
        HorizontalAlignment = HorizontalAlignment.Right,
        VerticalAlignment = VerticalAlignment.Center
      };

      var category = NormalizeNetworkProfileCategory(adapter.NetworkProfileCategory);
      var actionButton = new Button
      {
        Content = category switch
        {
          "Public" => "Auf Privat umstellen",
          "Private" => "Auf Öffentlich umstellen",
          "DomainAuthenticated" => "Domäne (gesperrt)",
          _ => "Netzprofil setzen"
        },
        CornerRadius = new CornerRadius(8),
        BorderThickness = new Thickness(1),
        BorderBrush = Application.Current.Resources["PanelBorderBrush"] as Brush,
        Background = Application.Current.Resources["SurfaceSoftBrush"] as Brush,
        MinWidth = 160,
        HorizontalAlignment = HorizontalAlignment.Right,
        IsEnabled = category is "Public" or "Private"
      };

      actionButton.Click += async (_, _) =>
      {
        if (_profileUpdateInProgress)
        {
          return;
        }

        var targetCategory = category == "Public" ? "Private" : "Public";
        var changed = await ApplyNetworkProfileChangeAsync(adapter.AdapterName, targetCategory);
        if (!changed)
        {
          actionButton.IsEnabled = true;
        }
      };

      _profileEditControls.Add(actionButton);
      row.Children.Add(actionButton);
      return row;
    }

    private async Task<bool> ApplyNetworkProfileChangeAsync(string? adapterName, string category)
    {
      if (_profileUpdateInProgress)
      {
        return false;
      }

      var changed = false;
      _profileUpdateInProgress = true;
      SetProfileEditorsEnabled(false);
      try
      {
        changed = await _setProfileCategoryAsync(adapterName, category);
        if (changed)
        {
          await ReloadAsync();
        }
      }
      finally
      {
        _profileUpdateInProgress = false;
        SetProfileEditorsEnabled(true);
      }

      return changed;
    }

    private void RebuildAdapterCards()
    {
      if (_adapterListStack is null)
      {
        return;
      }

      _profileEditControls.Clear();
      _adapterListStack.Children.Clear();
      foreach (var adapter in _adapters)
      {
        _adapterListStack.Children.Add(CreateAdapterCard(adapter));
      }

      SetProfileEditorsEnabled(!_profileUpdateInProgress);
      if (_adapterCountText is not null)
      {
        _adapterCountText.Text = $"{_adapters.Count} Adapter gefunden";
      }
    }

    private void SetProfileEditorsEnabled(bool enabled)
    {
      foreach (var control in _profileEditControls)
      {
        control.IsEnabled = enabled;
      }
    }

    private static IReadOnlyList<HostNetworkAdapterInfo> NormalizeAdapters(IReadOnlyList<HostNetworkAdapterInfo> adapters)
      => adapters
          .Where(adapter => adapter is not null)
          .ToList();

    public async Task ReloadAsync()
    {
      var reloaded = await _reloadAdaptersAsync();
      _adapters = NormalizeAdapters(reloaded);
      RebuildAdapterCards();
    }

    private Border CreateHighlightBadge(string label, Symbol icon, AdapterChipKind kind)
    {
      var (chipBackground, chipBorder, iconBackground, iconBorder, iconForeground, textForeground) = ResolveChipPalette(kind);

        var iconContainer = new Border
        {
            Width = 18,
            Height = 18,
            CornerRadius = new CornerRadius(5),
            Background = iconBackground,
            BorderThickness = new Thickness(1),
            BorderBrush = iconBorder,
            VerticalAlignment = VerticalAlignment.Center,
            Child = new SymbolIcon
            {
                Symbol = icon,
                Foreground = iconForeground,
                Width = 12,
              Height = 12,
              HorizontalAlignment = HorizontalAlignment.Center,
              VerticalAlignment = VerticalAlignment.Center
            }
        };

        var content = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            VerticalAlignment = VerticalAlignment.Center
        };
        content.Children.Add(iconContainer);
        content.Children.Add(new TextBlock
        {
            Text = label,
            FontSize = 12,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            Foreground = textForeground,
          VerticalAlignment = VerticalAlignment.Center,
          TextWrapping = TextWrapping.NoWrap
        });

        return new Border
        {
            MinHeight = 30,
            Padding = new Thickness(10, 5, 10, 5),
            CornerRadius = new CornerRadius(9),
            BorderThickness = new Thickness(1),
            BorderBrush = chipBorder,
            Background = chipBackground,
            Child = content
        };
    }

      private Border? CreateNetworkProfileBadge(string category)
      {
        if (string.Equals(category, "Unknown", StringComparison.OrdinalIgnoreCase))
        {
          return null;
        }

        var chipKind = ResolveNetworkProfileChipKind(category);
        var label = category switch
        {
          "Public" => "Öffentlich",
          "DomainAuthenticated" => "Domäne",
          "Private" => "Privat",
          _ => "Unbekannt"
        };

        var symbol = category switch
        {
          "Public" => Symbol.Important,
          "DomainAuthenticated" => Symbol.Contact,
          "Private" => Symbol.ProtectedDocument,
          _ => Symbol.Help
        };

        var (chipBackground, chipBorder, iconBackground, iconBorder, iconForeground, textForeground) = ResolveChipPalette(chipKind);

        var iconContainer = new Border
        {
          Width = 18,
          Height = 18,
          CornerRadius = new CornerRadius(5),
          Background = iconBackground,
          BorderThickness = new Thickness(1),
          BorderBrush = iconBorder,
          VerticalAlignment = VerticalAlignment.Center,
          Child = new SymbolIcon
          {
            Symbol = symbol,
            Foreground = iconForeground,
            Width = 12,
            Height = 12,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
          }
        };

        var content = new StackPanel
        {
          Orientation = Orientation.Horizontal,
          Spacing = 6,
          VerticalAlignment = VerticalAlignment.Center
        };
        content.Children.Add(iconContainer);
        content.Children.Add(new TextBlock
        {
          Text = label,
          FontSize = 12,
          FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
          Foreground = textForeground,
          VerticalAlignment = VerticalAlignment.Center,
          TextWrapping = TextWrapping.NoWrap
        });

        return new Border
        {
          MinHeight = 30,
          Padding = new Thickness(9, 5, 9, 5),
          CornerRadius = new CornerRadius(9),
          BorderThickness = new Thickness(1),
          BorderBrush = chipBorder,
          Background = chipBackground,
          Child = content
        };
      }

      private static string NormalizeNetworkProfileCategory(string? category)
      {
        if (string.IsNullOrWhiteSpace(category))
        {
          return "Unknown";
        }

        return category.Trim() switch
        {
          "Public" => "Public",
          "Private" => "Private",
          "DomainAuthenticated" => "DomainAuthenticated",
          _ => "Unknown"
        };
      }

      private static AdapterChipKind ResolveNetworkProfileChipKind(string category)
        => category switch
        {
          "Public" => AdapterChipKind.NetworkProfilePublic,
          "Private" => AdapterChipKind.NetworkProfileTrusted,
          "DomainAuthenticated" => AdapterChipKind.NetworkProfileTrusted,
          _ => AdapterChipKind.NetworkProfilePublic
        };

      private (Brush chipBackground, Brush chipBorder, Brush iconBackground, Brush iconBorder, Brush iconForeground, Brush textForeground) ResolveChipPalette(AdapterChipKind kind)
        {
          static SolidColorBrush Brush(byte a, byte r, byte g, byte b) => new(Color.FromArgb(a, r, g, b));

          if (kind == AdapterChipKind.DefaultSwitch)
          {
            if (_isDarkMode)
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

          if (kind == AdapterChipKind.NetworkProfilePublic)
          {
            if (_isDarkMode)
            {
              return (
                Brush(0xFF, 0x4A, 0x1E, 0x2A),
                Brush(0xFF, 0xE8, 0x4A, 0x5F),
                Brush(0xFF, 0xE8, 0x4A, 0x5F),
                Brush(0xFF, 0xE8, 0x4A, 0x5F),
                Brush(0xFF, 0x2C, 0x11, 0x18),
                Brush(0xFF, 0xFF, 0xDC, 0xE5));
            }

            return (
              Brush(0xFF, 0xFF, 0xE8, 0xED),
              Brush(0xFF, 0xD6, 0x3A, 0x53),
              Brush(0xFF, 0xD6, 0x3A, 0x53),
              Brush(0xFF, 0xD6, 0x3A, 0x53),
              Brush(0xFF, 0xFF, 0xFB, 0xFC),
              Brush(0xFF, 0x76, 0x1D, 0x33));
          }

          if (_isDarkMode)
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
