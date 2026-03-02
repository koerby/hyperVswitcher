using HyperTool.Models;
using HyperTool.WinUI.Helpers;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using Windows.Foundation;
using Windows.Graphics;
using Windows.UI;

namespace HyperTool.Guest.Views;

internal sealed class GuestTrayControlCenterWindow : Window
{
    private const int PanelCornerRadius = 18;
    public const int PopupWidth = 404;
    public const int PopupHeightWithUsb = 356;
    public const int PopupHeightCompact = 184;
    private readonly Grid _windowRoot = new();
    private readonly Border _panelRoot = new();
    private readonly Border _headerBorder = new();
    private readonly Border _usbCard = new();
    private readonly TextBlock _usbSelectedText = new();
    private readonly ComboBox _usbDeviceCombo = new();
    private readonly Button _refreshButton = new();
    private readonly Button _usbConnectButton = new();
    private readonly Button _usbDisconnectButton = new();
    private readonly Button _visibilityButton = new();
    private readonly Button _exitButton = new();
    private readonly Button _closeButton = new();

    private bool _isUpdatingUsbSelection;
    private bool _isTrayMenuEnabled = true;

    public event Action? CloseRequested;
    public event Action? RefreshUsbRequested;
    public event Action<string>? UsbSelected;
    public event Action? UsbConnectRequested;
    public event Action? UsbDisconnectRequested;
    public event Action? ToggleVisibilityRequested;
    public event Action? ExitRequested;

    public GuestTrayControlCenterWindow()
    {
        Title = "HyperTool Guest Control Center";
        ExtendsContentIntoTitleBar = false;

        BuildLayout();
        ConfigureWindowChrome();
        ApplyTheme(isDark: true);

        Activated += (_, args) =>
        {
            if (args.WindowActivationState == WindowActivationState.Deactivated)
            {
                CloseRequested?.Invoke();
            }
        };
    }

    public void ApplyTheme(bool isDark)
    {
        _panelRoot.RequestedTheme = isDark ? ElementTheme.Dark : ElementTheme.Light;

        var panelBackground = isDark ? Color.FromArgb(0xF6, 0x0F, 0x19, 0x2D) : Color.FromArgb(0xF8, 0xF7, 0xFB, 0xFF);
        var panelBorder = isDark ? Color.FromArgb(0xFF, 0x38, 0x55, 0x82) : Color.FromArgb(0xFF, 0xAD, 0xC3, 0xE6);
        var cardBackground = isDark ? Color.FromArgb(0xFF, 0x19, 0x25, 0x41) : Color.FromArgb(0xFF, 0xFC, 0xFD, 0xFF);
        var headerGradientStart = isDark ? Color.FromArgb(0xFF, 0x17, 0x24, 0x40) : Color.FromArgb(0xFF, 0xE8, 0xF1, 0xFF);
        var headerGradientEnd = isDark ? Color.FromArgb(0xFF, 0x10, 0x1B, 0x31) : Color.FromArgb(0xFF, 0xD8, 0xE7, 0xFC);
        var textPrimary = isDark ? Color.FromArgb(0xFF, 0xE8, 0xF1, 0xFF) : Color.FromArgb(0xFF, 0x1A, 0x2C, 0x48);
        var textSecondary = isDark ? Color.FromArgb(0xFF, 0xB5, 0xC8, 0xE6) : Color.FromArgb(0xFF, 0x4E, 0x67, 0x8C);
        var actionButtonBackground = isDark ? Color.FromArgb(0xFF, 0x2B, 0x3F, 0x66) : Color.FromArgb(0xFF, 0xE7, 0xF1, 0xFF);
        var actionButtonForeground = isDark ? Color.FromArgb(0xFF, 0xE9, 0xF2, 0xFF) : Color.FromArgb(0xFF, 0x1A, 0x2E, 0x4C);
        var actionButtonBorder = isDark ? Color.FromArgb(0xFF, 0x5D, 0x85, 0xBC) : Color.FromArgb(0xFF, 0x9A, 0xB9, 0xE3);

        _windowRoot.Background = new SolidColorBrush(Colors.Transparent);
        _panelRoot.Background = new SolidColorBrush(panelBackground);
        _panelRoot.BorderBrush = new SolidColorBrush(panelBorder);
        _headerBorder.Background = new LinearGradientBrush
        {
            StartPoint = new Point(0, 0),
            EndPoint = new Point(1, 1),
            GradientStops =
            {
                new GradientStop { Color = headerGradientStart, Offset = 0 },
                new GradientStop { Color = headerGradientEnd, Offset = 1 }
            }
        };

        _usbSelectedText.Foreground = new SolidColorBrush(textSecondary);
        _usbCard.Background = new SolidColorBrush(cardBackground);
        _usbCard.BorderBrush = new SolidColorBrush(actionButtonBorder);
        _usbCard.BorderThickness = new Thickness(1);

        _usbDeviceCombo.Background = new SolidColorBrush(cardBackground);
        _usbDeviceCombo.Foreground = new SolidColorBrush(textPrimary);
        _usbDeviceCombo.BorderBrush = new SolidColorBrush(actionButtonBorder);

        ApplyButtonColors(_refreshButton, actionButtonBackground, actionButtonForeground, actionButtonBorder);
        ApplyButtonColors(_usbConnectButton, actionButtonBackground, actionButtonForeground, actionButtonBorder);
        ApplyButtonColors(_usbDisconnectButton, actionButtonBackground, actionButtonForeground, actionButtonBorder);
        ApplyButtonColors(_visibilityButton, actionButtonBackground, actionButtonForeground, actionButtonBorder);
        ApplyButtonColors(_exitButton, actionButtonBackground, actionButtonForeground, actionButtonBorder);
        ApplyButtonColors(_closeButton, actionButtonBackground, actionButtonForeground, actionButtonBorder);
    }

    public void UpdateView(IReadOnlyList<UsbIpDeviceInfo> devices, string? selectedBusId, bool isMainWindowVisible, bool isTrayMenuEnabled)
    {
        _isTrayMenuEnabled = isTrayMenuEnabled;
        _usbCard.Visibility = isTrayMenuEnabled ? Visibility.Visible : Visibility.Collapsed;

        var selectedDisplay = devices.FirstOrDefault(item => string.Equals(item.BusId, selectedBusId, StringComparison.OrdinalIgnoreCase))?.DisplayName
            ?? "Kein USB-Gerät ausgewählt";
        _usbSelectedText.Text = selectedDisplay;

        _isUpdatingUsbSelection = true;
        try
        {
            _usbDeviceCombo.Items.Clear();
            foreach (var usb in devices)
            {
                _usbDeviceCombo.Items.Add(new ComboBoxItem
                {
                    Content = usb.DisplayName,
                    Tag = usb.BusId
                });
            }

            var selectedIndex = devices.ToList().FindIndex(item => string.Equals(item.BusId, selectedBusId, StringComparison.OrdinalIgnoreCase));
            _usbDeviceCombo.SelectedIndex = selectedIndex;
            _usbDeviceCombo.IsEnabled = devices.Count > 0;
        }
        finally
        {
            _isUpdatingUsbSelection = false;
        }

        var hasSelection = !string.IsNullOrWhiteSpace(selectedBusId);
        _refreshButton.IsEnabled = isTrayMenuEnabled;
        _usbConnectButton.IsEnabled = isTrayMenuEnabled && hasSelection;
        _usbDisconnectButton.IsEnabled = isTrayMenuEnabled && hasSelection;
        _visibilityButton.Content = isMainWindowVisible ? "⌂  Ausblenden" : "⌂  Einblenden";

        SetPanelSize(PopupWidth, isTrayMenuEnabled ? PopupHeightWithUsb : PopupHeightCompact);
    }

    private void BuildLayout()
    {
        _windowRoot.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        _panelRoot.Margin = new Thickness(0);
        _panelRoot.CornerRadius = new CornerRadius(PanelCornerRadius);
        _panelRoot.BorderThickness = new Thickness(1);
        _panelRoot.Padding = new Thickness(12);
        _panelRoot.Shadow = new ThemeShadow();

        var stack = new StackPanel { Spacing = 10 };

        _headerBorder.CornerRadius = new CornerRadius(14);
        _headerBorder.Padding = new Thickness(12, 10, 10, 10);

        var headerGrid = new Grid();
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        headerGrid.Children.Add(new Image
        {
            Width = 24,
            Height = 24,
            Source = new BitmapImage(new Uri("ms-appx:///Assets/HyperTool.Guest.Icon.Transparent.png")),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 10, 0)
        });

        var title = new TextBlock
        {
            Text = "HyperTool Guest",
            FontSize = 18,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            VerticalAlignment = VerticalAlignment.Center
        };
        Grid.SetColumn(title, 1);
        headerGrid.Children.Add(title);

        _closeButton.Content = "✕";
        _closeButton.Width = 32;
        _closeButton.Height = 32;
        _closeButton.Style = CreateRoundedButtonStyle();
        _closeButton.Click += (_, _) => CloseRequested?.Invoke();
        Grid.SetColumn(_closeButton, 2);
        headerGrid.Children.Add(_closeButton);

        _headerBorder.Child = headerGrid;
        stack.Children.Add(_headerBorder);

        _usbCard.CornerRadius = new CornerRadius(14);
        _usbCard.Padding = new Thickness(12);
        _usbCard.Background = new SolidColorBrush(Color.FromArgb(0xFF, 0x19, 0x25, 0x41));

        var usbStack = new StackPanel { Spacing = 8 };
        usbStack.Children.Add(new TextBlock { Text = "USB", FontSize = 12, Opacity = 0.8 });

        _usbSelectedText.Text = "Selected: -";
        _usbSelectedText.FontSize = 13;
        _usbSelectedText.TextWrapping = TextWrapping.WrapWholeWords;
        usbStack.Children.Add(_usbSelectedText);

        _usbDeviceCombo.MinHeight = 34;
        _usbDeviceCombo.Width = 340;
        _usbDeviceCombo.CornerRadius = new CornerRadius(8);
        _usbDeviceCombo.HorizontalAlignment = HorizontalAlignment.Stretch;
        _usbDeviceCombo.PlaceholderText = "USB-Gerät auswählen";
        _usbDeviceCombo.SelectionChanged += (_, _) =>
        {
            if (_isUpdatingUsbSelection)
            {
                return;
            }

            if (_usbDeviceCombo.SelectedItem is ComboBoxItem selectedItem && selectedItem.Tag is string busId)
            {
                UsbSelected?.Invoke(busId);
            }
        };
        usbStack.Children.Add(_usbDeviceCombo);

        var usbActions = new Grid { ColumnSpacing = 8 };
        usbActions.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        usbActions.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        usbActions.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        _refreshButton.Content = "⟳  Refresh";
        _refreshButton.Style = CreateActionButtonStyle();
        _refreshButton.HorizontalAlignment = HorizontalAlignment.Stretch;
        _refreshButton.HorizontalContentAlignment = HorizontalAlignment.Center;
        _refreshButton.Padding = new Thickness(8, 0, 8, 0);
        _refreshButton.Click += (_, _) => RefreshUsbRequested?.Invoke();
        Grid.SetColumn(_refreshButton, 0);
        usbActions.Children.Add(_refreshButton);

        _usbConnectButton.Content = "🔌  Connect";
        _usbConnectButton.Style = CreateActionButtonStyle();
        _usbConnectButton.HorizontalAlignment = HorizontalAlignment.Stretch;
        _usbConnectButton.HorizontalContentAlignment = HorizontalAlignment.Center;
        _usbConnectButton.Padding = new Thickness(8, 0, 8, 0);
        _usbConnectButton.Click += (_, _) => UsbConnectRequested?.Invoke();
        Grid.SetColumn(_usbConnectButton, 1);
        usbActions.Children.Add(_usbConnectButton);

        _usbDisconnectButton.Content = "⏏  Disconnect";
        _usbDisconnectButton.Style = CreateActionButtonStyle();
        _usbDisconnectButton.HorizontalAlignment = HorizontalAlignment.Stretch;
        _usbDisconnectButton.HorizontalContentAlignment = HorizontalAlignment.Center;
        _usbDisconnectButton.Padding = new Thickness(8, 0, 8, 0);
        _usbDisconnectButton.Click += (_, _) => UsbDisconnectRequested?.Invoke();
        Grid.SetColumn(_usbDisconnectButton, 2);
        usbActions.Children.Add(_usbDisconnectButton);

        usbStack.Children.Add(usbActions);
        _usbCard.Child = usbStack;
        stack.Children.Add(_usbCard);

        _visibilityButton.Content = "⌂  Einblenden";
        _visibilityButton.Style = CreateActionButtonStyle();
        _visibilityButton.HorizontalAlignment = HorizontalAlignment.Stretch;
        _visibilityButton.Click += (_, _) => ToggleVisibilityRequested?.Invoke();
        stack.Children.Add(_visibilityButton);

        _exitButton.Content = "⏻  Beenden";
        _exitButton.Style = CreateActionButtonStyle();
        _exitButton.HorizontalAlignment = HorizontalAlignment.Stretch;
        _exitButton.Click += (_, _) => ExitRequested?.Invoke();
        stack.Children.Add(_exitButton);

        _panelRoot.Child = stack;
        _windowRoot.Children.Add(_panelRoot);
        Content = _windowRoot;
    }

    private void ConfigureWindowChrome()
    {
        DwmWindowHelper.ApplyRoundedCorners(this);

        var presenter = AppWindow.Presenter as OverlappedPresenter;
        if (presenter is not null)
        {
            presenter.SetBorderAndTitleBar(false, false);
            presenter.IsResizable = false;
            presenter.IsMaximizable = false;
            presenter.IsMinimizable = false;
            presenter.IsAlwaysOnTop = true;
        }

        AppWindow.IsShownInSwitchers = false;
        SetPanelSize(PopupWidth, PopupHeightWithUsb);
    }

    private void SetPanelSize(int width, int height)
    {
        AppWindow.Resize(new SizeInt32(width, height));
        DwmWindowHelper.ApplyRoundedRegion(this, width, height, PanelCornerRadius);
    }

    private static Style CreateActionButtonStyle()
    {
        var style = new Style { TargetType = typeof(Button) };
        style.Setters.Add(new Setter(Control.CornerRadiusProperty, new CornerRadius(12)));
        style.Setters.Add(new Setter(FrameworkElement.HeightProperty, 38d));
        style.Setters.Add(new Setter(Control.HorizontalContentAlignmentProperty, HorizontalAlignment.Left));
        style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(12, 0, 12, 0)));
        style.Setters.Add(new Setter(Control.FontSizeProperty, 13d));
        style.Setters.Add(new Setter(Control.FontWeightProperty, Microsoft.UI.Text.FontWeights.SemiBold));
        style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));
        style.Setters.Add(new Setter(UIElement.OpacityProperty, 1.0));
        return style;
    }

    private static Style CreateRoundedButtonStyle()
    {
        var style = new Style { TargetType = typeof(Button) };
        style.Setters.Add(new Setter(Control.CornerRadiusProperty, new CornerRadius(10)));
        style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(10, 0, 10, 0)));
        style.Setters.Add(new Setter(Control.FontSizeProperty, 12d));
        style.Setters.Add(new Setter(Control.FontWeightProperty, Microsoft.UI.Text.FontWeights.SemiBold));
        style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));
        return style;
    }

    private static void ApplyButtonColors(Button button, Color background, Color foreground, Color border)
    {
        button.Background = new SolidColorBrush(background);
        button.Foreground = new SolidColorBrush(foreground);
        button.BorderBrush = new SolidColorBrush(border);
    }
}
