using HyperTool.Models;
using Microsoft.Win32;
using Serilog;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace HyperTool.Services;

public sealed class TrayService : ITrayService
{
    private NotifyIcon? _notifyIcon;
    private ContextMenuStrip? _contextMenu;

    private Action? _showAction;
    private Action? _hideAction;
    private Action? _toggleControlCenterAction;
    private Action? _toggleControlCenterCompactAction;
    private Action? _hideControlCenterAction;
    private Func<string>? _getUiTheme;
    private Func<bool>? _isWindowVisible;
    private Func<bool>? _isTrayMenuEnabled;
    private Func<IReadOnlyList<VmDefinition>>? _getVms;
    private Func<IReadOnlyList<HyperVSwitchInfo>>? _getSwitches;
    private Func<Task>? _refreshTrayDataAction;
    private Func<string, Task>? _startVmAction;
    private Func<string, Task>? _stopVmAction;
    private Func<string, Task>? _restartVmAction;
    private Func<string, Task>? _openConsoleAction;
    private Func<string, Task>? _createSnapshotAction;
    private Func<string, string, Task>? _connectVmToSwitchAction;
    private Func<string, Task>? _disconnectVmSwitchAction;
    private Action? _exitAction;
    private EventHandler? _trayStateChangedHandler;
    private Action<EventHandler>? _unsubscribeTrayStateChanged;
    private bool _isContextMenuOpen;
    private bool _hasPendingMenuRefresh;
    private TrayMenuPalette _menuPalette = TrayMenuPalette.CreateLight();
    private ToolStripProfessionalRenderer? _menuRenderer;
    private Form? _quickPanel;
    private readonly List<VmSelectionItem> _panelVms = [];
    private readonly List<string> _panelSwitches = [];
    private int _selectedVmIndex = -1;
    private int _selectedSwitchIndex = -1;
    private Label? _vmNameLabel;
    private Label? _vmSelectorLabel;
    private Label? _vmStatusLabel;
    private Panel? _switchButtonsPanel;
    private Label? _switchStatusLabel;
    private Button? _vmPrevButton;
    private Button? _vmNextButton;
    private Button? _switchPrevButton;
    private Button? _switchNextButton;
    private Button? _switchConnectButton;
    private Button? _switchDisconnectButton;
    private Button? _toggleVisibilityButton;
    private Button? _startVmButton;
    private Button? _shutdownVmButton;
    private Button? _snapshotVmButton;
    private ToolTip? _quickPanelToolTip;

    public void Initialize(
        Action showAction,
        Action hideAction,
        Action toggleControlCenterAction,
        Action toggleControlCenterCompactAction,
        Action hideControlCenterAction,
        Func<string> getUiTheme,
        Func<bool> isWindowVisible,
        Func<bool> isTrayMenuEnabled,
        Func<IReadOnlyList<VmDefinition>> getVms,
        Func<IReadOnlyList<HyperVSwitchInfo>> getSwitches,
        Func<Task> refreshTrayDataAction,
        Action<EventHandler> subscribeTrayStateChanged,
        Action<EventHandler> unsubscribeTrayStateChanged,
        Func<string, Task> startVmAction,
        Func<string, Task> stopVmAction,
        Func<string, Task> restartVmAction,
        Func<string, Task> openConsoleAction,
        Func<string, Task> createSnapshotAction,
        Func<string, string, Task> connectVmToSwitchAction,
        Func<string, Task> disconnectVmSwitchAction,
        Action exitAction)
    {
        _showAction = showAction;
        _hideAction = hideAction;
        _toggleControlCenterAction = toggleControlCenterAction;
        _toggleControlCenterCompactAction = toggleControlCenterCompactAction;
        _hideControlCenterAction = hideControlCenterAction;
        _getUiTheme = getUiTheme;
        _isWindowVisible = isWindowVisible;
        _isTrayMenuEnabled = isTrayMenuEnabled;
        _getVms = getVms;
        _getSwitches = getSwitches;
        _refreshTrayDataAction = refreshTrayDataAction;
        _startVmAction = startVmAction;
        _stopVmAction = stopVmAction;
        _restartVmAction = restartVmAction;
        _openConsoleAction = openConsoleAction;
        _createSnapshotAction = createSnapshotAction;
        _connectVmToSwitchAction = connectVmToSwitchAction;
        _disconnectVmSwitchAction = disconnectVmSwitchAction;
        _exitAction = exitAction;
        _unsubscribeTrayStateChanged = unsubscribeTrayStateChanged;

        _menuPalette = ResolveMenuPalette(IsDarkThemePreferred());
        _menuRenderer = new TrayMenuRenderer(_menuPalette);

        _contextMenu = new ContextMenuStrip
        {
            ShowImageMargin = true,
            ShowCheckMargin = false,
            Font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point),
            Renderer = _menuRenderer,
            BackColor = _menuPalette.MenuBackground,
            ForeColor = _menuPalette.MenuForeground,
            Padding = new Padding(8, 6, 8, 6)
        };

        _contextMenu.Opening += (_, _) =>
        {
            _isContextMenuOpen = true;
            UpdateTrayMenu();
        };
        _contextMenu.Closed += (_, _) =>
        {
            _isContextMenuOpen = false;
            if (_hasPendingMenuRefresh)
            {
                _hasPendingMenuRefresh = false;
                UpdateTrayMenuThreadSafe();
            }
        };

        _trayStateChangedHandler = (_, _) => UpdateTrayMenuThreadSafe();
        subscribeTrayStateChanged(_trayStateChangedHandler);

        UpdateTrayMenu();

        _notifyIcon = new NotifyIcon
        {
            Icon = ResolveTrayIcon(),
            Text = "HyperTool",
            ContextMenuStrip = null,
            Visible = true
        };

        _notifyIcon.MouseUp += (_, e) =>
        {
            try
            {
                if (e.Button == MouseButtons.Left)
                {
                    return;
                }

                if (e.Button == MouseButtons.Right)
                {
                    _hideControlCenterAction?.Invoke();
                    HideQuickPanel();

                    if (IsTrayMenuEnabled())
                    {
                        if (_toggleControlCenterAction is not null)
                        {
                            _toggleControlCenterAction();
                            return;
                        }
                    }

                    if (_toggleControlCenterCompactAction is not null)
                    {
                        _toggleControlCenterCompactAction();
                        return;
                    }

                    if (_contextMenu is not null)
                    {
                        _contextMenu.Show(Cursor.Position);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Tray mouse handler failed for button {Button}.", e.Button);
            }
        };

        _notifyIcon.MouseDoubleClick += (_, e) =>
        {
            try
            {
                if (e.Button != MouseButtons.Left)
                {
                    return;
                }

                _hideControlCenterAction?.Invoke();
                HideQuickPanel();

                if (IsWindowVisible())
                {
                    _hideAction?.Invoke();
                }
                else
                {
                    _showAction?.Invoke();
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Tray double-click handler failed.");
            }
        };
        Log.Information("Tray icon initialized.");
    }

    public void UpdateTrayMenu()
    {
        if (_contextMenu is null
            || _showAction is null
            || _hideAction is null
            || _exitAction is null)
        {
            return;
        }

        _contextMenu.Items.Clear();

        var visibilityText = IsWindowVisible() ? "Ausblenden" : "Anzeigen";
        _contextMenu.Items.Add(CreateMenuItem(visibilityText, "⌂", (_, _) => ExecuteMenuAction(() =>
        {
            if (IsWindowVisible())
            {
                _hideAction();
            }
            else
            {
                _showAction();
            }

            return Task.CompletedTask;
        }, "toggle-visibility")));

        _contextMenu.Items.Add(new ToolStripSeparator());

        _contextMenu.Items.Add(CreateMenuItem("Beenden", "⏻", (_, _) => ExecuteMenuAction(() =>
        {
            _exitAction();
            return Task.CompletedTask;
        }, "exit")));

        ApplyMenuStyling(_contextMenu.Items);
    }

    private static string NormalizeSwitchDisplayName(string? switchName)
    {
        return string.IsNullOrWhiteSpace(switchName) ? "Nicht verbunden" : switchName.Trim();
    }

    private async void ExecuteMenuAction(Func<Task> action, string actionName)
    {
        try
        {
            await action();
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Tray menu action failed: {ActionName}", actionName);
        }
    }

    public void Dispose()
    {
        if (_trayStateChangedHandler is not null)
        {
            _unsubscribeTrayStateChanged?.Invoke(_trayStateChangedHandler);
            _trayStateChangedHandler = null;
        }

        if (_notifyIcon is not null)
        {
            _notifyIcon.Visible = false;
            _notifyIcon.Dispose();
            _notifyIcon = null;
        }

        if (_quickPanel is not null)
        {
            _quickPanel.Close();
            _quickPanel.Dispose();
            _quickPanel = null;
        }

        if (_quickPanelToolTip is not null)
        {
            _quickPanelToolTip.Dispose();
            _quickPanelToolTip = null;
        }

        if (_contextMenu is not null)
        {
            _contextMenu.Dispose();
            _contextMenu = null;
        }
    }

    private void UpdateTrayMenuThreadSafe()
    {
        if (_contextMenu is null)
        {
            return;
        }

        if (_isContextMenuOpen)
        {
            _hasPendingMenuRefresh = true;
            return;
        }

        if (_contextMenu.IsHandleCreated)
        {
            _contextMenu.BeginInvoke(new Action(() =>
            {
                UpdateTrayMenu();
                UpdateQuickPanelDataThreadSafe();
            }));
            return;
        }

        UpdateTrayMenu();
        UpdateQuickPanelDataThreadSafe();
    }

    private bool IsTrayMenuEnabled()
    {
        return _isTrayMenuEnabled?.Invoke() ?? true;
    }

    private bool IsWindowVisible()
    {
        return _isWindowVisible?.Invoke() ?? true;
    }

    private void ApplyMenuStyling(ToolStripItemCollection items)
    {
        foreach (ToolStripItem item in items)
        {
            item.BackColor = _menuPalette.MenuBackground;
            item.ForeColor = _menuPalette.MenuForeground;

            if (item is ToolStripSeparator separator)
            {
                separator.Margin = new Padding(8, 4, 8, 4);
                continue;
            }

            if (item is not ToolStripMenuItem menuItem)
            {
                continue;
            }

            menuItem.Padding = new Padding(12, 6, 12, 6);
            menuItem.Margin = new Padding(2, 2, 2, 2);

            if (menuItem.DropDown is ToolStripDropDownMenu dropDown)
            {
                dropDown.ShowImageMargin = true;
                dropDown.ShowCheckMargin = true;
                dropDown.BackColor = _menuPalette.MenuBackground;
                dropDown.ForeColor = _menuPalette.MenuForeground;
                dropDown.Font = _contextMenu?.Font ?? dropDown.Font;
                if (_menuRenderer is not null)
                {
                    dropDown.Renderer = _menuRenderer;
                }

                dropDown.Padding = new Padding(6, 4, 6, 4);
            }

            if (menuItem.DropDownItems.Count > 0)
            {
                ApplyMenuStyling(menuItem.DropDownItems);
            }
        }
    }

    private ToolStripMenuItem CreateMenuItem(string text, string glyph, EventHandler? onClick = null)
    {
        var item = new ToolStripMenuItem(text)
        {
            Image = CreateGlyphImage(glyph, _menuPalette.MenuForeground)
        };

        if (onClick is not null)
        {
            item.Click += onClick;
        }

        return item;
    }

    private static Bitmap CreateGlyphImage(string glyph, Color color)
    {
        var bmp = new Bitmap(16, 16);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);
        using var font = new Font("Segoe UI Symbol", 10f, FontStyle.Regular, GraphicsUnit.Pixel);
        using var brush = new SolidBrush(color);
        var rect = new RectangleF(0, 0, 16, 16);
        var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        g.DrawString(glyph, font, brush, rect, sf);
        return bmp;
    }

    private void ToggleQuickPanel()
    {
        try
        {
            EnsureQuickPanel();
            if (_quickPanel is null)
            {
                return;
            }

            if (_quickPanel.Visible)
            {
                HideQuickPanel();
                return;
            }

            RefreshQuickPanelData();

            var cursor = Cursor.Position;
            var screen = Screen.FromPoint(cursor).WorkingArea;
            var x = Math.Clamp(cursor.X - _quickPanel.Width + 24, screen.Left + 8, screen.Right - _quickPanel.Width - 8);
            var y = Math.Clamp(cursor.Y - _quickPanel.Height - 12, screen.Top + 8, screen.Bottom - _quickPanel.Height - 8);

            _quickPanel.Location = new Point(x, y);
            _quickPanel.Show();
            _quickPanel.BringToFront();
            _quickPanel.TopMost = true;
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Failed to toggle tray quick panel.");
        }
    }

    private void HideQuickPanel()
    {
        if (_quickPanel is null)
        {
            return;
        }

        _quickPanel.Hide();
    }

    private void EnsureQuickPanel()
    {
        if (_quickPanel is not null)
        {
            return;
        }

        var isDark = IsDarkThemePreferred();
        var panelBack = isDark ? Color.FromArgb(0x0F, 0x19, 0x2D) : Color.FromArgb(0xEE, 0xF4, 0xFF);
        var panelCard = isDark ? Color.FromArgb(0x1A, 0x26, 0x41) : Color.FromArgb(0xFC, 0xFE, 0xFF);
        var panelTile = isDark ? Color.FromArgb(0x23, 0x35, 0x58) : Color.FromArgb(0xE8, 0xF1, 0xFF);
        var panelBorder = isDark ? Color.FromArgb(0x42, 0x5D, 0x8E) : Color.FromArgb(0x96, 0xB7, 0xE4);
        var textPrimary = isDark ? Color.FromArgb(0xE8, 0xF1, 0xFF) : Color.FromArgb(0x14, 0x2D, 0x4A);
        var textSecondary = isDark ? Color.FromArgb(0xBA, 0xCA, 0xE4) : Color.FromArgb(0x4A, 0x67, 0x8F);
        var accent = isDark ? Color.FromArgb(0x62, 0xC3, 0xFF) : Color.FromArgb(0x3C, 0x9B, 0xF2);

        _quickPanelToolTip = new ToolTip
        {
            AutomaticDelay = 200,
            ReshowDelay = 100,
            AutoPopDelay = 4000,
            InitialDelay = 250,
            ShowAlways = true,
            IsBalloon = false
        };

        _quickPanel = new Form
        {
            FormBorderStyle = FormBorderStyle.None,
            StartPosition = FormStartPosition.Manual,
            ShowInTaskbar = false,
            TopMost = true,
            Width = 420,
            Height = 506,
            BackColor = panelBorder,
            Padding = new Padding(1)
        };

        _quickPanel.Deactivate += (_, _) => HideQuickPanel();
        _quickPanel.Resize += (_, _) => ApplyRoundedRegion(_quickPanel, 12);

        var root = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = panelBack,
            Padding = new Padding(10, 10, 10, 8)
        };
        _quickPanel.Controls.Add(root);

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 7,
            BackColor = panelBack
        };
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 112));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 134));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 128));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 10));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
        root.Controls.Add(layout);

        var headerRow = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = panelBack
        };
        layout.Controls.Add(headerRow, 0, 0);

        var iconBox = new PictureBox
        {
            Width = 26,
            Height = 26,
            Left = 0,
            Top = 10,
            SizeMode = PictureBoxSizeMode.Zoom,
            Image = ResolvePanelIconImage()
        };
        headerRow.Controls.Add(iconBox);

        var header = new Label
        {
            Text = "HyperTool",
            ForeColor = textPrimary,
            Font = new Font("Segoe UI Semibold", 15f, FontStyle.Bold),
            AutoSize = true,
            Left = 36,
            Top = 8
        };
        headerRow.Controls.Add(header);

        var closeButton = new Button
        {
            Text = "✕",
            Width = 34,
            Height = 28,
            Left = _quickPanel.Width - 60,
            Top = 10,
            FlatStyle = FlatStyle.Flat,
            BackColor = panelTile,
            ForeColor = textPrimary,
            Font = new Font("Segoe UI Symbol", 11f, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleCenter,
            TabStop = false
        };
        closeButton.FlatAppearance.BorderColor = panelBorder;
        closeButton.FlatAppearance.BorderSize = 1;
        closeButton.FlatAppearance.MouseOverBackColor = ControlPaint.Light(accent, 0.05f);
        closeButton.Click += (_, _) => HideQuickPanel();
        closeButton.Resize += (_, _) => ApplyRoundedRegion(closeButton, 8);
        ApplyRoundedRegion(closeButton, 8);
        headerRow.Controls.Add(closeButton);

        var vmCard = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = panelCard,
            Padding = new Padding(10)
        };
        vmCard.Resize += (_, _) => ApplyRoundedRegion(vmCard, 10);
        ApplyRoundedRegion(vmCard, 10);
        layout.Controls.Add(vmCard, 0, 1);

        var vmHeader = new Label
        {
            Text = "Virtuelle Maschine",
            ForeColor = textSecondary,
            Font = new Font("Segoe UI", 9f, FontStyle.Regular),
            Left = 0,
            Top = 0,
            Width = 360,
            Height = 18
        };
        vmCard.Controls.Add(vmHeader);

        _vmNameLabel = new Label
        {
            Text = "-",
            ForeColor = textPrimary,
            Font = new Font("Segoe UI Semibold", 12f, FontStyle.Bold),
            Left = 0,
            Top = 20,
            Width = 360,
            Height = 28,
            TextAlign = ContentAlignment.MiddleLeft
        };
        vmCard.Controls.Add(_vmNameLabel);

        _vmStatusLabel = new Label
        {
            Text = "Status: -",
            ForeColor = textSecondary,
            Font = new Font("Segoe UI", 10f, FontStyle.Regular),
            Left = 0,
            Top = 48,
            Width = 360,
            Height = 24,
            TextAlign = ContentAlignment.MiddleLeft
        };
        vmCard.Controls.Add(_vmStatusLabel);

        var vmSelectRow = new Panel
        {
            Left = 0,
            Top = 74,
            Width = 360,
            Height = 32,
            BackColor = panelCard
        };

        _vmPrevButton = CreatePanelIconButton("◀", "Vorherige VM", panelTile, panelBorder, textPrimary, accent);
        _vmPrevButton.Width = 38;
        _vmPrevButton.Height = 30;
        _vmPrevButton.Location = new Point(0, 1);
        _vmPrevButton.Click += (_, _) =>
        {
            if (_panelVms.Count == 0)
            {
                return;
            }

            _selectedVmIndex = (_selectedVmIndex - 1 + _panelVms.Count) % _panelVms.Count;
            UpdateVmStatusText();
            RebuildSwitchButtons();
            UpdateQuickPanelButtonState();
        };
        vmSelectRow.Controls.Add(_vmPrevButton);

        _vmNextButton = CreatePanelIconButton("▶", "Nächste VM", panelTile, panelBorder, textPrimary, accent);
        _vmNextButton.Width = 38;
        _vmNextButton.Height = 30;
        _vmNextButton.Location = new Point(342, 1);
        _vmNextButton.Click += (_, _) =>
        {
            if (_panelVms.Count == 0)
            {
                return;
            }

            _selectedVmIndex = (_selectedVmIndex + 1) % _panelVms.Count;
            UpdateVmStatusText();
            RebuildSwitchButtons();
            UpdateQuickPanelButtonState();
        };
        vmSelectRow.Controls.Add(_vmNextButton);

        _vmSelectorLabel = new Label
        {
            Left = 44,
            Top = 3,
            Width = 294,
            Height = 26,
            TextAlign = ContentAlignment.MiddleCenter,
            ForeColor = textPrimary,
            BackColor = panelTile,
            BorderStyle = BorderStyle.FixedSingle,
            Font = new Font("Segoe UI", 9f, FontStyle.Regular)
        };
        vmSelectRow.Controls.Add(_vmSelectorLabel);
        vmCard.Controls.Add(vmSelectRow);

        var actionsPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            BackColor = panelBack,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(0)
        };
        actionsPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 31));
        actionsPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 31));
        actionsPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 31));
        actionsPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 31));
        layout.Controls.Add(actionsPanel, 0, 2);

        _startVmButton = CreatePanelMenuButton("▶", "Start VM", panelTile, panelBorder, textPrimary, accent, 374);
        _startVmButton.Click += (_, _) => ExecuteMenuAction(async () =>
        {
            var vmItem = GetSelectedVm();
            if (vmItem is null || _startVmAction is null)
            {
                return;
            }

            await _startVmAction(vmItem.Name);
            await RefreshTrayDataAndPanelAsync();
        }, "panel-start-console");
        _startVmButton.Dock = DockStyle.Fill;
        actionsPanel.Controls.Add(_startVmButton, 0, 0);

        _snapshotVmButton = CreatePanelMenuButton("📸", "Snapshots", panelTile, panelBorder, textPrimary, accent, 374);
        _snapshotVmButton.Click += (_, _) => ExecuteSelectedVmAction(_createSnapshotAction, "panel-snapshot");
        _snapshotVmButton.Dock = DockStyle.Fill;
        actionsPanel.Controls.Add(_snapshotVmButton, 0, 1);

        _toggleVisibilityButton = CreatePanelMenuButton("⌂", "Anzeigen / Ausblenden", panelTile, panelBorder, textPrimary, accent, 374);
        _toggleVisibilityButton.Click += (_, _) => ExecuteMenuAction(() =>
        {
            if (IsWindowVisible())
            {
                _hideAction?.Invoke();
            }
            else
            {
                _showAction?.Invoke();
            }

            UpdateVisibilityButton();
            return Task.CompletedTask;
        }, "panel-toggle-visibility");
        _toggleVisibilityButton.Dock = DockStyle.Fill;
        actionsPanel.Controls.Add(_toggleVisibilityButton, 0, 2);

        _shutdownVmButton = CreatePanelMenuButton("🖥", "Open Console", panelTile, panelBorder, textPrimary, accent, 374);
        _shutdownVmButton.Click += (_, _) => ExecuteMenuAction(async () =>
        {
            var vmItem = GetSelectedVm();
            if (vmItem is null || _startVmAction is null)
            {
                return;
            }

            await _startVmAction(vmItem.Name);
            await RefreshTrayDataAndPanelAsync();
        }, "panel-open-console");
        _shutdownVmButton.Dock = DockStyle.Fill;
        actionsPanel.Controls.Add(_shutdownVmButton, 0, 3);

        var networkLabel = new Label
        {
            Text = "Netzwerk",
            ForeColor = textSecondary,
            Font = new Font("Segoe UI Semibold", 9f, FontStyle.Regular),
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft
        };
        layout.Controls.Add(networkLabel, 0, 3);

        var networkCard = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = panelCard,
            Padding = new Padding(10)
        };
        networkCard.Resize += (_, _) => ApplyRoundedRegion(networkCard, 10);
        ApplyRoundedRegion(networkCard, 10);
        layout.Controls.Add(networkCard, 0, 4);

        _switchButtonsPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = panelCard,
            Padding = new Padding(0)
        };
        networkCard.Controls.Add(_switchButtonsPanel);

        var footer = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = panelBack
        };
        layout.Controls.Add(footer, 0, 6);

        var exitButton = CreatePanelMenuButton("⏻", "Beenden", panelTile, panelBorder, textPrimary, accent, 374);
        exitButton.Dock = DockStyle.Fill;
        exitButton.Click += (_, _) => ExecuteMenuAction(() =>
        {
            _exitAction?.Invoke();
            return Task.CompletedTask;
        }, "panel-exit");
        footer.Controls.Add(exitButton);

        RefreshQuickPanelData();
    }

    private Button CreatePanelIconButton(string symbol, string tooltip, Color backColor, Color borderColor, Color textColor, Color hoverColor)
    {
        var button = new Button
        {
            Text = symbol,
            Width = 44,
            Height = 32,
            FlatStyle = FlatStyle.Flat,
            BackColor = backColor,
            ForeColor = textColor,
            Font = new Font("Segoe UI Symbol", 11f, FontStyle.Bold),
            Padding = new Padding(0),
            Margin = new Padding(0, 0, 6, 0),
            TextAlign = ContentAlignment.MiddleCenter
        };
        button.FlatAppearance.BorderColor = borderColor;
        button.FlatAppearance.BorderSize = 1;
        button.FlatAppearance.MouseOverBackColor = hoverColor;
        button.FlatAppearance.MouseDownBackColor = ControlPaint.Dark(hoverColor, 0.08f);
        button.Resize += (_, _) => ApplyRoundedRegion(button, 10);
        ApplyRoundedRegion(button, 10);
        _quickPanelToolTip?.SetToolTip(button, tooltip);
        return button;
    }

    private Button CreatePanelMenuButton(string symbol, string text, Color backColor, Color borderColor, Color textColor, Color hoverColor, int width)
    {
        var button = new Button
        {
            Text = $"  {symbol}  {text}",
            Width = width,
            Height = 30,
            FlatStyle = FlatStyle.Flat,
            BackColor = backColor,
            ForeColor = textColor,
            Font = new Font("Segoe UI", 10f, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(8, 0, 8, 0),
            Margin = new Padding(0)
        };

        button.FlatAppearance.BorderColor = borderColor;
        button.FlatAppearance.BorderSize = 1;
        button.FlatAppearance.MouseOverBackColor = hoverColor;
        button.FlatAppearance.MouseDownBackColor = ControlPaint.Dark(hoverColor, 0.08f);
        button.Resize += (_, _) => ApplyRoundedRegion(button, 10);
        ApplyRoundedRegion(button, 10);
        return button;
    }

    private static void ConfigureModernComboBox(ComboBox comboBox, Color backColor, Color textColor, Color accent)
    {
        comboBox.DrawMode = DrawMode.OwnerDrawFixed;
        comboBox.ItemHeight = 20;
        comboBox.DrawItem += (_, e) =>
        {
            if (e.Index < 0)
            {
                return;
            }

            var selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            using var bg = new SolidBrush(selected ? accent : backColor);
            using var fg = new SolidBrush(selected ? Color.FromArgb(0x0F, 0x24, 0x3C) : textColor);
            e.Graphics.FillRectangle(bg, e.Bounds);

            var text = comboBox.Items[e.Index]?.ToString() ?? string.Empty;
            var textRect = new RectangleF(e.Bounds.X + 6, e.Bounds.Y + 2, e.Bounds.Width - 8, e.Bounds.Height - 4);
            e.Graphics.DrawString(text, comboBox.Font, fg, textRect);

            e.DrawFocusRectangle();
        };
    }

    private static void ApplyRoundedRegion(Control control, int radius)
    {
        if (control.Width <= 0 || control.Height <= 0)
        {
            return;
        }

        using var path = new GraphicsPath();
        var diameter = radius * 2;
        path.AddArc(0, 0, diameter, diameter, 180, 90);
        path.AddArc(control.Width - diameter, 0, diameter, diameter, 270, 90);
        path.AddArc(control.Width - diameter, control.Height - diameter, diameter, diameter, 0, 90);
        path.AddArc(0, control.Height - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();
        control.Region = new Region(path);
    }

    private void ExecuteSelectedVmAction(Func<string, Task>? vmAction, string actionName)
    {
        ExecuteMenuAction(async () =>
        {
            var vmItem = GetSelectedVm();
            if (vmItem is null || vmAction is null)
            {
                return;
            }

            await vmAction(vmItem.Name);
            await RefreshTrayDataAndPanelAsync();
        }, actionName);
    }

    private async Task RefreshTrayDataAndPanelAsync()
    {
        if (_refreshTrayDataAction is not null)
        {
            await _refreshTrayDataAction();
        }

        RefreshQuickPanelData();
    }

    private void RefreshQuickPanelData()
    {
        var previousVm = GetSelectedVm()?.Name;
        _panelVms.Clear();

        var vms = _getVms?.Invoke() ?? [];
        foreach (var vm in vms)
        {
            var label = string.IsNullOrWhiteSpace(vm.Label) ? vm.Name : vm.Label;
            _panelVms.Add(new VmSelectionItem(
                vm.Name,
                label,
                NormalizeSwitchDisplayName(vm.RuntimeSwitchName),
                string.IsNullOrWhiteSpace(vm.RuntimeState) ? "Unbekannt" : vm.RuntimeState));
        }

        if (_panelVms.Count == 0)
        {
            _selectedVmIndex = -1;
        }
        else if (!string.IsNullOrWhiteSpace(previousVm))
        {
            var idx = _panelVms.FindIndex(vm => string.Equals(vm.Name, previousVm, StringComparison.OrdinalIgnoreCase));
            _selectedVmIndex = idx >= 0 ? idx : 0;
        }
        else if (_selectedVmIndex < 0 || _selectedVmIndex >= _panelVms.Count)
        {
            _selectedVmIndex = 0;
        }

        UpdateVmStatusText();
        RebuildSwitchButtons();
        UpdateVisibilityButton();
        UpdateQuickPanelButtonState();
    }

    private void UpdateVmStatusText()
    {
        if (_vmStatusLabel is null || _vmNameLabel is null || _vmSelectorLabel is null)
        {
            return;
        }

        var vm = GetSelectedVm();
        if (vm is null)
        {
            _vmNameLabel.Text = "-";
            _vmSelectorLabel.Text = "Keine VM";
            _vmStatusLabel.Text = "Status: -";
            return;
        }

        _vmNameLabel.Text = vm.Label;
        _vmSelectorLabel.Text = $"{_selectedVmIndex + 1} / {_panelVms.Count}";

        _vmStatusLabel.Text = $"{vm.RuntimeState} · {vm.RuntimeSwitchName}";
    }

    private void RebuildSwitchButtons()
    {
        if (_switchButtonsPanel is null)
        {
            return;
        }

        _switchButtonsPanel.Controls.Clear();

        var isDark = IsDarkThemePreferred();
        var panelBorder = isDark ? Color.FromArgb(0x4E, 0x66, 0x8A) : Color.FromArgb(0x92, 0xB3, 0xDF);
        var textPrimary = isDark ? Color.FromArgb(0xE7, 0xF0, 0xFF) : Color.FromArgb(0x13, 0x2A, 0x44);
        var accent = isDark ? Color.FromArgb(0x67, 0xC8, 0xFF) : Color.FromArgb(0x5A, 0xAE, 0xF2);
        var panelTile = isDark ? Color.FromArgb(0x2A, 0x37, 0x58) : Color.FromArgb(0xDB, 0xEA, 0xFF);

        var selectedVm = GetSelectedVm();
        if (selectedVm is null)
        {
            return;
        }

        var title = new Label
        {
            Text = $"Aktiv: {NormalizeSwitchDisplayName(selectedVm.RuntimeSwitchName)}",
            ForeColor = textPrimary,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Width = 360,
            Height = 20,
            Left = 0,
            Top = 0,
            TextAlign = ContentAlignment.MiddleLeft,
            Margin = new Padding(0)
        };
        _switchButtonsPanel.Controls.Add(title);

        _panelSwitches.Clear();
        foreach (var vmSwitch in _getSwitches?.Invoke() ?? [])
        {
            if (!string.IsNullOrWhiteSpace(vmSwitch.Name))
            {
                _panelSwitches.Add(vmSwitch.Name);
            }
        }

        if (_panelSwitches.Count == 0)
        {
            _selectedSwitchIndex = -1;
        }
        else
        {
            var idx = _panelSwitches.FindIndex(name => string.Equals(name, selectedVm.RuntimeSwitchName, StringComparison.OrdinalIgnoreCase));
            if (idx >= 0)
            {
                _selectedSwitchIndex = idx;
            }
            else if (_selectedSwitchIndex < 0 || _selectedSwitchIndex >= _panelSwitches.Count)
            {
                _selectedSwitchIndex = 0;
            }
        }

        var selectRow = new Panel
        {
            Width = 360,
            Height = 34,
            Left = 0,
            Top = 24,
            BackColor = Color.Transparent,
            Margin = new Padding(0)
        };

        _switchPrevButton = CreatePanelIconButton("◀", "Vorheriger Switch", panelTile, panelBorder, textPrimary, accent);
        _switchPrevButton.Width = 38;
        _switchPrevButton.Height = 30;
        _switchPrevButton.Location = new Point(0, 2);
        _switchPrevButton.Click += (_, _) =>
        {
            if (_panelSwitches.Count == 0)
            {
                return;
            }

            _selectedSwitchIndex = (_selectedSwitchIndex - 1 + _panelSwitches.Count) % _panelSwitches.Count;
            UpdateSwitchSelectorLabel();
            UpdateQuickPanelButtonState();
        };
        selectRow.Controls.Add(_switchPrevButton);

        _switchNextButton = CreatePanelIconButton("▶", "Nächster Switch", panelTile, panelBorder, textPrimary, accent);
        _switchNextButton.Width = 38;
        _switchNextButton.Height = 30;
        _switchNextButton.Location = new Point(322, 2);
        _switchNextButton.Click += (_, _) =>
        {
            if (_panelSwitches.Count == 0)
            {
                return;
            }

            _selectedSwitchIndex = (_selectedSwitchIndex + 1) % _panelSwitches.Count;
            UpdateSwitchSelectorLabel();
            UpdateQuickPanelButtonState();
        };
        selectRow.Controls.Add(_switchNextButton);

        _switchStatusLabel = new Label
        {
            Left = 44,
            Top = 4,
            Width = 274,
            Height = 26,
            BackColor = panelTile,
            ForeColor = textPrimary,
            BorderStyle = BorderStyle.FixedSingle,
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font("Segoe UI", 9f, FontStyle.Regular)
        };
        selectRow.Controls.Add(_switchStatusLabel);
        _switchButtonsPanel.Controls.Add(selectRow);

        _switchConnectButton = CreatePanelMenuButton("⇄", "Verbinden", panelTile, panelBorder, textPrimary, accent, 360);
        _switchConnectButton.Left = 0;
        _switchConnectButton.Top = 62;
        _switchConnectButton.Click += (_, _) => ExecuteMenuAction(async () =>
        {
            var switchName = GetSelectedSwitchName();
            if (_connectVmToSwitchAction is null || string.IsNullOrWhiteSpace(switchName))
            {
                return;
            }

            await _connectVmToSwitchAction(selectedVm.Name, switchName);
            await RefreshTrayDataAndPanelAsync();
        }, $"panel-connect-{selectedVm.Name}");
        _switchButtonsPanel.Controls.Add(_switchConnectButton);

        _switchDisconnectButton = CreatePanelMenuButton("⛔", "Trennen", panelTile, panelBorder, textPrimary, accent, 360);
        _switchDisconnectButton.Left = 0;
        _switchDisconnectButton.Top = 94;
        _switchDisconnectButton.Click += (_, _) => ExecuteMenuAction(async () =>
        {
            if (_disconnectVmSwitchAction is null)
            {
                return;
            }

            await _disconnectVmSwitchAction(selectedVm.Name);
            await RefreshTrayDataAndPanelAsync();
        }, $"panel-disconnect-{selectedVm.Name}");
        _switchButtonsPanel.Controls.Add(_switchDisconnectButton);

        UpdateSwitchSelectorLabel();
    }

    private void UpdateVisibilityButton()
    {
        if (_toggleVisibilityButton is null)
        {
            return;
        }

        var visible = IsWindowVisible();
        _toggleVisibilityButton.Text = visible ? "  ⌂  Ausblenden" : "  ⌂  Anzeigen";
        _quickPanelToolTip?.SetToolTip(_toggleVisibilityButton, visible ? "Fenster ausblenden" : "Fenster anzeigen");
    }

    private void UpdateQuickPanelButtonState()
    {
        var hasVm = GetSelectedVm() is not null;
        var hasSwitch = !string.IsNullOrWhiteSpace(GetSelectedSwitchName());
        var trayEnabled = IsTrayMenuEnabled();

        if (_startVmButton is not null) _startVmButton.Enabled = trayEnabled && hasVm;
        if (_shutdownVmButton is not null) _shutdownVmButton.Enabled = trayEnabled && hasVm;
        if (_snapshotVmButton is not null) _snapshotVmButton.Enabled = trayEnabled && hasVm;
        if (_vmPrevButton is not null) _vmPrevButton.Enabled = trayEnabled && _panelVms.Count > 1;
        if (_vmNextButton is not null) _vmNextButton.Enabled = trayEnabled && _panelVms.Count > 1;
        if (_switchPrevButton is not null) _switchPrevButton.Enabled = trayEnabled && _panelSwitches.Count > 1;
        if (_switchNextButton is not null) _switchNextButton.Enabled = trayEnabled && _panelSwitches.Count > 1;
        if (_switchConnectButton is not null) _switchConnectButton.Enabled = trayEnabled && hasVm && hasSwitch;
        if (_switchDisconnectButton is not null) _switchDisconnectButton.Enabled = trayEnabled && hasVm;
    }

    private VmSelectionItem? GetSelectedVm()
    {
        if (_selectedVmIndex < 0 || _selectedVmIndex >= _panelVms.Count)
        {
            return null;
        }

        return _panelVms[_selectedVmIndex];
    }

    private string? GetSelectedSwitchName()
    {
        if (_selectedSwitchIndex < 0 || _selectedSwitchIndex >= _panelSwitches.Count)
        {
            return null;
        }

        return _panelSwitches[_selectedSwitchIndex];
    }

    private void UpdateSwitchSelectorLabel()
    {
        if (_switchStatusLabel is null)
        {
            return;
        }

        var selectedSwitch = GetSelectedSwitchName();
        _switchStatusLabel.Text = string.IsNullOrWhiteSpace(selectedSwitch)
            ? "Kein Switch verfügbar"
            : $"{_selectedSwitchIndex + 1} / {_panelSwitches.Count}: {selectedSwitch}";
    }

    private void UpdateQuickPanelDataThreadSafe()
    {
        if (_quickPanel is null || _quickPanel.IsDisposed || !_quickPanel.Visible)
        {
            return;
        }

        if (_quickPanel.InvokeRequired)
        {
            _quickPanel.BeginInvoke(new Action(RefreshQuickPanelData));
            return;
        }

        RefreshQuickPanelData();
    }

    private bool IsDarkThemePreferred()
    {
        try
        {
            var uiTheme = _getUiTheme?.Invoke();
            if (!string.IsNullOrWhiteSpace(uiTheme))
            {
                return string.Equals(uiTheme, "Dark", StringComparison.OrdinalIgnoreCase);
            }
        }
        catch
        {
        }

        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            var value = key?.GetValue("AppsUseLightTheme");
            return value is not null && Convert.ToInt32(value) == 0;
        }
        catch
        {
            return true;
        }
    }

    private static TrayMenuPalette ResolveMenuPalette(bool isDark)
    {
        return isDark ? TrayMenuPalette.CreateDark() : TrayMenuPalette.CreateLight();
    }

    private static Icon ResolveTrayIcon()
    {
        try
        {
            var preferredTrayIconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "HyperTool.Tray.ico");
            if (File.Exists(preferredTrayIconPath))
            {
                using var icon = new Icon(preferredTrayIconPath);
                return OptimizeIconForTray(icon);
            }
        }
        catch
        {
        }

        try
        {
            var iconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "HyperTool.ico");
            if (File.Exists(iconPath))
            {
                using var icon = new Icon(iconPath);
                return OptimizeIconForTray(icon);
            }
        }
        catch
        {
        }

        try
        {
            var processPath = Environment.ProcessPath;
            if (!string.IsNullOrWhiteSpace(processPath))
            {
                var extracted = Icon.ExtractAssociatedIcon(processPath);
                if (extracted is not null)
                {
                    return OptimizeIconForTray(extracted);
                }
            }
        }
        catch
        {
        }

        return SystemIcons.Application;
    }

    private static Icon OptimizeIconForTray(Icon icon)
    {
        try
        {
            using var bitmap = icon.ToBitmap();
            using var fitted = FitBitmapToTray(bitmap);
            var optimized = CreateIconFromBitmap(fitted);
            if (optimized is not null)
            {
                return optimized;
            }
        }
        catch
        {
        }

        return (Icon)icon.Clone();
    }

    private static Bitmap FitBitmapToTray(Bitmap source)
    {
        var bounds = GetOpaqueBounds(source);
        var output = new Bitmap(64, 64);

        using var g = Graphics.FromImage(output);
        g.SmoothingMode = SmoothingMode.HighQuality;
        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
        g.Clear(Color.Transparent);

        var target = new Rectangle(0, 0, 64, 64);
        g.DrawImage(source, target, bounds, GraphicsUnit.Pixel);
        return output;
    }

    private static Rectangle GetOpaqueBounds(Bitmap bitmap)
    {
        var minX = bitmap.Width;
        var minY = bitmap.Height;
        var maxX = -1;
        var maxY = -1;

        for (var y = 0; y < bitmap.Height; y++)
        {
            for (var x = 0; x < bitmap.Width; x++)
            {
                if (bitmap.GetPixel(x, y).A <= 10)
                {
                    continue;
                }

                if (x < minX) minX = x;
                if (y < minY) minY = y;
                if (x > maxX) maxX = x;
                if (y > maxY) maxY = y;
            }
        }

        return maxX < minX || maxY < minY
            ? new Rectangle(0, 0, bitmap.Width, bitmap.Height)
            : Rectangle.FromLTRB(minX, minY, maxX + 1, maxY + 1);
    }

    private static Icon? CreateIconFromBitmap(Bitmap bitmap)
    {
        IntPtr iconHandle = IntPtr.Zero;

        try
        {
            iconHandle = bitmap.GetHicon();
            using var icon = Icon.FromHandle(iconHandle);
            return (Icon)icon.Clone();
        }
        catch
        {
            return null;
        }
        finally
        {
            if (iconHandle != IntPtr.Zero)
            {
                _ = DestroyIcon(iconHandle);
            }
        }
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyIcon(IntPtr hIcon);

    private static Image? ResolvePanelIconImage()
    {
        var iconCandidates = new[]
        {
            Path.Combine(AppContext.BaseDirectory, "Assets", "HyperTool.Tray.ico"),
            Path.Combine(AppContext.BaseDirectory, "Assets", "HyperTool.ico")
        };

        foreach (var iconPath in iconCandidates)
        {
            try
            {
                if (File.Exists(iconPath))
                {
                    using var icon = new Icon(iconPath);
                    return icon.ToBitmap();
                }
            }
            catch
            {
            }
        }

        try
        {
            using var fallback = ResolveTrayIcon();
            return fallback.ToBitmap();
        }
        catch
        {
            return null;
        }
    }

    private sealed record VmSelectionItem(string Name, string Label, string RuntimeSwitchName, string RuntimeState)
    {
        public override string ToString() => Label;
    }

    private readonly record struct TrayMenuPalette(
        Color MenuBackground,
        Color MenuForeground,
        Color ItemHoverBackground,
        Color ItemPressedBackground,
        Color ItemBorder,
        Color Separator)
    {
        public static TrayMenuPalette CreateLight()
        {
            return new TrayMenuPalette(
                MenuBackground: Color.FromArgb(246, 251, 255),
                MenuForeground: Color.FromArgb(20, 42, 70),
                ItemHoverBackground: Color.FromArgb(224, 238, 255),
                ItemPressedBackground: Color.FromArgb(210, 228, 252),
                ItemBorder: Color.FromArgb(150, 186, 230),
                Separator: Color.FromArgb(182, 207, 235));
        }

        public static TrayMenuPalette CreateDark()
        {
            return new TrayMenuPalette(
                MenuBackground: Color.FromArgb(18, 31, 53),
                MenuForeground: Color.FromArgb(232, 242, 255),
                ItemHoverBackground: Color.FromArgb(40, 59, 92),
                ItemPressedBackground: Color.FromArgb(53, 74, 112),
                ItemBorder: Color.FromArgb(82, 120, 178),
                Separator: Color.FromArgb(71, 101, 148));
        }
    }

    private sealed class TrayMenuRenderer : ToolStripProfessionalRenderer
    {
        public TrayMenuRenderer(TrayMenuPalette palette)
            : base(new TrayMenuColorTable(palette))
        {
            RoundedEdges = false;
        }
    }

    private sealed class TrayMenuColorTable : ProfessionalColorTable
    {
        private readonly TrayMenuPalette _palette;

        public TrayMenuColorTable(TrayMenuPalette palette)
        {
            _palette = palette;
        }

        public override Color ToolStripDropDownBackground => _palette.MenuBackground;
        public override Color MenuItemSelected => _palette.ItemHoverBackground;
        public override Color MenuItemSelectedGradientBegin => _palette.ItemHoverBackground;
        public override Color MenuItemSelectedGradientEnd => _palette.ItemHoverBackground;
        public override Color MenuItemPressedGradientBegin => _palette.ItemPressedBackground;
        public override Color MenuItemPressedGradientMiddle => _palette.ItemPressedBackground;
        public override Color MenuItemPressedGradientEnd => _palette.ItemPressedBackground;
        public override Color MenuBorder => _palette.ItemBorder;
        public override Color SeparatorDark => _palette.Separator;
        public override Color SeparatorLight => _palette.Separator;
        public override Color ImageMarginGradientBegin => _palette.MenuBackground;
        public override Color ImageMarginGradientMiddle => _palette.MenuBackground;
        public override Color ImageMarginGradientEnd => _palette.MenuBackground;
    }
}
