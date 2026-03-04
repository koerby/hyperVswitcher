using HyperTool.Models;
using HyperTool.Services;
using HyperTool.Guest.Views;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml;
using System.Diagnostics;
using System.IO.Pipes;
using System.Net;
using System.Net.Sockets;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Collections.Concurrent;
using Windows.Graphics;
using Windows.UI;

namespace HyperTool.Guest;

public sealed partial class App : Application
{
    private const int GuestUsbAutoRefreshSeconds = 4;
    private const int SharedFolderCatalogFetchMaxAttempts = 3;
    private static readonly TimeSpan RecurringWarnRateLimitInterval = TimeSpan.FromMinutes(2);
    private const string GuestWindowTitle = "HyperTool Guest";
    private const string GuestHeadline = "HyperTool Guest";
    private const string GuestIconUri = "ms-appx:///Assets/HyperTool.Guest.Icon.Transparent.png";

    private const string SingleInstanceMutexName = @"Local\HyperTool.Guest.SingleInstance";
    private const string SingleInstancePipeName = "HyperTool.Guest.SingleInstance.Activate";
    private const int SwRestore = 9;
    private const int SwShow = 5;

    private readonly IUsbIpService _usbService = new UsbIpdCliService();

    private GuestMainWindow? _mainWindow;
    private ITrayService? _trayService;
    private GuestTrayControlCenterWindow? _trayControlCenterWindow;
    private GuestConfig? _config;
    private string _configPath = GuestConfigService.DefaultConfigPath;

    private bool _isExitRequested;
    private bool _isExitAnimationRunning;
    private bool _isThemeWindowReopenInProgress;
    private bool _isTrayFunctional;
    private bool _minimizeToTray = true;

    private Mutex? _singleInstanceMutex;
    private CancellationTokenSource? _singleInstanceServerCts;
    private Task? _singleInstanceServerTask;
    private bool _singleInstanceOwned;
    private bool _pendingSingleInstanceShow;

    private readonly List<UsbIpDeviceInfo> _usbDevices = [];
    private string? _selectedUsbBusId;
    private bool _isUsbClientAvailable;
    private bool _usbClientMissingLogged;
    private bool _currentUsbTransportUseHyperVSocket = true;
    private HyperVSocketUsbGuestProxy? _usbHyperVSocketProxy;
    private bool _usbHyperVSocketTransportActive;
    private bool _usbIpFallbackActive;
    private bool _usbHyperVSocketServiceReachable;
    private int _usbHyperVSocketProbeFailureCount;
    private CancellationTokenSource? _usbDiagnosticsCts;
    private Task? _usbDiagnosticsTask;
    private CancellationTokenSource? _usbAutoRefreshCts;
    private Task? _usbAutoRefreshTask;
    private readonly GuestWinFspMountService _winFspMountService = GuestWinFspMountRegistry.Instance;
    private readonly SemaphoreSlim _sharedFolderReconnectGate = new(1, 1);
    private CancellationTokenSource? _sharedFolderAutoMountCts;
    private Task? _sharedFolderAutoMountTask;
    private readonly SemaphoreSlim _usbRefreshGate = new(1, 1);
    private readonly ConcurrentDictionary<string, RateLimitedWarnState> _rateLimitedWarnStates = new(StringComparer.Ordinal);
    private DateTimeOffset? _sharedFolderReconnectLastRunUtc;
    private string _sharedFolderReconnectLastSummary = "Noch kein Lauf";

    private sealed class RateLimitedWarnState
    {
        public DateTimeOffset LastLoggedAtUtc { get; set; }
        public int SuppressedCount { get; set; }
    }

    private sealed class SharedFolderReconnectCycleResult
    {
        public int Attempted { get; init; }
        public int NewlyMounted { get; init; }
        public int Failed { get; init; }
    }

    public App()
    {
        Microsoft.UI.Xaml.Application.LoadComponent(this, new Uri("ms-appx:///App.xaml"));

        UnhandledException += (_, e) =>
        {
            var ex = e.Exception;
            GuestLogger.Error("ui.unhandled", ex?.Message ?? "Unbekannter UI-Fehler", new
            {
                exceptionType = ex?.GetType().FullName,
                stackTrace = ex?.StackTrace,
                exception = ex?.ToString()
            });
            e.Handled = true;
        };
    }

    private static void EnsureControlResources()
    {
        if (Current is not Application app)
        {
            return;
        }

        var resources = app.Resources;

        EnsureBrushResource(resources, "TabViewScrollButtonBackground", Color.FromArgb(0x00, 0x00, 0x00, 0x00));
        EnsureBrushResource(resources, "TabViewScrollButtonBackgroundPointerOver", Color.FromArgb(0x1F, 0x80, 0x80, 0x80));
        EnsureBrushResource(resources, "TabViewScrollButtonBackgroundPressed", Color.FromArgb(0x33, 0x80, 0x80, 0x80));
        EnsureBrushResource(resources, "TabViewScrollButtonForeground", Colors.WhiteSmoke);
        EnsureBrushResource(resources, "TabViewScrollButtonForegroundPointerOver", Colors.White);
        EnsureBrushResource(resources, "TabViewScrollButtonForegroundPressed", Colors.White);
        EnsureBrushResource(resources, "TabViewScrollButtonForegroundDisabled", Color.FromArgb(0x8F, 0xB0, 0xB0, 0xB0));

        EnsureBrushResource(resources, "TabViewButtonBackground", Color.FromArgb(0x00, 0x00, 0x00, 0x00));
        EnsureBrushResource(resources, "TabViewButtonBackgroundPointerOver", Color.FromArgb(0x1F, 0x90, 0x90, 0x90));
        EnsureBrushResource(resources, "TabViewButtonBackgroundPressed", Color.FromArgb(0x33, 0x90, 0x90, 0x90));
        EnsureBrushResource(resources, "TabViewButtonForeground", Colors.WhiteSmoke);
        EnsureBrushResource(resources, "TabViewButtonForegroundPointerOver", Colors.White);
        EnsureBrushResource(resources, "TabViewButtonForegroundPressed", Colors.White);
        EnsureBrushResource(resources, "TabViewButtonForegroundDisabled", Color.FromArgb(0x7F, 0xA5, 0xA5, 0xA5));
        EnsureBrushResource(resources, "TabViewButtonBorderBrush", Color.FromArgb(0x00, 0x00, 0x00, 0x00));
        EnsureBrushResource(resources, "TabViewButtonBorderBrushPointerOver", Color.FromArgb(0x3F, 0x90, 0x90, 0x90));
        EnsureBrushResource(resources, "TabViewButtonBorderBrushPressed", Color.FromArgb(0x55, 0x90, 0x90, 0x90));

        try
        {
            if (!resources.MergedDictionaries.OfType<XamlControlsResources>().Any())
            {
                resources.MergedDictionaries.Add(new XamlControlsResources());
            }
        }
        catch
        {
        }
    }

    private static void EnsureBrushResource(ResourceDictionary resources, string key, Color color)
    {
        if (resources.ContainsKey(key))
        {
            return;
        }

        resources[key] = new SolidColorBrush(color);
    }

    protected override async void OnLaunched(LaunchActivatedEventArgs args)
    {
        EnsureControlResources();

        var parsedArgs = SplitArgs(args.Arguments);
        var skipStartScreen = parsedArgs.Any(arg => string.Equals(arg, "--skip-start-screen", StringComparison.OrdinalIgnoreCase));

        if (IsCommandMode(parsedArgs))
        {
            var code = await GuestCli.ExecuteAsync(parsedArgs);
            Environment.ExitCode = code;
            Exit();
            return;
        }

        if (!TryInitializeSingleInstance())
        {
            Exit();
            return;
        }

        _configPath = ResolveConfigPath(parsedArgs);
        _config = GuestConfigService.LoadOrCreate(_configPath, out _);
        _currentUsbTransportUseHyperVSocket = _config.Usb?.UseHyperVSocket != false;
        GuestLogger.Initialize(_config.Logging);
        await RefreshUsbClientAvailabilityAsync();
        UpdateUsbTransportBridge();
        StartUsbDiagnosticsLoop();
        StartUsbAutoRefreshLoop();
        StartSharedFolderAutoMountLoop();

        ApplyStartWithWindows(_config);

        _mainWindow = new GuestMainWindow(
            _config,
            RefreshUsbDevicesAsync,
            ConnectUsbAsync,
            DisconnectUsbAsync,
            SaveConfigAsync,
            RestartForThemeChangeAsync,
            RunTransportDiagnosticsTestAsync,
            DiscoverUsbHostAddressAsync,
            FetchHostSharedFoldersAsync,
            _isUsbClientAvailable);
        _mainWindow.UpdateHostFeatureAvailability(
            usbFeatureEnabledByHost: _config.Usb?.HostFeatureEnabled != false,
            sharedFoldersFeatureEnabledByHost: _config.SharedFolders?.HostFeatureEnabled != false,
            hostName: _config.Usb?.HostName);
        UpdateUsbDiagnosticsPanel();
        UpdateSharedFolderReconnectStatusPanel();

        _mainWindow.AppWindow.Closing += OnMainWindowClosing;

        _minimizeToTray = _config.Ui.MinimizeToTray;
        _mainWindow.ApplyTheme(_config.Ui.Theme);

        if (!skipStartScreen)
        {
            try
            {
                _mainWindow.PrepareStartupSplash();
            }
            catch
            {
            }
        }

        TryInitializeTray();

        if (_config.Ui.StartMinimized && _isTrayFunctional && !_pendingSingleInstanceShow)
        {
            _mainWindow.AppWindow.Hide();
        }
        else
        {
            _mainWindow.Activate();
            _mainWindow.ApplyTheme(_config.Ui.Theme);

            if (!skipStartScreen)
            {
                try
                {
                    await _mainWindow.PlayStartupAnimationAsync();
                    _mainWindow.ApplyTheme(_config.Ui.Theme);
                }
                catch
                {
                }
            }
        }

        _ = RunDeferredStartupTasksAsync();

        if (_pendingSingleInstanceShow)
        {
            BringMainWindowToFront();
        }
    }

    private async Task SaveConfigAsync(GuestConfig config)
    {
        var previousTransportMode = _currentUsbTransportUseHyperVSocket;
        var nextTransportMode = config.Usb?.UseHyperVSocket != false;

        if (previousTransportMode != nextTransportMode)
        {
            await DisconnectAllAttachedUsbForTransportSwitchAsync();
        }

        _config = config;
        GuestConfigService.Save(_configPath, config);

        ApplyStartWithWindows(config);
        _minimizeToTray = config.Ui.MinimizeToTray;
        UpdateUsbTransportBridge();
        _currentUsbTransportUseHyperVSocket = nextTransportMode;
        UpdateTrayControlCenterView();
        TriggerSharedFolderReconnectCycle();
        UpdateSharedFolderReconnectStatusPanel();

        await Task.CompletedTask;
    }

    private async Task RunDeferredStartupTasksAsync()
    {
        try
        {
            var identity = await FetchHostIdentityViaHyperVSocketAsync(CancellationToken.None);
            if (identity is not null)
            {
                ApplyHostIdentityToConfig(identity, persistConfig: true);
            }
        }
        catch
        {
        }

        try
        {
            await RefreshUsbDevicesAsync();
        }
        catch (Exception ex)
        {
            GuestLogger.Warn("startup.usb_refresh_failed", ex.Message);
        }

        try
        {
            if (_mainWindow is not null)
            {
                await _mainWindow.CheckForUpdatesOnStartupIfEnabledAsync();
            }
        }
        catch (Exception ex)
        {
            GuestLogger.Warn("startup.updatecheck_failed", ex.Message);
        }
    }

    private void ApplyStartWithWindows(GuestConfig config)
    {
        var startupService = new StartupService();
        if (!startupService.SetStartWithWindows(config.Ui.StartWithWindows, "HyperTool.Guest", Environment.ProcessPath ?? string.Empty, out var startupError)
            && !string.IsNullOrWhiteSpace(startupError))
        {
            GuestLogger.Warn("startup.apply_failed", startupError);
        }
    }

    private void OnMainWindowClosing(Microsoft.UI.Windowing.AppWindow sender, Microsoft.UI.Windowing.AppWindowClosingEventArgs args)
    {
        if (_isExitRequested || _isThemeWindowReopenInProgress)
        {
            return;
        }

        if (_isTrayFunctional && _minimizeToTray)
        {
            args.Cancel = true;
            sender.Hide();
            UpdateTrayControlCenterView();
            return;
        }

        args.Cancel = true;
        _ = ExitWithAnimationAsync();
    }

    private async Task ExitWithAnimationAsync(bool showExitScreen = true)
    {
        if (_isExitAnimationRunning)
        {
            return;
        }

        _isExitAnimationRunning = true;

        try
        {
            var exitAnimationStopwatch = Stopwatch.StartNew();
            const int inlineExitAnimationDurationMs = 2000;

            if (_mainWindow is not null)
            {
                try
                {
                    await _mainWindow.PlayExitAnimationAsync();
                }
                catch
                {
                }
            }

            _isExitRequested = true;

            _trayControlCenterWindow?.Close();
            _trayControlCenterWindow = null;

            StopUsbDiagnosticsLoop();
            StopUsbAutoRefreshLoop();
            StopSharedFolderAutoMountLoop();

            await DisconnectAllAttachedUsbOnExitAsync();

            _usbHyperVSocketProxy?.Dispose();
            _usbHyperVSocketProxy = null;

            _winFspMountService.Dispose();

            _trayService?.Dispose();
            _trayService = null;

            ShutdownSingleInstanceInfrastructure();

            if (showExitScreen)
            {
                var remainingMs = inlineExitAnimationDurationMs - (int)exitAnimationStopwatch.ElapsedMilliseconds;
                if (remainingMs > 0)
                {
                    try
                    {
                        await Task.Delay(remainingMs);
                    }
                    catch
                    {
                    }
                }
            }

            try
            {
                _mainWindow?.AppWindow.Hide();
            }
            catch
            {
            }

            _mainWindow?.Close();
            Exit();
        }
        finally
        {
            _isExitAnimationRunning = false;
        }
    }

    private async Task RestartForThemeChangeAsync(string targetTheme)
    {
        if (_isThemeWindowReopenInProgress)
        {
            return;
        }

        if (_mainWindow is null || _config is null)
        {
            return;
        }

        _isThemeWindowReopenInProgress = true;
        var previousWindow = _mainWindow;

        try
        {
            PointInt32 previousPosition;
            SizeInt32 previousSize;
            bool previousVisible;
            var previousMenuIndex = previousWindow.SelectedMenuIndex;

            try
            {
                previousPosition = previousWindow.AppWindow.Position;
                previousSize = previousWindow.AppWindow.Size;
                previousVisible = previousWindow.AppWindow.IsVisible;
            }
            catch
            {
                previousPosition = new PointInt32(120, 120);
                previousSize = new SizeInt32(GuestMainWindow.DefaultWindowWidth, GuestMainWindow.DefaultWindowHeight);
                previousVisible = true;
            }

            _trayControlCenterWindow?.Close();
            _trayControlCenterWindow = null;

            _trayService?.Dispose();
            _trayService = null;

            try
            {
                previousWindow.AppWindow.Show();
            }
            catch
            {
            }

            try
            {
                previousWindow.Activate();
            }
            catch
            {
            }

            try
            {
                await Task.Delay(70);
                await previousWindow.PlayThemeReloadSplashAsync(targetTheme);
            }
            catch
            {
            }

            _config.Ui.Theme = GuestConfigService.NormalizeTheme(targetTheme);

            var nextWindow = new GuestMainWindow(
                _config,
                RefreshUsbDevicesAsync,
                ConnectUsbAsync,
                DisconnectUsbAsync,
                SaveConfigAsync,
                RestartForThemeChangeAsync,
                RunTransportDiagnosticsTestAsync,
                DiscoverUsbHostAddressAsync,
                FetchHostSharedFoldersAsync,
                _isUsbClientAvailable);

            try
            {
                nextWindow.PrepareLifecycleGuard("Design wird neu geladen …");
            }
            catch
            {
            }

            nextWindow.AppWindow.Closing += OnMainWindowClosing;
            _mainWindow = nextWindow;
            _minimizeToTray = _config.Ui.MinimizeToTray;
            _mainWindow.ApplyTheme(_config.Ui.Theme);
            _mainWindow.SelectMenuIndex(previousMenuIndex);

            TryInitializeTray();
            await RefreshUsbDevicesAsync();

            try
            {
                nextWindow.AppWindow.Move(previousPosition);
                nextWindow.AppWindow.Resize(previousSize);
            }
            catch
            {
            }

            if (previousVisible)
            {
                try
                {
                    previousWindow.AppWindow.Hide();
                }
                catch
                {
                }

                await Task.Delay(70);

                try
                {
                    previousWindow.Close();
                }
                catch
                {
                }

                await Task.Delay(50);

                try
                {
                    nextWindow.AppWindow.Show();
                }
                catch
                {
                }

                nextWindow.Activate();
                nextWindow.ApplyTheme(_config.Ui.Theme);

                try
                {
                    await nextWindow.DismissLifecycleGuardAsync();
                }
                catch
                {
                }
            }
            else if (_isTrayFunctional && _minimizeToTray)
            {
                try
                {
                    nextWindow.AppWindow.Hide();
                }
                catch
                {
                }
            }
            else
            {
                nextWindow.Activate();
                nextWindow.ApplyTheme(_config.Ui.Theme);

                try
                {
                    await nextWindow.DismissLifecycleGuardAsync();
                }
                catch
                {
                }
            }

            UpdateTrayControlCenterView();
            UpdateUsbDiagnosticsPanel();
            UpdateSharedFolderReconnectStatusPanel();
        }
        catch (Exception ex)
        {
            GuestLogger.Warn("theme.reopen_failed", ex.Message, new { targetTheme });

            try
            {
                _mainWindow = previousWindow;
                _mainWindow.AppWindow.Show();
                _mainWindow.Activate();
            }
            catch
            {
            }
        }
        finally
        {
            _isThemeWindowReopenInProgress = false;
        }
    }

    private void TryInitializeTray()
    {
        if (_mainWindow is null)
        {
            return;
        }

        try
        {
            _trayService = new HyperTool.Services.TrayService();
            InitializeTrayServiceCompat(_trayService);

            _isTrayFunctional = true;

            EnsureTrayControlCenterWindow();
            UpdateTrayControlCenterView();
        }
        catch (Exception ex)
        {
            _isTrayFunctional = false;
            GuestLogger.Warn("tray.init_failed", ex.Message);
        }
    }

    private void InitializeTrayServiceCompat(ITrayService trayService)
    {
        var initializeMethod = trayService.GetType().GetMethod("Initialize");
        if (initializeMethod is null)
        {
            throw new InvalidOperationException("TrayService.Initialize wurde nicht gefunden.");
        }

        var commonArguments = new object?[]
        {
            (Action)(() =>
            {
                _mainWindow!.AppWindow.Show();
                _mainWindow.Activate();
                UpdateTrayControlCenterView();
            }),
            (Action)(() =>
            {
                _mainWindow!.AppWindow.Hide();
                UpdateTrayControlCenterView();
            }),
            (Action)ToggleTrayControlCenter,
            (Action)ToggleTrayControlCenter,
            (Action)HideTrayControlCenter,
            (Func<string>)(() => _mainWindow!.CurrentTheme),
            (Func<bool>)(() => _mainWindow!.AppWindow.IsVisible),
            (Func<bool>)(() => false),
            (Func<IReadOnlyList<VmDefinition>>)(() => Array.Empty<VmDefinition>()),
            (Func<IReadOnlyList<HyperVSwitchInfo>>)(() => Array.Empty<HyperVSwitchInfo>()),
            (Func<Task>)(async () => await RefreshUsbDevicesAsync()),
            (Action<EventHandler>)(_ => { }),
            (Action<EventHandler>)(_ => { }),
            (Func<string, Task>)(_ => Task.CompletedTask),
            (Func<string, Task>)(_ => Task.CompletedTask),
            (Func<string, Task>)(_ => Task.CompletedTask),
            (Func<string, Task>)(_ => Task.CompletedTask),
            (Func<string, Task>)(_ => Task.CompletedTask),
            (Func<string, string, Task>)((_, _) => Task.CompletedTask),
            (Func<string, Task>)(_ => Task.CompletedTask)
        };

        var parameterCount = initializeMethod.GetParameters().Length;
        if (parameterCount == 25)
        {
            var argsWithSelectedUsb = new List<object?>(commonArguments)
            {
                (Func<UsbIpDeviceInfo?>)GetSelectedUsbDevice,
                (Func<Task>)(async () => await RefreshUsbDevicesAsync()),
                (Func<Task>)(async () => await ConnectSelectedUsbFromTrayAsync()),
                (Func<Task>)(async () => await DisconnectSelectedUsbFromTrayAsync()),
                (Action)(() => _ = ExitWithAnimationAsync())
            };

            initializeMethod.Invoke(trayService, argsWithSelectedUsb.ToArray());
            return;
        }

        if (parameterCount == 24)
        {
            var argsWithoutSelectedUsb = new List<object?>(commonArguments)
            {
                (Func<Task>)(async () => await RefreshUsbDevicesAsync()),
                (Func<Task>)(async () => await ConnectSelectedUsbFromTrayAsync()),
                (Func<Task>)(async () => await DisconnectSelectedUsbFromTrayAsync()),
                (Action)(() => _ = ExitWithAnimationAsync())
            };

            initializeMethod.Invoke(trayService, argsWithoutSelectedUsb.ToArray());
            return;
        }

        throw new InvalidOperationException($"Unbekannte TrayService.Initialize-Signatur mit {parameterCount} Parametern.");
    }

    private void EnsureTrayControlCenterWindow()
    {
        if (_trayControlCenterWindow is not null)
        {
            return;
        }

        var window = new GuestTrayControlCenterWindow();
        _trayControlCenterWindow = window;

        window.CloseRequested += () =>
        {
            try
            {
                if (window.AppWindow.IsVisible)
                {
                    window.AppWindow.Hide();
                }
            }
            catch
            {
            }
        };

        window.RefreshUsbRequested += async () => await RefreshUsbDevicesAsync();
        window.UsbSelected += busId => _selectedUsbBusId = busId;
        window.UsbConnectRequested += async () => await ConnectSelectedUsbFromTrayAsync();
        window.UsbDisconnectRequested += async () => await DisconnectSelectedUsbFromTrayAsync();
        window.ToggleVisibilityRequested += () =>
        {
            if (_mainWindow is null)
            {
                return;
            }

            if (_mainWindow.AppWindow.IsVisible)
            {
                _mainWindow.AppWindow.Hide();
            }
            else
            {
                _mainWindow.AppWindow.Show();
                _mainWindow.Activate();
            }

            UpdateTrayControlCenterView();
        };
        window.ExitRequested += () => _ = ExitWithAnimationAsync();
    }

    private void ToggleTrayControlCenter()
    {
        if (_mainWindow is null)
        {
            return;
        }

        EnsureTrayControlCenterWindow();
        if (_trayControlCenterWindow is null)
        {
            return;
        }

        if (_trayControlCenterWindow.AppWindow.IsVisible)
        {
            _trayControlCenterWindow.AppWindow.Hide();
            return;
        }

        _ = RefreshUsbDevicesAsync();
        UpdateTrayControlCenterView();

        PositionTrayControlCenterNearTray();

        _trayControlCenterWindow.AppWindow.Show();
        _trayControlCenterWindow.Activate();
    }

    private void PositionTrayControlCenterNearTray()
    {
        if (_trayControlCenterWindow is null)
        {
            return;
        }

        var popupWidth = GuestTrayControlCenterWindow.PopupWidth;
        var popupHeight = _minimizeToTray
            ? GuestTrayControlCenterWindow.PopupHeightWithUsb
            : GuestTrayControlCenterWindow.PopupHeightCompact;

        if (!GetCursorPos(out var cursorPos))
        {
            cursorPos = new NativePoint { X = 0, Y = 0 };
        }

        var displayArea = DisplayArea.GetFromPoint(new PointInt32(cursorPos.X, cursorPos.Y), DisplayAreaFallback.Primary);
        var work = displayArea.WorkArea;
        var bounds = displayArea.OuterBounds;

        var x = cursorPos.X - popupWidth + 24;
        var y = work.Y + work.Height - popupHeight - 8;

        var taskbarAtBottom = work.Y + work.Height < bounds.Y + bounds.Height;
        var taskbarAtTop = work.Y > bounds.Y;
        var taskbarAtLeft = work.X > bounds.X;
        var taskbarAtRight = work.X + work.Width < bounds.X + bounds.Width;

        if (taskbarAtTop)
        {
            y = work.Y + 8;
        }
        else if (taskbarAtBottom)
        {
            y = work.Y + work.Height - popupHeight - 8;
        }

        if (taskbarAtLeft)
        {
            x = work.X + 8;
        }
        else if (taskbarAtRight)
        {
            x = work.X + work.Width - popupWidth - 8;
        }

        x = Math.Clamp(x, work.X + 8, work.X + work.Width - popupWidth - 8);
        y = Math.Clamp(y, work.Y + 8, work.Y + work.Height - popupHeight - 8);

        try
        {
            _trayControlCenterWindow.AppWindow.Move(new PointInt32(x, y));
        }
        catch
        {
        }
    }

    private void HideTrayControlCenter()
    {
        if (_trayControlCenterWindow?.AppWindow.IsVisible == true)
        {
            _trayControlCenterWindow.AppWindow.Hide();
        }
    }

    private void UpdateTrayControlCenterView()
    {
        if (_mainWindow?.DispatcherQueue is { } queue && !queue.HasThreadAccess)
        {
            _ = queue.TryEnqueue(UpdateTrayControlCenterView);
            return;
        }

        if (_trayControlCenterWindow is null || _mainWindow is null || _config is null)
        {
            return;
        }

        _trayControlCenterWindow.ApplyTheme(_mainWindow.CurrentTheme == "dark");
        _trayControlCenterWindow.UpdateView(_usbDevices, _selectedUsbBusId, _mainWindow.AppWindow.IsVisible, _minimizeToTray, _isUsbClientAvailable);
    }

    private UsbIpDeviceInfo? GetSelectedUsbDevice()
    {
        var selectedFromMain = _mainWindow?.GetSelectedUsbDevice();
        if (selectedFromMain is not null)
        {
            _selectedUsbBusId = selectedFromMain.BusId;
            return selectedFromMain;
        }

        if (string.IsNullOrWhiteSpace(_selectedUsbBusId))
        {
            return null;
        }

        return _usbDevices.FirstOrDefault(item => string.Equals(item.BusId, _selectedUsbBusId, StringComparison.OrdinalIgnoreCase));
    }

    private void UpdateUsbViews()
    {
        void apply()
        {
            _mainWindow?.UpdateUsbDevices(_usbDevices);
            UpdateTrayControlCenterView();
        }

        if (_mainWindow?.DispatcherQueue is { } queue && !queue.HasThreadAccess)
        {
            _ = queue.TryEnqueue(apply);
            return;
        }

        apply();
    }

    private Task<IReadOnlyList<UsbIpDeviceInfo>> RefreshUsbDevicesAsync()
    {
        return RefreshUsbDevicesAsync(emitLogs: true);
    }

    private async Task<IReadOnlyList<UsbIpDeviceInfo>> RefreshUsbDevicesAsync(bool emitLogs)
    {
        if (!await _usbRefreshGate.WaitAsync(0))
        {
            return _usbDevices;
        }

        try
        {
        await RefreshUsbClientAvailabilityAsync();

        if (_config?.Usb?.Enabled == false)
        {
            _usbDevices.Clear();
            UpdateUsbViews();

            if (emitLogs)
            {
                GuestLogger.Info("usb.refresh.blocked_guest_policy", "USB Refresh blockiert: Guest hat USB lokal deaktiviert.");
            }

            return _usbDevices;
        }

        if (_config?.Usb?.HostFeatureEnabled == false)
        {
            _usbDevices.Clear();
            UpdateUsbViews();

            if (emitLogs)
            {
                GuestLogger.Info("usb.refresh.blocked_host_policy", "USB Refresh blockiert: Host hat USB Share deaktiviert.");
            }

            return _usbDevices;
        }

        if (!_isUsbClientAvailable)
        {
            _usbDevices.Clear();
            UpdateUsbViews();

            if (!_usbClientMissingLogged)
            {
                _usbClientMissingLogged = true;
                GuestLogger.Warn("usb.client.missing", "USB/IP-Client nicht installiert. USB-Funktionen sind deaktiviert.", new
                {
                    source = "https://github.com/vadimgrn/usbip-win2"
                });
            }

            return _usbDevices;
        }

        var operationId = CreateUsbOperationId("refresh");
        var stopwatch = Stopwatch.StartNew();
        var hostResolution = ResolveUsbHostAddressDiagnostics();

        ApplyUsbTransportResolution(hostResolution);
        var previousCount = _usbDevices.Count;
        var fallbackToIpUsedForLog = false;

        try
        {
            IReadOnlyList<UsbIpDeviceInfo> list;
            var useRemoteHostList = !string.IsNullOrWhiteSpace(hostResolution.ResolvedIpv4);
            var fallbackToIpUsed = false;
            string? fallbackHostAddress = null;
            if (useRemoteHostList)
            {
                try
                {
                    list = await _usbService.GetRemoteDevicesAsync(hostResolution.ResolvedIpv4, CancellationToken.None);
                }
                catch when (string.Equals(hostResolution.Source, "hyperv-socket", StringComparison.OrdinalIgnoreCase))
                {
                    var ipFallbackResolution = ResolveUsbHostAddressDiagnostics(preferHyperVSocket: false);
                    if (string.IsNullOrWhiteSpace(ipFallbackResolution.ResolvedIpv4))
                    {
                        throw;
                    }

                    fallbackToIpUsed = true;
                    fallbackToIpUsedForLog = true;
                    fallbackHostAddress = ipFallbackResolution.ResolvedIpv4;
                    list = await _usbService.GetRemoteDevicesAsync(ipFallbackResolution.ResolvedIpv4, CancellationToken.None);
                }
            }
            else
            {
                list = await _usbService.GetDevicesAsync(CancellationToken.None);
            }

            var retryTriggered = false;
            if (useRemoteHostList && list.Count == 0 && previousCount > 0)
            {
                retryTriggered = true;
                await Task.Delay(900);
                list = await _usbService.GetRemoteDevicesAsync(hostResolution.ResolvedIpv4, CancellationToken.None);
            }

            var effectiveHostAddress = fallbackToIpUsed && !string.IsNullOrWhiteSpace(fallbackHostAddress)
                ? fallbackHostAddress
                : hostResolution.ResolvedIpv4;
            var effectiveRemoteList = !string.IsNullOrWhiteSpace(effectiveHostAddress);

            if (effectiveRemoteList)
            {
                await CleanupDanglingGuestAttachmentsAsync(list, effectiveHostAddress!);
            }

            var autoConnectApplied = await ApplyUsbAutoConnectAsync(list);
            if (autoConnectApplied > 0)
            {
                await Task.Delay(260);

                if (effectiveRemoteList)
                {
                    list = await _usbService.GetRemoteDevicesAsync(effectiveHostAddress!, CancellationToken.None);
                }
                else
                {
                    list = await _usbService.GetDevicesAsync(CancellationToken.None);
                }
            }

            _usbDevices.Clear();
            _usbDevices.AddRange(list);

            ApplyUsbTransportResolution(hostResolution, fallbackToIpUsed);

            UpdateUsbViews();

            if (emitLogs)
            {
                GuestLogger.Info("usb.refresh.success", "USB-Geräte aktualisiert.", new
                {
                    operationId,
                    elapsedMs = stopwatch.ElapsedMilliseconds,
                    count = _usbDevices.Count,
                    selectedBusId = _selectedUsbBusId,
                    attachedCount = _usbDevices.Count(item => item.IsAttached),
                    sharedCount = _usbDevices.Count(item => item.IsShared),
                    autoConnectApplied,
                    retryTriggered,
                    previousCount,
                    listSource = useRemoteHostList ? "remote-host" : "local-state",
                    hostAddress = hostResolution.ResolvedIpv4,
                    hostSource = hostResolution.Source,
                    hostInput = hostResolution.RawInput,
                    transportPath = fallbackToIpUsed
                        ? "ip-fallback"
                        : (string.Equals(hostResolution.Source, "hyperv-socket", StringComparison.OrdinalIgnoreCase) ? "hyperv" : "ip-fallback"),
                    fallbackToIpUsed,
                    fallbackHostAddress
                });
            }

            return _usbDevices;
        }
        catch (Exception ex)
        {
            if (emitLogs)
            {
                GuestLogger.Warn("usb.refresh.failed", ex.Message, new
                {
                    operationId,
                    elapsedMs = stopwatch.ElapsedMilliseconds,
                    selectedBusId = _selectedUsbBusId,
                    exceptionType = ex.GetType().FullName,
                    hResult = ex.HResult,
                    hostAddress = hostResolution.ResolvedIpv4,
                    hostSource = hostResolution.Source,
                    hostInput = hostResolution.RawInput,
                    transportPath = fallbackToIpUsedForLog
                        ? "ip-fallback"
                        : (string.Equals(hostResolution.Source, "hyperv-socket", StringComparison.OrdinalIgnoreCase) ? "hyperv" : "ip-fallback"),
                    hostResolutionFailure = hostResolution.FailureReason
                });
            }
            return _usbDevices;
        }
        }
        finally
        {
            _usbRefreshGate.Release();
        }
    }

    private async Task CleanupDanglingGuestAttachmentsAsync(IReadOnlyList<UsbIpDeviceInfo> remoteDevices, string hostAddress)
    {
        if (remoteDevices.Count == 0 && _usbDevices.Count == 0)
        {
            return;
        }

        var remoteBusIds = new HashSet<string>(
            remoteDevices
                .Where(device => !string.IsNullOrWhiteSpace(device.BusId))
                .Select(device => device.BusId.Trim()),
            StringComparer.OrdinalIgnoreCase);

        var danglingAttachedBusIds = _usbDevices
            .Where(device =>
                device.IsAttached
                && !string.IsNullOrWhiteSpace(device.BusId)
                && !remoteBusIds.Contains(device.BusId.Trim()))
            .Select(device => device.BusId.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (danglingAttachedBusIds.Count == 0)
        {
            return;
        }

        foreach (var busId in danglingAttachedBusIds)
        {
            try
            {
                await _usbService.DetachFromHostAsync(busId, hostAddress, CancellationToken.None);
                await TrySendUsbConnectionEventAckAsync(busId, "usb-disconnected", CancellationToken.None);

                GuestLogger.Info("usb.cleanup.dangling_detach", "Dangling USB-Anhang im Guest wurde hart getrennt.", new
                {
                    busId,
                    hostAddress,
                    reason = "host-device-missing"
                });
            }
            catch (Exception ex)
            {
                GuestLogger.Warn("usb.cleanup.dangling_detach_failed", ex.Message, new
                {
                    busId,
                    hostAddress,
                    exceptionType = ex.GetType().FullName
                });
            }
        }
    }

    private async Task<int> ConnectUsbAsync(string busId)
    {
        if (_config?.Usb?.HostFeatureEnabled == false)
        {
            GuestLogger.Warn("usb.connect.blocked_host_policy", "USB Connect blockiert: Host hat USB Share deaktiviert.", new
            {
                busId
            });
            return 1;
        }

        if (!_isUsbClientAvailable)
        {
            GuestLogger.Warn("usb.connect.blocked", "USB/IP-Client nicht installiert. Connect ist deaktiviert.", new
            {
                busId,
                source = "https://github.com/vadimgrn/usbip-win2"
            });
            return 1;
        }

        if (string.IsNullOrWhiteSpace(busId))
        {
            GuestLogger.Warn("usb.connect.invalid", "USB Connect mit leerer BUSID angefordert.");
            return 1;
        }

        var operationId = CreateUsbOperationId("connect");
        var stopwatch = Stopwatch.StartNew();
        var hostResolution = ResolveUsbHostAddressDiagnostics();
        ApplyUsbTransportResolution(hostResolution);
        var beforeDevice = FindUsbByBusId(busId);

        GuestLogger.Info("usb.connect.begin", "USB Host-Attach gestartet.", new
        {
            operationId,
            busId,
            selectedBusId = _selectedUsbBusId,
            hostAddress = hostResolution.ResolvedIpv4,
            hostSource = hostResolution.Source,
            hostInput = hostResolution.RawInput,
            sharePath = _config?.SharePath,
            beforeState = beforeDevice?.StateText,
            beforeClientIp = beforeDevice?.ClientIpAddress,
            beforePersistedGuid = beforeDevice?.PersistedGuid
        });

        if (string.IsNullOrWhiteSpace(hostResolution.ResolvedIpv4))
        {
            GuestLogger.Warn("usb.connect.failed", "USB-Host-Adresse fehlt oder konnte nicht auf IPv4 aufgelöst werden. Bitte usb.hostAddress oder SharePath konfigurieren.", new
            {
                operationId,
                busId,
                elapsedMs = stopwatch.ElapsedMilliseconds,
                hostSource = hostResolution.Source,
                hostInput = hostResolution.RawInput,
                transportPath = string.Equals(hostResolution.Source, "hyperv-socket", StringComparison.OrdinalIgnoreCase)
                    ? "hyperv"
                    : "ip-fallback",
                hostResolutionFailure = hostResolution.FailureReason,
                sharePath = _config?.SharePath
            });
            return 1;
        }

        try
        {
            await _usbService.AttachToHostAsync(busId, hostResolution.ResolvedIpv4, CancellationToken.None);
            _selectedUsbBusId = busId;
            ApplyUsbTransportResolution(hostResolution);

            GuestLogger.Info("usb.connect.success", "USB Host-Attach erfolgreich.", new
            {
                operationId,
                busId,
                elapsedMs = stopwatch.ElapsedMilliseconds,
                hostAddress = hostResolution.ResolvedIpv4,
                hostSource = hostResolution.Source,
                transportPath = string.Equals(hostResolution.Source, "hyperv-socket", StringComparison.OrdinalIgnoreCase)
                    ? "hyperv"
                    : "ip-fallback"
            });

                    await TrySendUsbConnectionEventAckAsync(busId, "usb-connected", CancellationToken.None);

            return 0;
        }
        catch (Exception ex)
        {
            if (ShouldRetryUsbAttach(ex))
            {
                try
                {
                    GuestLogger.Warn("usb.connect.retry", "USB Attach wird nach Recovery erneut versucht.", new
                    {
                        operationId,
                        busId,
                        elapsedMs = stopwatch.ElapsedMilliseconds,
                        hostAddress = hostResolution.ResolvedIpv4,
                        hostSource = hostResolution.Source,
                        reason = ex.Message,
                        transportPath = string.Equals(hostResolution.Source, "hyperv-socket", StringComparison.OrdinalIgnoreCase)
                            ? "hyperv"
                            : "ip-fallback"
                    });

                    await TryRecoverAndRetryUsbAttachAsync(busId, hostResolution.ResolvedIpv4, CancellationToken.None);
                    _selectedUsbBusId = busId;
                    ApplyUsbTransportResolution(hostResolution);

                    GuestLogger.Info("usb.connect.success", "USB Host-Attach erfolgreich (Retry).", new
                    {
                        operationId,
                        busId,
                        elapsedMs = stopwatch.ElapsedMilliseconds,
                        hostAddress = hostResolution.ResolvedIpv4,
                        hostSource = hostResolution.Source,
                        transportPath = string.Equals(hostResolution.Source, "hyperv-socket", StringComparison.OrdinalIgnoreCase)
                            ? "hyperv"
                            : "ip-fallback",
                        recovery = "retry-after-detach"
                    });

                    await TrySendUsbConnectionEventAckAsync(busId, "usb-connected", CancellationToken.None);
                    return 0;
                }
                catch (Exception retryEx)
                {
                    GuestLogger.Warn("usb.connect.retry.failed", retryEx.Message, new
                    {
                        operationId,
                        busId,
                        elapsedMs = stopwatch.ElapsedMilliseconds,
                        hostAddress = hostResolution.ResolvedIpv4,
                        hostSource = hostResolution.Source,
                        transportPath = string.Equals(hostResolution.Source, "hyperv-socket", StringComparison.OrdinalIgnoreCase)
                            ? "hyperv"
                            : "ip-fallback",
                        exceptionType = retryEx.GetType().FullName
                    });
                }
            }

            if (string.Equals(hostResolution.Source, "hyperv-socket", StringComparison.OrdinalIgnoreCase))
            {
                var ipFallbackResolution = ResolveUsbHostAddressDiagnostics(preferHyperVSocket: false);
                if (!string.IsNullOrWhiteSpace(ipFallbackResolution.ResolvedIpv4))
                {
                    try
                    {
                        await _usbService.AttachToHostAsync(busId, ipFallbackResolution.ResolvedIpv4, CancellationToken.None);
                        _selectedUsbBusId = busId;
                        ApplyUsbTransportResolution(ipFallbackResolution, fallbackToIp: true);

                        GuestLogger.Info("usb.connect.success", "USB Host-Attach erfolgreich (IP-Fallback).", new
                        {
                            operationId,
                            busId,
                            elapsedMs = stopwatch.ElapsedMilliseconds,
                            hostAddress = ipFallbackResolution.ResolvedIpv4,
                            hostSource = ipFallbackResolution.Source,
                            transportPath = "ip-fallback",
                            fallback = "ip"
                        });

                        await TrySendUsbConnectionEventAckAsync(busId, "usb-connected", CancellationToken.None);

                        return 0;
                    }
                    catch (Exception fallbackEx)
                    {
                        GuestLogger.Warn("usb.connect.fallback.failed", fallbackEx.Message, new
                        {
                            operationId,
                            busId,
                            elapsedMs = stopwatch.ElapsedMilliseconds,
                            fallbackHostAddress = ipFallbackResolution.ResolvedIpv4,
                            fallbackHostSource = ipFallbackResolution.Source,
                            transportPath = "ip-fallback",
                            exceptionType = fallbackEx.GetType().FullName
                        });
                    }
                }
            }

            GuestLogger.Warn("usb.connect.failed", ex.Message, new
            {
                operationId,
                busId,
                elapsedMs = stopwatch.ElapsedMilliseconds,
                hostAddress = hostResolution.ResolvedIpv4,
                hostSource = hostResolution.Source,
                transportPath = string.Equals(hostResolution.Source, "hyperv-socket", StringComparison.OrdinalIgnoreCase)
                    ? "hyperv"
                    : "ip-fallback",
                exceptionType = ex.GetType().FullName
            });
            return 1;
        }
    }

    private async Task TryRecoverAndRetryUsbAttachAsync(string busId, string hostAddress, CancellationToken cancellationToken)
    {
        try
        {
            await _usbService.DetachFromHostAsync(busId, hostAddress, cancellationToken);
        }
        catch
        {
        }

        await Task.Delay(TimeSpan.FromMilliseconds(850), cancellationToken);
        await _usbService.AttachToHostAsync(busId, hostAddress, cancellationToken);
    }

    private static bool ShouldRetryUsbAttach(Exception exception)
    {
        var message = exception.Message ?? string.Empty;
        return message.Contains("device in error state", StringComparison.OrdinalIgnoreCase)
               || message.Contains("error state", StringComparison.OrdinalIgnoreCase)
               || message.Contains("resource busy", StringComparison.OrdinalIgnoreCase)
               || message.Contains("temporarily unavailable", StringComparison.OrdinalIgnoreCase)
               || message.Contains("already in use", StringComparison.OrdinalIgnoreCase);
    }

    private async Task<string?> DiscoverUsbHostAddressAsync()
    {
        var hostIdentity = await FetchHostIdentityViaHyperVSocketAsync(CancellationToken.None);
        if (hostIdentity is not null)
        {
            ApplyHostIdentityToConfig(hostIdentity, persistConfig: true);

            var discoveredHostName = (hostIdentity.HostName ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(discoveredHostName))
            {
                GuestLogger.Info("usb.host.discovery.hostname", "Hostname per Hyper-V Socket gefunden.", new
                {
                    requester = Environment.MachineName,
                    discoveredHostName
                });

                return discoveredHostName;
            }
        }

        try
        {
            var discoveryResult = await UsbHostDiscoveryService.DiscoverHostAsync(
                requesterComputerName: Environment.MachineName,
                timeout: TimeSpan.FromSeconds(3),
                cancellationToken: CancellationToken.None);

            var discovered = (discoveryResult?.HostAddress ?? string.Empty).Trim();
            var discoveredBroadcastHostName = (discoveryResult?.HostComputerName ?? string.Empty).Trim();

            if (!string.IsNullOrWhiteSpace(discovered) || !string.IsNullOrWhiteSpace(discoveredBroadcastHostName))
            {
                _config ??= new GuestConfig();
                _config.Usb ??= new GuestUsbSettings();

                if (!string.IsNullOrWhiteSpace(discoveredBroadcastHostName))
                {
                    _config.Usb.HostName = discoveredBroadcastHostName;
                    _config.Usb.HostAddress = discoveredBroadcastHostName;
                }
                else
                {
                    _config.Usb.HostAddress = discovered;
                }

                GuestLogger.Info("usb.host.discovery.success", "Host per Broadcast gefunden.", new
                {
                    requester = Environment.MachineName,
                    discoveredHostName = discoveredBroadcastHostName,
                    discoveredAddress = discovered,
                    transportPath = !string.IsNullOrWhiteSpace(discoveredBroadcastHostName)
                        ? "hostname-broadcast"
                        : "ip-broadcast"
                });

                GuestConfigService.Save(_configPath, _config);

                if (!string.IsNullOrWhiteSpace(discoveredBroadcastHostName))
                {
                    return discoveredBroadcastHostName;
                }
            }
            else
            {
                GuestLogger.Warn("usb.host.discovery.empty", "Keine Host-Adresse per Broadcast gefunden.", new
                {
                    requester = Environment.MachineName
                });
            }

            return discovered;
        }
        catch (Exception ex)
        {
            GuestLogger.Warn("usb.host.discovery.failed", ex.Message, new
            {
                requester = Environment.MachineName,
                exceptionType = ex.GetType().FullName
            });
            return null;
        }
    }

    private async Task<HostIdentityInfo?> FetchHostIdentityViaHyperVSocketAsync(CancellationToken cancellationToken)
    {
        try
        {
            var client = new HyperVSocketHostIdentityGuestClient();
            return await client.FetchHostIdentityAsync(cancellationToken);
        }
        catch
        {
            return null;
        }
    }

    private async Task<int> DisconnectUsbAsync(string busId)
    {
        if (_config?.Usb?.HostFeatureEnabled == false)
        {
            GuestLogger.Warn("usb.disconnect.blocked_host_policy", "USB Disconnect blockiert: Host hat USB Share deaktiviert.", new
            {
                busId
            });
            return 1;
        }

        if (!_isUsbClientAvailable)
        {
            GuestLogger.Warn("usb.disconnect.blocked", "USB/IP-Client nicht installiert. Disconnect ist deaktiviert.", new
            {
                busId,
                source = "https://github.com/vadimgrn/usbip-win2"
            });
            return 1;
        }

        if (string.IsNullOrWhiteSpace(busId))
        {
            GuestLogger.Warn("usb.disconnect.invalid", "USB Disconnect mit leerer BUSID angefordert.");
            return 1;
        }

        var operationId = CreateUsbOperationId("disconnect");
        var stopwatch = Stopwatch.StartNew();
        var hostResolution = ResolveUsbHostAddressDiagnostics();
        ApplyUsbTransportResolution(hostResolution);
        var beforeDevice = FindUsbByBusId(busId);

        GuestLogger.Info("usb.disconnect.begin", "USB Host-Detach gestartet.", new
        {
            operationId,
            busId,
            selectedBusId = _selectedUsbBusId,
            hostAddress = hostResolution.ResolvedIpv4,
            hostSource = hostResolution.Source,
            hostInput = hostResolution.RawInput,
            beforeState = beforeDevice?.StateText,
            beforeClientIp = beforeDevice?.ClientIpAddress,
            beforePersistedGuid = beforeDevice?.PersistedGuid
        });

        try
        {
            await _usbService.DetachFromHostAsync(busId, hostResolution.ResolvedIpv4, CancellationToken.None);
            _selectedUsbBusId = busId;
            ApplyUsbTransportResolution(hostResolution);

            GuestLogger.Info("usb.disconnect.success", "USB Host-Detach erfolgreich.", new
            {
                operationId,
                busId,
                elapsedMs = stopwatch.ElapsedMilliseconds,
                hostAddress = hostResolution.ResolvedIpv4,
                hostSource = hostResolution.Source,
                transportPath = string.Equals(hostResolution.Source, "hyperv-socket", StringComparison.OrdinalIgnoreCase)
                    ? "hyperv"
                    : "ip-fallback"
            });

                    await TrySendUsbConnectionEventAckAsync(busId, "usb-disconnected", CancellationToken.None);

            return 0;
        }
        catch (Exception ex)
        {
            if (string.Equals(hostResolution.Source, "hyperv-socket", StringComparison.OrdinalIgnoreCase))
            {
                var ipFallbackResolution = ResolveUsbHostAddressDiagnostics(preferHyperVSocket: false);
                if (!string.IsNullOrWhiteSpace(ipFallbackResolution.ResolvedIpv4))
                {
                    try
                    {
                        await _usbService.DetachFromHostAsync(busId, ipFallbackResolution.ResolvedIpv4, CancellationToken.None);
                        _selectedUsbBusId = busId;
                        ApplyUsbTransportResolution(ipFallbackResolution, fallbackToIp: true);

                        GuestLogger.Info("usb.disconnect.success", "USB Host-Detach erfolgreich (IP-Fallback).", new
                        {
                            operationId,
                            busId,
                            elapsedMs = stopwatch.ElapsedMilliseconds,
                            hostAddress = ipFallbackResolution.ResolvedIpv4,
                            hostSource = ipFallbackResolution.Source,
                            transportPath = "ip-fallback",
                            fallback = "ip"
                        });

                        await TrySendUsbConnectionEventAckAsync(busId, "usb-disconnected", CancellationToken.None);

                        return 0;
                    }
                    catch (Exception fallbackEx)
                    {
                        GuestLogger.Warn("usb.disconnect.fallback.failed", fallbackEx.Message, new
                        {
                            operationId,
                            busId,
                            elapsedMs = stopwatch.ElapsedMilliseconds,
                            fallbackHostAddress = ipFallbackResolution.ResolvedIpv4,
                            fallbackHostSource = ipFallbackResolution.Source,
                            transportPath = "ip-fallback",
                            exceptionType = fallbackEx.GetType().FullName
                        });
                    }
                }
            }

            GuestLogger.Warn("usb.disconnect.failed", ex.Message, new
            {
                operationId,
                busId,
                elapsedMs = stopwatch.ElapsedMilliseconds,
                hostAddress = hostResolution.ResolvedIpv4,
                hostSource = hostResolution.Source,
                transportPath = string.Equals(hostResolution.Source, "hyperv-socket", StringComparison.OrdinalIgnoreCase)
                    ? "hyperv"
                    : "ip-fallback",
                exceptionType = ex.GetType().FullName
            });
            return 1;
        }
    }

    private UsbHostResolution ResolveUsbHostAddressDiagnostics(bool preferHyperVSocket = true)
    {
        if (preferHyperVSocket
            && _usbHyperVSocketProxy?.IsRunning == true
            && _usbHyperVSocketServiceReachable)
        {
            return new UsbHostResolution(
                HyperVSocketUsbTunnelDefaults.LoopbackAddress,
                "hyperv-socket",
                "AF_HYPERV",
                string.Empty);
        }

        var configured = (_config?.Usb?.HostAddress ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(configured))
        {
            var resolvedConfigured = ResolveToIpv4(configured);
            if (!string.IsNullOrWhiteSpace(resolvedConfigured))
            {
                return new UsbHostResolution(resolvedConfigured, "usb.hostAddress", configured, string.Empty);
            }

            return new UsbHostResolution(string.Empty, "usb.hostAddress", configured, "Configured host konnte nicht in IPv4 aufgelöst werden.");
        }

        var sharePath = (_config?.SharePath ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(sharePath))
        {
            var defaultGatewayIpv4 = ResolveDefaultGatewayIpv4();
            if (!string.IsNullOrWhiteSpace(defaultGatewayIpv4))
            {
                return new UsbHostResolution(defaultGatewayIpv4, "default-gateway", "network", string.Empty);
            }

            return new UsbHostResolution(string.Empty, "sharePath", string.Empty, "SharePath ist leer.");
        }

        if (Uri.TryCreate(sharePath, UriKind.Absolute, out var shareUri) && shareUri.IsUnc)
        {
            var host = shareUri.Host;
            var resolvedFromShare = ResolveToIpv4(host);
            if (!string.IsNullOrWhiteSpace(resolvedFromShare))
            {
                return new UsbHostResolution(resolvedFromShare, "sharePath", host, string.Empty);
            }

            var defaultGatewayIpv4 = ResolveDefaultGatewayIpv4();
            if (!string.IsNullOrWhiteSpace(defaultGatewayIpv4))
            {
                return new UsbHostResolution(defaultGatewayIpv4, "default-gateway", host, string.Empty);
            }

            return new UsbHostResolution(string.Empty, "sharePath", host, "UNC-Host konnte nicht in IPv4 aufgelöst werden.");
        }

        var fallbackGatewayIpv4 = ResolveDefaultGatewayIpv4();
        if (!string.IsNullOrWhiteSpace(fallbackGatewayIpv4))
        {
            return new UsbHostResolution(fallbackGatewayIpv4, "default-gateway", sharePath, string.Empty);
        }

        return new UsbHostResolution(string.Empty, "sharePath", sharePath, "SharePath ist kein gültiger UNC-Pfad.");
    }

    private async Task RefreshUsbClientAvailabilityAsync()
    {
        try
        {
            _isUsbClientAvailable = await _usbService.IsUsbClientAvailableAsync(CancellationToken.None);
            if (_isUsbClientAvailable)
            {
                _usbClientMissingLogged = false;
            }
        }
        catch
        {
            _isUsbClientAvailable = false;
        }

        if (_mainWindow?.DispatcherQueue is { } queue && !queue.HasThreadAccess)
        {
            _ = queue.TryEnqueue(() => _mainWindow?.UpdateUsbClientAvailability(_isUsbClientAvailable));
        }
        else
        {
            _mainWindow?.UpdateUsbClientAvailability(_isUsbClientAvailable);
        }

        UpdateTrayControlCenterView();
    }

    private static string ResolveToIpv4(string value)
    {
        var candidate = (value ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(candidate))
        {
            return string.Empty;
        }

        if (IPAddress.TryParse(candidate, out var parsedIp) && parsedIp.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
        {
            return parsedIp.ToString();
        }

        try
        {
            var addresses = Dns.GetHostAddresses(candidate);
            var ipv4 = addresses.FirstOrDefault(address => address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);
            return ipv4?.ToString() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string ResolveDefaultGatewayIpv4()
    {
        try
        {
            foreach (var networkInterface in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (networkInterface.OperationalStatus != OperationalStatus.Up)
                {
                    continue;
                }

                if (networkInterface.NetworkInterfaceType is NetworkInterfaceType.Loopback or NetworkInterfaceType.Tunnel)
                {
                    continue;
                }

                var properties = networkInterface.GetIPProperties();
                var gateway = properties.GatewayAddresses
                    .Select(item => item.Address)
                    .FirstOrDefault(address =>
                        address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork
                        && !IPAddress.Any.Equals(address)
                        && !IPAddress.None.Equals(address));

                if (gateway is not null)
                {
                    return gateway.ToString();
                }
            }
        }
        catch
        {
        }

        return string.Empty;
    }

    private UsbIpDeviceInfo? FindUsbByBusId(string busId)
    {
        if (string.IsNullOrWhiteSpace(busId))
        {
            return null;
        }

        return _usbDevices.FirstOrDefault(item => string.Equals(item.BusId, busId, StringComparison.OrdinalIgnoreCase));
    }

    private static string CreateUsbOperationId(string action)
    {
        return $"usb-{action}-{Guid.NewGuid():N}";
    }

    private sealed record UsbHostResolution(string ResolvedIpv4, string Source, string RawInput, string FailureReason);

    private async Task ConnectSelectedUsbFromTrayAsync()
    {
        var selected = GetSelectedUsbDevice();
        if (selected is null)
        {
            return;
        }

        await ConnectUsbAsync(selected.BusId);
        await RefreshUsbDevicesAsync();
    }

    private async Task DisconnectSelectedUsbFromTrayAsync()
    {
        var selected = GetSelectedUsbDevice();
        if (selected is null)
        {
            return;
        }

        var result = await DisconnectUsbAsync(selected.BusId);
        if (result == 0)
        {
            await Task.Delay(TimeSpan.FromSeconds(3));
        }

        await RefreshUsbDevicesAsync();
    }

    private async Task DisconnectAllAttachedUsbOnExitAsync()
    {
        try
        {
            await RefreshUsbDevicesAsync();
        }
        catch (Exception ex)
        {
            GuestLogger.Warn("usb.exit_refresh.failed", ex.Message, new
            {
                exceptionType = ex.GetType().FullName
            });
        }

        var attachedBusIds = _usbDevices
            .Where(device => device.IsAttached && !string.IsNullOrWhiteSpace(device.BusId))
            .Select(device => device.BusId.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (attachedBusIds.Count == 0)
        {
            return;
        }

        GuestLogger.Info("usb.exit_disconnect.begin", "Guest-App beendet, trenne verbundene USB-Geräte.", new
        {
            count = attachedBusIds.Count,
            busIds = attachedBusIds
        });

        var disconnectedCount = 0;
        var failedCount = 0;

        foreach (var busId in attachedBusIds)
        {
            try
            {
                var result = await DisconnectUsbAsync(busId);
                if (result == 0)
                {
                    disconnectedCount++;
                }
                else
                {
                    failedCount++;
                }
            }
            catch (Exception ex)
            {
                failedCount++;
                GuestLogger.Warn("usb.exit_disconnect.item_failed", ex.Message, new
                {
                    busId,
                    exceptionType = ex.GetType().FullName
                });
            }
        }

        GuestLogger.Info("usb.exit_disconnect.done", "USB-Disconnect beim Beenden abgeschlossen.", new
        {
            requestedCount = attachedBusIds.Count,
            disconnectedCount,
            failedCount
        });
    }

    private async Task DisconnectAllAttachedUsbForTransportSwitchAsync()
    {
        try
        {
            await RefreshUsbDevicesAsync(emitLogs: false);
        }
        catch
        {
        }

        var attachedBusIds = _usbDevices
            .Where(device => device.IsAttached && !string.IsNullOrWhiteSpace(device.BusId))
            .Select(device => device.BusId.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (attachedBusIds.Count == 0)
        {
            return;
        }

        var hostResolution = ResolveUsbHostAddressDiagnostics();
        var disconnectedCount = 0;
        var failedCount = 0;

        foreach (var busId in attachedBusIds)
        {
            try
            {
                await _usbService.DetachFromHostAsync(busId, hostResolution.ResolvedIpv4, CancellationToken.None);
                await TrySendUsbConnectionEventAckAsync(busId, "usb-disconnected", CancellationToken.None);
                disconnectedCount++;
                continue;
            }
            catch
            {
            }

            if (string.Equals(hostResolution.Source, "hyperv-socket", StringComparison.OrdinalIgnoreCase))
            {
                var ipFallbackResolution = ResolveUsbHostAddressDiagnostics(preferHyperVSocket: false);
                if (!string.IsNullOrWhiteSpace(ipFallbackResolution.ResolvedIpv4))
                {
                    try
                    {
                        await _usbService.DetachFromHostAsync(busId, ipFallbackResolution.ResolvedIpv4, CancellationToken.None);
                        await TrySendUsbConnectionEventAckAsync(busId, "usb-disconnected", CancellationToken.None);
                        disconnectedCount++;
                        continue;
                    }
                    catch
                    {
                    }
                }
            }

            failedCount++;
        }

        try
        {
            await RefreshUsbDevicesAsync(emitLogs: false);
        }
        catch
        {
        }

        GuestLogger.Info("usb.transport.switch.disconnect", "USB-Geräte vor Transportwechsel getrennt.", new
        {
            requestedCount = attachedBusIds.Count,
            disconnectedCount,
            failedCount
        });
    }

    private void StartUsbAutoRefreshLoop()
    {
        StopUsbAutoRefreshLoop();

        var cts = new CancellationTokenSource();
        _usbAutoRefreshCts = cts;
        _usbAutoRefreshTask = Task.Run(() => RunUsbAutoRefreshLoopAsync(cts.Token));
    }

    private void StopUsbAutoRefreshLoop()
    {
        try
        {
            _usbAutoRefreshCts?.Cancel();
        }
        catch
        {
        }

        _usbAutoRefreshCts?.Dispose();
        _usbAutoRefreshCts = null;
        _usbAutoRefreshTask = null;
    }

    private async Task RunUsbAutoRefreshLoopAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                if (!_isExitRequested && _config?.Usb?.Enabled != false)
                {
                    var wasHostUsbEnabled = _config?.Usb?.HostFeatureEnabled != false;

                    try
                    {
                        var identity = await FetchHostIdentityViaHyperVSocketAsync(cancellationToken);
                        if (identity is not null)
                        {
                            ApplyHostIdentityToConfig(identity, persistConfig: false);
                        }
                    }
                    catch
                    {
                    }

                    var isHostUsbEnabled = _config?.Usb?.HostFeatureEnabled != false;
                    if (!wasHostUsbEnabled && isHostUsbEnabled && _isUsbClientAvailable)
                    {
                        await RefreshUsbDevicesAsync(emitLogs: false);
                    }
                }

                if (!_isExitRequested && _isUsbClientAvailable && _config?.Usb?.Enabled != false && HasConfiguredUsbAutoConnect())
                {
                    await RefreshUsbDevicesAsync(emitLogs: false);
                }
            }
            catch
            {
            }

            try
            {
                await Task.Delay(TimeSpan.FromSeconds(GuestUsbAutoRefreshSeconds), cancellationToken);
            }
            catch (OperationCanceledException)
            {
                break;
            }
        }
    }

    private void StartSharedFolderAutoMountLoop()
    {
        StopSharedFolderAutoMountLoop();

        var cts = new CancellationTokenSource();
        _sharedFolderAutoMountCts = cts;
        _sharedFolderAutoMountTask = Task.Run(() => RunSharedFolderAutoMountLoopAsync(cts.Token));
        UpdateSharedFolderReconnectStatusPanel();
    }

    private void StopSharedFolderAutoMountLoop()
    {
        try
        {
            _sharedFolderAutoMountCts?.Cancel();
        }
        catch
        {
        }

        _sharedFolderAutoMountCts?.Dispose();
        _sharedFolderAutoMountCts = null;
        _sharedFolderAutoMountTask = null;
        UpdateSharedFolderReconnectStatusPanel();
    }

    private void TriggerSharedFolderReconnectCycle()
    {
        _ = Task.Run(async () =>
        {
            try
            {
                if (!_isExitRequested)
                {
                    await RefreshSharedFolderHostNameFromSocketAsync(CancellationToken.None);
                }

                if (!_isExitRequested && HasConfiguredSharedFolderAutoMount())
                {
                    var result = await EnsureEnabledSharedFolderMappingsAsync(CancellationToken.None);
                    _sharedFolderReconnectLastRunUtc = DateTimeOffset.UtcNow;
                    _sharedFolderReconnectLastSummary = BuildSharedFolderReconnectSummary(result);
                }
                else
                {
                    _sharedFolderReconnectLastSummary = "Keine aktiven Mappings konfiguriert.";
                }
            }
            catch
            {
            }
            finally
            {
                UpdateSharedFolderReconnectStatusPanel();
            }
        });
    }

    private async Task RunSharedFolderAutoMountLoopAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                if (!_isExitRequested)
                {
                    await RefreshSharedFolderHostNameFromSocketAsync(cancellationToken);
                }

                if (!_isExitRequested && HasConfiguredSharedFolderAutoMount())
                {
                    var result = await EnsureEnabledSharedFolderMappingsAsync(cancellationToken);
                    _sharedFolderReconnectLastRunUtc = DateTimeOffset.UtcNow;
                    _sharedFolderReconnectLastSummary = BuildSharedFolderReconnectSummary(result);
                    UpdateSharedFolderReconnectStatusPanel();
                }
                else
                {
                    _sharedFolderReconnectLastSummary = "Keine aktiven Mappings konfiguriert.";
                    UpdateSharedFolderReconnectStatusPanel();
                }
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch
            {
            }

            try
            {
                await Task.Delay(TimeSpan.FromSeconds(GetSharedFolderReconnectIntervalSeconds()), cancellationToken);
            }
            catch (OperationCanceledException)
            {
                break;
            }
        }
    }

    private bool HasConfiguredSharedFolderAutoMount()
    {
        if (_config?.SharedFolders is null || !_config.SharedFolders.Enabled || _config.SharedFolders.HostFeatureEnabled == false)
        {
            return false;
        }

        var mappings = _config?.SharedFolders?.Mappings;
        return mappings is not null
               && mappings.Any(mapping => mapping.Enabled
                                           && !string.IsNullOrWhiteSpace(mapping.SharePath)
                                           && !string.IsNullOrWhiteSpace(mapping.DriveLetter));
    }

    private int GetSharedFolderReconnectIntervalSeconds()
    {
        return Math.Clamp(_config?.PollIntervalSeconds ?? 15, 5, 3600);
    }

    private static string BuildSharedFolderReconnectSummary(SharedFolderReconnectCycleResult result)
    {
        return $"Lauf: {result.Attempted} geprüft, {result.NewlyMounted} neu verbunden, {result.Failed} Fehler";
    }

    private void UpdateSharedFolderReconnectStatusPanel()
    {
        var reconnectActive = !_isExitRequested && HasConfiguredSharedFolderAutoMount();
        var lastRunUtc = _sharedFolderReconnectLastRunUtc;
        var summary = _sharedFolderReconnectLastSummary;

        void apply()
        {
            _mainWindow?.UpdateSharedFolderReconnectStatus(reconnectActive, lastRunUtc, summary);
        }

        if (_mainWindow?.DispatcherQueue is { } queue && !queue.HasThreadAccess)
        {
            _ = queue.TryEnqueue(apply);
            return;
        }

        apply();
    }

    private async Task<SharedFolderReconnectCycleResult> EnsureEnabledSharedFolderMappingsAsync(CancellationToken cancellationToken)
    {
        await _sharedFolderReconnectGate.WaitAsync(cancellationToken);
        try
        {
            var attempted = 0;
            var newlyMounted = 0;
            var failed = 0;

            if (_config?.SharedFolders is null || !_config.SharedFolders.Enabled)
            {
                var legacyLetters = _config?.SharedFolders?.Mappings?
                    .Select(mapping => GuestConfigService.NormalizeDriveLetter(mapping.DriveLetter))
                    .Where(letter => !string.IsNullOrWhiteSpace(letter))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList()
                    ?? [];

                if (!string.IsNullOrWhiteSpace(_config?.SharedFolders?.BaseDriveLetter))
                {
                    legacyLetters.Add(GuestConfigService.NormalizeDriveLetter(_config.SharedFolders.BaseDriveLetter));
                }

                await _winFspMountService.UnmountManyAsync(legacyLetters, cancellationToken);
                return new SharedFolderReconnectCycleResult();
            }

            var mappings = _config?.SharedFolders?.Mappings?
                .Where(mapping => mapping.Enabled
                                  && !string.IsNullOrWhiteSpace(mapping.SharePath)
                                  && !string.IsNullOrWhiteSpace(mapping.DriveLetter))
                .Select(mapping => new GuestSharedFolderMapping
                {
                    Id = mapping.Id,
                    Label = mapping.Label,
                    SharePath = mapping.SharePath,
                    DriveLetter = mapping.DriveLetter,
                    Persistent = mapping.Persistent,
                    Enabled = mapping.Enabled
                })
                .ToList()
                ?? [];

            if (mappings.Count == 0)
            {
                var baseDrive = GuestConfigService.NormalizeDriveLetter(_config?.SharedFolders?.BaseDriveLetter);
                await _winFspMountService.UnmountAsync(baseDrive, cancellationToken);
                return new SharedFolderReconnectCycleResult();
            }

            var baseDriveLetter = GuestConfigService.NormalizeDriveLetter(_config?.SharedFolders?.BaseDriveLetter);
            attempted = mappings.Count;

            try
            {
                await _winFspMountService.EnsureCatalogMountedAsync(baseDriveLetter, mappings, cancellationToken);
                newlyMounted = mappings.Count;
                ResetRateLimitedWarning("sharedfolders.reconnect.failed", baseDriveLetter);
                GuestLogger.Info("sharedfolders.reconnect.catalog_ready", "HyperTool-File-Katalog bereit.", new
                {
                    driveLetter = baseDriveLetter,
                    count = mappings.Count,
                    mode = "hypertool-file-catalog"
                });
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                WarnRateLimited(
                    "sharedfolders.reconnect.failed",
                    ex.Message,
                    new
                    {
                        driveLetter = baseDriveLetter,
                        count = mappings.Count,
                        exceptionType = ex.GetType().FullName
                    },
                    baseDriveLetter,
                    RecurringWarnRateLimitInterval);
                failed = mappings.Count;
            }

            return new SharedFolderReconnectCycleResult
            {
                Attempted = attempted,
                NewlyMounted = newlyMounted,
                Failed = failed
            };
        }
        finally
        {
            _sharedFolderReconnectGate.Release();
        }
    }

    private async Task RefreshSharedFolderHostNameFromSocketAsync(CancellationToken cancellationToken)
    {
        var identity = await FetchHostIdentityViaHyperVSocketAsync(cancellationToken);
        if (identity is null)
        {
            return;
        }

        ApplyHostIdentityToConfig(identity, persistConfig: true);
    }

    private void ApplyHostIdentityToConfig(HostIdentityInfo identity, bool persistConfig)
    {
        if (_config is null)
        {
            return;
        }

        _config.Usb ??= new GuestUsbSettings();
        _config.SharedFolders ??= new GuestSharedFolderSettings();

        var normalizedHostName = (identity.HostName ?? string.Empty).Trim();

        var changed = false;
        if (!string.IsNullOrWhiteSpace(normalizedHostName)
            && !string.Equals(_config.Usb.HostName, normalizedHostName, StringComparison.OrdinalIgnoreCase))
        {
            _config.Usb.HostName = normalizedHostName;
            changed = true;
        }

        if (!string.IsNullOrWhiteSpace(normalizedHostName)
            && !string.Equals(_config.Usb.HostAddress, normalizedHostName, StringComparison.OrdinalIgnoreCase))
        {
            _config.Usb.HostAddress = normalizedHostName;
            changed = true;
        }

        var usbEnabled = identity.Features?.UsbSharingEnabled != false;
        if (_config.Usb.HostFeatureEnabled != usbEnabled)
        {
            _config.Usb.HostFeatureEnabled = usbEnabled;
            changed = true;
        }

        var sharedFoldersEnabled = identity.Features?.SharedFoldersEnabled != false;
        if (_config.SharedFolders.HostFeatureEnabled != sharedFoldersEnabled)
        {
            _config.SharedFolders.HostFeatureEnabled = sharedFoldersEnabled;
            changed = true;
        }

        if (!changed)
        {
            return;
        }

        if (persistConfig)
        {
            GuestConfigService.Save(_configPath, _config);
        }

        _mainWindow?.UpdateHostFeatureAvailability(
            usbFeatureEnabledByHost: _config.Usb.HostFeatureEnabled,
            sharedFoldersFeatureEnabledByHost: _config.SharedFolders.HostFeatureEnabled,
            hostName: _config.Usb.HostName);

        GuestLogger.Info("guest.hostidentity.updated", "Host-Feature-Status per Hyper-V Socket aktualisiert.", new
        {
            hostName = _config.Usb.HostName,
            usbSharingEnabled = _config.Usb.HostFeatureEnabled,
            sharedFoldersEnabled = _config.SharedFolders.HostFeatureEnabled
        });
    }

    private bool HasConfiguredUsbAutoConnect()
    {
        var keys = _config?.Usb?.AutoConnectDeviceKeys;
        return keys is not null && keys.Any(static key => !string.IsNullOrWhiteSpace(key));
    }

    private string? ResolveSharedFolderHostTarget()
    {
        var hostName = (_config?.Usb?.HostName ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(hostName))
        {
            return hostName;
        }

        return (_config?.Usb?.HostAddress ?? string.Empty).Trim();
    }

    private async Task<int> ApplyUsbAutoConnectAsync(IReadOnlyList<UsbIpDeviceInfo> devices)
    {
        var configuredKeys = _config?.Usb?.AutoConnectDeviceKeys;
        if (configuredKeys is null || configuredKeys.Count == 0)
        {
            return 0;
        }

        var keySet = new HashSet<string>(
            configuredKeys
                .Where(static key => !string.IsNullOrWhiteSpace(key))
                .Select(static key => key.Trim()),
            StringComparer.OrdinalIgnoreCase);

        if (keySet.Count == 0)
        {
            return 0;
        }

        var candidates = devices
            .Where(device =>
                !string.IsNullOrWhiteSpace(device.BusId)
                && !device.IsAttached
                && keySet.Contains(BuildUsbAutoConnectKey(device)))
            .Select(device => device.BusId.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var connectedCount = 0;
        foreach (var busId in candidates)
        {
            var result = await ConnectUsbAsync(busId);
            if (result == 0)
            {
                connectedCount++;
            }
        }

        return connectedCount;
    }

    private static string BuildUsbAutoConnectKey(UsbIpDeviceInfo device)
    {
        if (!string.IsNullOrWhiteSpace(device.HardwareId))
        {
            return "hardware:" + device.HardwareId.Trim();
        }

        if (!string.IsNullOrWhiteSpace(device.Description))
        {
            return "description:" + device.Description.Trim();
        }

        if (!string.IsNullOrWhiteSpace(device.BusId))
        {
            return "busid:" + device.BusId.Trim();
        }

        return string.Empty;
    }

    private void StartUsbDiagnosticsLoop()
    {
        StopUsbDiagnosticsLoop();
        var cts = new CancellationTokenSource();
        _usbDiagnosticsCts = cts;
        _usbDiagnosticsTask = Task.Run(() => RunUsbDiagnosticsLoopAsync(cts.Token));
    }

    private void StopUsbDiagnosticsLoop()
    {
        try
        {
            _usbDiagnosticsCts?.Cancel();
        }
        catch
        {
        }

        _usbDiagnosticsCts?.Dispose();
        _usbDiagnosticsCts = null;
        _usbDiagnosticsTask = null;
    }

    private async Task RunUsbDiagnosticsLoopAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                var probeSucceeded = await ProbeHyperVSocketServiceAsync(cancellationToken);
                if (probeSucceeded)
                {
                    _usbHyperVSocketProbeFailureCount = 0;
                    _usbHyperVSocketServiceReachable = true;
                }
                else
                {
                    _usbHyperVSocketProbeFailureCount++;
                    if (_usbHyperVSocketProbeFailureCount >= 3)
                    {
                        _usbHyperVSocketServiceReachable = false;
                    }
                }

                UpdateUsbDiagnosticsPanel();
            }
            catch
            {
            }

            try
            {
                await Task.Delay(TimeSpan.FromSeconds(2), cancellationToken);
            }
            catch (OperationCanceledException)
            {
                break;
            }
        }
    }

    private async Task<bool> ProbeHyperVSocketServiceAsync(CancellationToken cancellationToken)
    {
        if (_usbHyperVSocketProxy?.IsRunning != true)
        {
            return false;
        }

        var serviceIdText = (_config?.Usb?.HyperVSocketServiceId ?? HyperVSocketUsbTunnelDefaults.ServiceIdString).Trim();
        if (!Guid.TryParse(serviceIdText, out var serviceId))
        {
            return false;
        }

        for (var attempt = 0; attempt < 2; attempt++)
        {
            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            linkedCts.CancelAfter(TimeSpan.FromMilliseconds(1500));

            try
            {
                using var socket = new System.Net.Sockets.Socket((System.Net.Sockets.AddressFamily)34, System.Net.Sockets.SocketType.Stream, (System.Net.Sockets.ProtocolType)1);
                linkedCts.Token.ThrowIfCancellationRequested();
                socket.Connect(new HyperVSocketEndPoint(HyperVSocketUsbTunnelDefaults.VmIdParent, serviceId));
                return true;
            }
            catch
            {
                if (attempt == 0)
                {
                    await Task.Delay(120, cancellationToken);
                }
            }
        }

        return false;
    }

    private async Task<(bool hyperVSocketActive, bool registryServiceOk)> RunTransportDiagnosticsTestAsync()
    {
        UpdateUsbTransportBridge();

        var hyperVSocketActive = _usbHyperVSocketProxy?.IsRunning == true;
        var registryServiceOk = false;

        if (hyperVSocketActive)
        {
            try
            {
                registryServiceOk = await ProbeHyperVSocketServiceAsync(CancellationToken.None);
            }
            catch
            {
                registryServiceOk = false;
            }
        }

        if (registryServiceOk)
        {
            _usbHyperVSocketProbeFailureCount = 0;
            _usbHyperVSocketServiceReachable = true;
        }
        else
        {
            _usbHyperVSocketProbeFailureCount = Math.Max(_usbHyperVSocketProbeFailureCount + 1, 3);
            _usbHyperVSocketServiceReachable = false;
        }

        ApplyUsbTransportResolution(ResolveUsbHostAddressDiagnostics());
        UpdateUsbDiagnosticsPanel();

        GuestLogger.Info("usb.transport.hyperv.test", "Hyper-V Socket Test ausgeführt.", new
        {
            hyperVSocketActive,
            registryServiceOk
        });

        try
        {
            await SendHyperVDiagnosticsAckAsync(Environment.MachineName, hyperVSocketActive, registryServiceOk, CancellationToken.None);
            GuestLogger.Info("usb.transport.hyperv.test.host_ack_sent", "Hyper-V Socket Test an Host übermittelt.", new
            {
                guestComputerName = Environment.MachineName,
                hyperVSocketActive,
                registryServiceOk
            });
        }
        catch (Exception ex)
        {
            GuestLogger.Warn("usb.transport.hyperv.test.host_ack_failed", ex.Message, new
            {
                guestComputerName = Environment.MachineName,
                hyperVSocketActive,
                registryServiceOk,
                exceptionType = ex.GetType().FullName
            });
        }

        return (hyperVSocketActive, registryServiceOk);
    }

    private static async Task SendHyperVDiagnosticsAckAsync(string guestComputerName, bool hyperVSocketActive, bool registryServiceOk, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(guestComputerName))
        {
            return;
        }

        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        linkedCts.CancelAfter(TimeSpan.FromMilliseconds(1200));

        using var socket = new System.Net.Sockets.Socket((System.Net.Sockets.AddressFamily)34, System.Net.Sockets.SocketType.Stream, (System.Net.Sockets.ProtocolType)1);
        linkedCts.Token.ThrowIfCancellationRequested();
        socket.Connect(new HyperVSocketEndPoint(HyperVSocketUsbTunnelDefaults.VmIdParent, HyperVSocketUsbTunnelDefaults.DiagnosticsServiceId));

        await using var stream = new System.Net.Sockets.NetworkStream(socket, ownsSocket: true);
        await using var writer = new StreamWriter(stream, Encoding.UTF8, bufferSize: 256, leaveOpen: false)
        {
            NewLine = "\n"
        };

        var payload = JsonSerializer.Serialize(new
        {
            guestComputerName,
            hyperVSocketActive,
            registryServiceOk,
            sentAtUtc = DateTime.UtcNow.ToString("O")
        });

        await writer.WriteLineAsync(payload.AsMemory(), linkedCts.Token);
        await writer.FlushAsync(linkedCts.Token);
    }

    private static async Task TrySendUsbConnectionEventAckAsync(string busId, string eventType, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(busId) || string.IsNullOrWhiteSpace(eventType))
        {
            return;
        }

        try
        {
            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            linkedCts.CancelAfter(TimeSpan.FromMilliseconds(1200));

            using var socket = new System.Net.Sockets.Socket((System.Net.Sockets.AddressFamily)34, System.Net.Sockets.SocketType.Stream, (System.Net.Sockets.ProtocolType)1);
            linkedCts.Token.ThrowIfCancellationRequested();
            socket.Connect(new HyperVSocketEndPoint(HyperVSocketUsbTunnelDefaults.VmIdParent, HyperVSocketUsbTunnelDefaults.DiagnosticsServiceId));

            await using var stream = new System.Net.Sockets.NetworkStream(socket, ownsSocket: true);
            await using var writer = new StreamWriter(stream, Encoding.UTF8, bufferSize: 256, leaveOpen: false)
            {
                NewLine = "\n"
            };

            var payload = JsonSerializer.Serialize(new HyperVSocketDiagnosticsAck
            {
                GuestComputerName = Environment.MachineName,
                BusId = busId.Trim(),
                EventType = eventType.Trim(),
                SentAtUtc = DateTime.UtcNow.ToString("O")
            });

            await writer.WriteLineAsync(payload.AsMemory(), linkedCts.Token);
            await writer.FlushAsync(linkedCts.Token);
        }
        catch
        {
        }
    }

    private void ApplyUsbTransportResolution(UsbHostResolution resolution, bool fallbackToIp = false)
    {
        var usesHyperVSocket = string.Equals(resolution.Source, "hyperv-socket", StringComparison.OrdinalIgnoreCase)
            && !string.IsNullOrWhiteSpace(resolution.ResolvedIpv4);

        var fallbackActive = fallbackToIp
            || ((_config?.Usb?.UseHyperVSocket ?? false)
                && !usesHyperVSocket
                && !string.IsNullOrWhiteSpace(resolution.ResolvedIpv4));

        var changed = _usbHyperVSocketTransportActive != usesHyperVSocket || _usbIpFallbackActive != fallbackActive;
        _usbHyperVSocketTransportActive = usesHyperVSocket;
        _usbIpFallbackActive = fallbackActive;

        if (!changed)
        {
            return;
        }

        UpdateUsbDiagnosticsPanel();
    }

    private void UpdateUsbDiagnosticsPanel()
    {
        void apply()
        {
            _mainWindow?.UpdateTransportDiagnostics(
                hyperVSocketActive: _usbHyperVSocketTransportActive,
                registryServicePresent: _usbHyperVSocketServiceReachable,
                fallbackActive: _usbIpFallbackActive);
        }

        if (_mainWindow?.DispatcherQueue is { } queue && !queue.HasThreadAccess)
        {
            _ = queue.TryEnqueue(apply);
            return;
        }

        apply();
    }

    private void UpdateUsbTransportBridge()
    {
        var usbSettings = _config?.Usb;
        if (usbSettings is null || !usbSettings.UseHyperVSocket)
        {
            if (_usbHyperVSocketProxy is not null)
            {
                try
                {
                    _usbHyperVSocketProxy.Dispose();
                }
                catch
                {
                }

                _usbHyperVSocketProxy = null;
            }

            _usbHyperVSocketTransportActive = false;
            _usbHyperVSocketServiceReachable = false;
            _usbHyperVSocketProbeFailureCount = 0;
            UpdateUsbDiagnosticsPanel();
            return;
        }

        var serviceIdText = string.IsNullOrWhiteSpace(usbSettings.HyperVSocketServiceId)
            ? HyperVSocketUsbTunnelDefaults.ServiceIdString
            : usbSettings.HyperVSocketServiceId;

        if (!Guid.TryParse(serviceIdText, out var serviceId))
        {
            _usbHyperVSocketProxy?.Dispose();
            _usbHyperVSocketProxy = null;
            _usbHyperVSocketTransportActive = false;
            _usbHyperVSocketServiceReachable = false;
            _usbHyperVSocketProbeFailureCount = 0;
            UpdateUsbDiagnosticsPanel();

            GuestLogger.Warn("usb.transport.hyperv.invalid_guid", "Hyper-V Socket ServiceId ist ungültig. USB fällt auf IP-Fallback zurück.", new
            {
                serviceId = serviceIdText
            });
            return;
        }

        if (_usbHyperVSocketProxy?.IsRunning == true)
        {
            return;
        }

        try
        {
            _usbHyperVSocketProxy?.Dispose();
            _usbHyperVSocketProxy = new HyperVSocketUsbGuestProxy(serviceId);
            _usbHyperVSocketProxy.Start();
            _usbHyperVSocketTransportActive = false;
            UpdateUsbDiagnosticsPanel();

            GuestLogger.Info("usb.transport.hyperv.started", "Hyper-V Socket USB Proxy aktiv (primärer Transport).", new
            {
                serviceId = serviceId.ToString("D"),
                loopback = HyperVSocketUsbTunnelDefaults.LoopbackAddress,
                port = HyperVSocketUsbTunnelDefaults.UsbIpTcpPort
            });
        }
        catch (Exception ex)
        {
            _usbHyperVSocketProxy?.Dispose();
            _usbHyperVSocketProxy = null;
            _usbHyperVSocketTransportActive = false;
            _usbHyperVSocketServiceReachable = false;
            _usbHyperVSocketProbeFailureCount = 0;
            UpdateUsbDiagnosticsPanel();

            GuestLogger.Warn("usb.transport.hyperv.failed", ex.Message, new
            {
                serviceId = serviceId.ToString("D"),
                fallback = "ip"
            });
        }
    }

    private static string BuildSharedFolderReconnectRateLimitKey(GuestSharedFolderMapping mapping)
    {
        if (!string.IsNullOrWhiteSpace(mapping.Id))
        {
            return mapping.Id.Trim();
        }

        var driveLetter = GuestConfigService.NormalizeDriveLetter(mapping.DriveLetter ?? string.Empty);
        var sharePath = (mapping.SharePath ?? string.Empty).Trim();
        return $"{driveLetter}|{sharePath}";
    }

    private void WarnRateLimited(string eventName, string message, object? data, string scopeKey, TimeSpan minInterval)
    {
        var key = $"{eventName}:{scopeKey}";
        var now = DateTimeOffset.UtcNow;

        while (true)
        {
            var existing = _rateLimitedWarnStates.GetOrAdd(
                key,
                _ => new RateLimitedWarnState
                {
                    LastLoggedAtUtc = now,
                    SuppressedCount = 0
                });

            var sinceLastLog = now - existing.LastLoggedAtUtc;
            if (sinceLastLog >= minInterval)
            {
                var suppressedCount = existing.SuppressedCount;
                var replacement = new RateLimitedWarnState
                {
                    LastLoggedAtUtc = now,
                    SuppressedCount = 0
                };

                if (_rateLimitedWarnStates.TryUpdate(key, replacement, existing))
                {
                    var finalMessage = suppressedCount > 0
                        ? $"{message} (weitere {suppressedCount} gleichartige Warnung(en) unterdrückt)"
                        : message;
                    GuestLogger.Warn(eventName, finalMessage, data);
                    return;
                }

                continue;
            }

            var suppressedReplacement = new RateLimitedWarnState
            {
                LastLoggedAtUtc = existing.LastLoggedAtUtc,
                SuppressedCount = existing.SuppressedCount + 1
            };

            if (_rateLimitedWarnStates.TryUpdate(key, suppressedReplacement, existing))
            {
                return;
            }
        }
    }

    private void ResetRateLimitedWarning(string eventName, string scopeKey)
    {
        var key = $"{eventName}:{scopeKey}";
        _rateLimitedWarnStates.TryRemove(key, out _);
    }

    private async Task<IReadOnlyList<HostSharedFolderDefinition>> FetchHostSharedFoldersAsync()
    {
        if (_config?.SharedFolders?.HostFeatureEnabled == false)
        {
            throw new InvalidOperationException("Shared-Folder Funktion ist durch den Host deaktiviert.");
        }

        if (_config?.Usb?.UseHyperVSocket != true)
        {
            throw new InvalidOperationException("Host-Liste per Hyper-V Socket ist deaktiviert. Aktiviere in den USB-Einstellungen 'Hyper-V Socket verwenden (bevorzugt)'.");
        }

        var client = new HyperVSocketSharedFolderCatalogGuestClient();
        Exception? lastError = null;

        for (var attempt = 1; attempt <= SharedFolderCatalogFetchMaxAttempts; attempt++)
        {
            try
            {
                var catalog = await client.FetchCatalogAsync(CancellationToken.None);
                return catalog;
            }
            catch (Exception ex) when (attempt < SharedFolderCatalogFetchMaxAttempts && IsTransientCatalogFetchError(ex))
            {
                lastError = ex;
                await Task.Delay(TimeSpan.FromMilliseconds(250 * attempt));
            }
            catch (Exception ex)
            {
                lastError = ex;
                break;
            }
        }

        var errorMessage = lastError?.Message ?? "Unbekannter Fehler";
        GuestLogger.Warn("sharedfolders.catalog.fetch_failed", errorMessage, new
        {
            transport = "hyperv-socket",
            attempts = SharedFolderCatalogFetchMaxAttempts,
            exceptionType = lastError?.GetType().FullName
        });

        var userMessage = "Host-Liste konnte nicht geladen werden (Hyper-V Socket abgelehnt). Prüfe, ob HyperTool Host läuft, und starte Host/Guest ggf. neu.";
        if (lastError is SocketException socketException
            && socketException.SocketErrorCode == SocketError.ConnectionRefused)
        {
            userMessage = "Host-Liste konnte nicht geladen werden (Hyper-V Socket verweigert). Starte HyperTool Host einmalig als Administrator, damit der Shared-Folder-Hyper-V-Dienst registriert wird.";
        }

        throw new InvalidOperationException(
            userMessage,
            lastError);
    }

    private static bool IsTransientCatalogFetchError(Exception exception)
    {
        if (exception is not SocketException socketException)
        {
            return false;
        }

        return socketException.SocketErrorCode is SocketError.ConnectionRefused
            or SocketError.TimedOut
            or SocketError.NetworkDown
            or SocketError.NetworkUnreachable
            or SocketError.HostDown
            or SocketError.HostUnreachable
            or SocketError.TryAgain
            or SocketError.WouldBlock;
    }

    private bool TryInitializeSingleInstance()
    {
        try
        {
            _singleInstanceMutex = new Mutex(initiallyOwned: true, SingleInstanceMutexName, out var createdNew);
            if (!createdNew)
            {
                try
                {
                    _singleInstanceMutex.Dispose();
                }
                catch
                {
                }

                _singleInstanceMutex = null;
                TryNotifyRunningInstanceToShow();
                return false;
            }

            _singleInstanceOwned = true;
            _singleInstanceServerCts = new CancellationTokenSource();
            _singleInstanceServerTask = Task.Run(() => RunSingleInstancePipeServerAsync(_singleInstanceServerCts.Token));
            return true;
        }
        catch
        {
            return true;
        }
    }

    private void TryNotifyRunningInstanceToShow()
    {
        try
        {
            using var client = new NamedPipeClientStream(".", SingleInstancePipeName, PipeDirection.Out, PipeOptions.Asynchronous);
            client.Connect(450);
            using var writer = new StreamWriter(client, Encoding.UTF8, leaveOpen: false);
            writer.WriteLine("SHOW");
            writer.Flush();
        }
        catch
        {
        }
    }

    private async Task RunSingleInstancePipeServerAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                using var server = new NamedPipeServerStream(
                    SingleInstancePipeName,
                    PipeDirection.In,
                    maxNumberOfServerInstances: 1,
                    PipeTransmissionMode.Byte,
                    PipeOptions.Asynchronous);

                await server.WaitForConnectionAsync(cancellationToken);

                using var reader = new StreamReader(server, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true);
                var command = (await reader.ReadLineAsync())?.Trim();
                if (string.IsNullOrWhiteSpace(command))
                {
                    continue;
                }

                if (string.Equals(command, "SHOW", StringComparison.OrdinalIgnoreCase))
                {
                    if (_mainWindow is null)
                    {
                        _pendingSingleInstanceShow = true;
                        continue;
                    }

                    _ = _mainWindow.DispatcherQueue.TryEnqueue(BringMainWindowToFront);
                    continue;
                }

                if (string.Equals(command, "REFRESH_HOST_FEATURES", StringComparison.OrdinalIgnoreCase))
                {
                    _ = Task.Run(() => HandleHostFeatureRefreshSignalAsync(cancellationToken), cancellationToken);
                }
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch
            {
                try
                {
                    await Task.Delay(250, cancellationToken);
                }
                catch
                {
                }
            }
        }
    }

    private async Task HandleHostFeatureRefreshSignalAsync(CancellationToken cancellationToken)
    {
        try
        {
            var identity = await FetchHostIdentityViaHyperVSocketAsync(cancellationToken);
            if (identity is not null)
            {
                ApplyHostIdentityToConfig(identity, persistConfig: true);
            }

            await RefreshUsbDevicesAsync(emitLogs: false);
            TriggerSharedFolderReconnectCycle();
        }
        catch
        {
        }
    }

    private void BringMainWindowToFront()
    {
        _pendingSingleInstanceShow = false;

        if (_mainWindow is null)
        {
            return;
        }

        try
        {
            _mainWindow.AppWindow.Show();
        }
        catch
        {
        }

        try
        {
            _mainWindow.Activate();
        }
        catch
        {
        }

        try
        {
            var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(_mainWindow);
            if (hwnd != nint.Zero)
            {
                if (IsIconic(hwnd))
                {
                    _ = ShowWindow(hwnd, SwRestore);
                }
                else
                {
                    _ = ShowWindow(hwnd, SwShow);
                }

                _ = SetForegroundWindow(hwnd);
            }
        }
        catch
        {
        }
    }

    private void ShutdownSingleInstanceInfrastructure()
    {
        try
        {
            _singleInstanceServerCts?.Cancel();
        }
        catch
        {
        }

        _singleInstanceServerCts?.Dispose();
        _singleInstanceServerCts = null;
        _singleInstanceServerTask = null;

        if (_singleInstanceMutex is null)
        {
            return;
        }

        try
        {
            if (_singleInstanceOwned)
            {
                _singleInstanceMutex.ReleaseMutex();
            }
        }
        catch
        {
        }
        finally
        {
            _singleInstanceMutex.Dispose();
            _singleInstanceMutex = null;
            _singleInstanceOwned = false;
        }
    }

    private static bool IsCommandMode(string[] args)
    {
        if (args.Length == 0)
        {
            return false;
        }

        var known = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "status",
            "once",
            "run",
            "unmap",
            "handshake",
            "install-autostart",
            "remove-autostart",
            "autostart-status"
        };

        return args.Any(arg => known.Contains(arg));
    }

    private static string ResolveConfigPath(string[] args)
    {
        var index = Array.FindIndex(args, item => string.Equals(item, "--config", StringComparison.OrdinalIgnoreCase));
        if (index >= 0 && index + 1 < args.Length)
        {
            return Path.GetFullPath(args[index + 1]);
        }

        return GuestConfigService.DefaultConfigPath;
    }

    private static string[] SplitArgs(string? commandLine)
    {
        if (string.IsNullOrWhiteSpace(commandLine))
        {
            return [];
        }

        var result = new List<string>();
        var current = new StringBuilder();
        var inQuotes = false;

        foreach (var ch in commandLine)
        {
            if (ch == '"')
            {
                inQuotes = !inQuotes;
                continue;
            }

            if (char.IsWhiteSpace(ch) && !inQuotes)
            {
                if (current.Length > 0)
                {
                    result.Add(current.ToString());
                    current.Clear();
                }

                continue;
            }

            current.Append(ch);
        }

        if (current.Length > 0)
        {
            result.Add(current.ToString());
        }

        return result.ToArray();
    }

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetCursorPos(out NativePoint lpPoint);

    [StructLayout(LayoutKind.Sequential)]
    private struct NativePoint
    {
        public int X;
        public int Y;
    }
}
