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
using System.Runtime.InteropServices;
using System.Text;
using Windows.Graphics;
using Windows.UI;

namespace HyperTool.Guest;

public sealed partial class App : Application
{
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
        GuestLogger.Initialize(_config.Logging);
        await RefreshUsbClientAvailabilityAsync();

        ApplyStartWithWindows(_config);

        _mainWindow = new GuestMainWindow(
            _config,
            RefreshUsbDevicesAsync,
            ConnectUsbAsync,
            DisconnectUsbAsync,
            SaveConfigAsync,
            RestartForThemeChangeAsync,
            _isUsbClientAvailable);

        _mainWindow.AppWindow.Closing += OnMainWindowClosing;

        _minimizeToTray = _config.Ui.MinimizeToTray;
        _mainWindow.ApplyTheme(_config.Ui.Theme);

        TryInitializeTray();
        await RefreshUsbDevicesAsync();

        if (_config.Ui.StartMinimized && _isTrayFunctional && !_pendingSingleInstanceShow)
        {
            _mainWindow.AppWindow.Hide();
        }
        else
        {
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

        if (_pendingSingleInstanceShow)
        {
            BringMainWindowToFront();
        }
    }

    private async Task SaveConfigAsync(GuestConfig config)
    {
        _config = config;
        GuestConfigService.Save(_configPath, config);

        ApplyStartWithWindows(config);
        _minimizeToTray = config.Ui.MinimizeToTray;
        UpdateTrayControlCenterView();

        await Task.CompletedTask;
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
            PointInt32 exitPosition;
            SizeInt32 exitSize;

            if (_mainWindow is not null)
            {
                exitPosition = _mainWindow.AppWindow.Position;
                exitSize = _mainWindow.AppWindow.Size;

                try
                {
                    await _mainWindow.PlayExitFadeAsync();
                }
                catch
                {
                }

                try
                {
                    _mainWindow.AppWindow.Hide();
                }
                catch
                {
                }
            }
            else
            {
                exitPosition = new PointInt32(120, 120);
                exitSize = new SizeInt32(980, 640);
            }

            _isExitRequested = true;

            _trayControlCenterWindow?.Close();
            _trayControlCenterWindow = null;

            _trayService?.Dispose();
            _trayService = null;

            ShutdownSingleInstanceInfrastructure();

            if (showExitScreen)
            {
                try
                {
                    var exitWindow = new HyperTool.WinUI.Views.ExitScreenWindow(GuestWindowTitle, GuestHeadline, GuestIconUri);
                    exitWindow.ConfigureBounds(exitPosition, exitSize);
                    exitWindow.Activate();
                    await exitWindow.PlayAndCloseAsync();
                }
                catch
                {
                }
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
                _isUsbClientAvailable);

            nextWindow.AppWindow.Closing += OnMainWindowClosing;
            _mainWindow = nextWindow;
            _minimizeToTray = _config.Ui.MinimizeToTray;
            _mainWindow.ApplyTheme(_config.Ui.Theme);

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
            }

            UpdateTrayControlCenterView();
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
        if (_trayControlCenterWindow is null || _mainWindow is null || _config is null)
        {
            return;
        }

        _trayControlCenterWindow.ApplyTheme(_mainWindow.CurrentTheme == "dark");
        _trayControlCenterWindow.UpdateView(_usbDevices, _selectedUsbBusId, _mainWindow.AppWindow.IsVisible, _minimizeToTray);

        if (_trayControlCenterWindow.AppWindow.IsVisible)
        {
            PositionTrayControlCenterNearTray();
        }
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

    private async Task<IReadOnlyList<UsbIpDeviceInfo>> RefreshUsbDevicesAsync()
    {
        if (!_isUsbClientAvailable)
        {
            _usbDevices.Clear();
            _mainWindow?.UpdateUsbDevices(_usbDevices);
            UpdateTrayControlCenterView();

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
        var previousCount = _usbDevices.Count;

        try
        {
            IReadOnlyList<UsbIpDeviceInfo> list;
            var useRemoteHostList = !string.IsNullOrWhiteSpace(hostResolution.ResolvedIpv4);
            if (useRemoteHostList)
            {
                list = await _usbService.GetRemoteDevicesAsync(hostResolution.ResolvedIpv4, CancellationToken.None);
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

            _usbDevices.Clear();
            _usbDevices.AddRange(list);

            _mainWindow?.UpdateUsbDevices(_usbDevices);
            UpdateTrayControlCenterView();

            GuestLogger.Info("usb.refresh.success", "USB-Geräte aktualisiert.", new
            {
                operationId,
                elapsedMs = stopwatch.ElapsedMilliseconds,
                count = _usbDevices.Count,
                selectedBusId = _selectedUsbBusId,
                attachedCount = _usbDevices.Count(item => item.IsAttached),
                sharedCount = _usbDevices.Count(item => item.IsShared),
                retryTriggered,
                previousCount,
                listSource = useRemoteHostList ? "remote-host" : "local-state",
                hostAddress = hostResolution.ResolvedIpv4,
                hostSource = hostResolution.Source,
                hostInput = hostResolution.RawInput
            });

            return _usbDevices;
        }
        catch (Exception ex)
        {
            GuestLogger.Warn("usb.refresh.failed", ex.Message, new
            {
                operationId,
                elapsedMs = stopwatch.ElapsedMilliseconds,
                selectedBusId = _selectedUsbBusId,
                exceptionType = ex.GetType().FullName,
                hostAddress = hostResolution.ResolvedIpv4,
                hostSource = hostResolution.Source,
                hostInput = hostResolution.RawInput,
                hostResolutionFailure = hostResolution.FailureReason
            });
            return _usbDevices;
        }
    }

    private async Task<int> ConnectUsbAsync(string busId)
    {
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
                hostResolutionFailure = hostResolution.FailureReason,
                sharePath = _config?.SharePath
            });
            return 1;
        }

        try
        {
            await _usbService.AttachToHostAsync(busId, hostResolution.ResolvedIpv4, CancellationToken.None);
            _selectedUsbBusId = busId;

            GuestLogger.Info("usb.connect.success", "USB Host-Attach erfolgreich.", new
            {
                operationId,
                busId,
                elapsedMs = stopwatch.ElapsedMilliseconds,
                hostAddress = hostResolution.ResolvedIpv4,
                hostSource = hostResolution.Source
            });

            return 0;
        }
        catch (Exception ex)
        {
            GuestLogger.Warn("usb.connect.failed", ex.Message, new
            {
                operationId,
                busId,
                elapsedMs = stopwatch.ElapsedMilliseconds,
                hostAddress = hostResolution.ResolvedIpv4,
                hostSource = hostResolution.Source,
                exceptionType = ex.GetType().FullName
            });
            return 1;
        }
    }

    private async Task<int> DisconnectUsbAsync(string busId)
    {
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

            GuestLogger.Info("usb.disconnect.success", "USB Host-Detach erfolgreich.", new
            {
                operationId,
                busId,
                elapsedMs = stopwatch.ElapsedMilliseconds,
                hostAddress = hostResolution.ResolvedIpv4,
                hostSource = hostResolution.Source
            });

            return 0;
        }
        catch (Exception ex)
        {
            GuestLogger.Warn("usb.disconnect.failed", ex.Message, new
            {
                operationId,
                busId,
                elapsedMs = stopwatch.ElapsedMilliseconds,
                hostAddress = hostResolution.ResolvedIpv4,
                hostSource = hostResolution.Source,
                exceptionType = ex.GetType().FullName
            });
            return 1;
        }
    }

    private UsbHostResolution ResolveUsbHostAddressDiagnostics()
    {
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

            return new UsbHostResolution(string.Empty, "sharePath", host, "UNC-Host konnte nicht in IPv4 aufgelöst werden.");
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

        await DisconnectUsbAsync(selected.BusId);
        await RefreshUsbDevicesAsync();
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
                var command = await reader.ReadLineAsync();
                if (!string.Equals(command?.Trim(), "SHOW", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (_mainWindow is null)
                {
                    _pendingSingleInstanceShow = true;
                    continue;
                }

                _ = _mainWindow.DispatcherQueue.TryEnqueue(BringMainWindowToFront);
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
