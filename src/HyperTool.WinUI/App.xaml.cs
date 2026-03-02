using HyperTool.Services;
using HyperTool.ViewModels;
using HyperTool.WinUI.Helpers;
using HyperTool.WinUI.Services;
using HyperTool.WinUI.Views;
using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Serilog;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Windows.Graphics;

namespace HyperTool.WinUI;

public sealed partial class App : Application
{
    private const string SingleInstanceMutexName = @"Local\HyperTool.WinUI.SingleInstance";
    private const string SingleInstancePipeName = "HyperTool.WinUI.SingleInstance.Activate";

    private ITrayService? _trayService;
    private ITrayControlCenterService? _trayControlCenterService;
    private bool _isExitRequested;
    private bool _isExitSequenceRunning;
    private bool _isRestartInProgress;
    private bool _isFatalErrorShown;
    private bool _isTrayFunctional;
    private bool _singleInstanceOwned;
    private bool _pendingSingleInstanceShow;
    private bool _isThemeWindowReopenInProgress;
    private bool _usbShutdownCleanupDone;

    private MainWindow? _mainWindow;
    private MainViewModel? _mainViewModel;
    private IThemeService? _themeService;
    private bool _minimizeToTray = true;
    private static bool _firstChanceTracingRegistered;
    private Mutex? _singleInstanceMutex;
    private CancellationTokenSource? _singleInstanceServerCts;
    private Task? _singleInstanceServerTask;

    public App()
    {
        RegisterFirstChanceTracing();
        UnhandledException += OnUnhandledException;
    }

    private static void RegisterFirstChanceTracing()
    {
        if (_firstChanceTracingRegistered)
        {
            return;
        }

        _firstChanceTracingRegistered = true;
        AppDomain.CurrentDomain.FirstChanceException += OnFirstChanceException;
    }

    private static void OnFirstChanceException(object? sender, FirstChanceExceptionEventArgs e)
    {
        if (e.Exception is not (InvalidCastException or System.Runtime.InteropServices.COMException))
        {
            return;
        }

        try
        {
            var traceDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "HyperTool", "diagnostics");
            Directory.CreateDirectory(traceDir);
            var traceFile = Path.Combine(traceDir, "firstchance.log");
            File.AppendAllText(traceFile, $"[{DateTime.Now:O}] {e.Exception}\n\n");
        }
        catch
        {
        }
    }

    protected override async void OnLaunched(LaunchActivatedEventArgs args)
    {
        EnsureControlResources();

        if (args.Arguments?.Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Any(arg => string.Equals(arg, "--restart-hns", StringComparison.OrdinalIgnoreCase)) == true)
        {
            RunRestartHnsHelperMode();
            return;
        }

        if (!TryInitializeSingleInstance())
        {
            Current.Exit();
            return;
        }

        RegisterGlobalExceptionHandlers();

        try
        {
            var logPath = InitializeLogging();
            Log.Information("Logging initialized at {LogPath}", logPath);

            IConfigService configService = new ConfigService();
            IHyperVService hyperVService = new HyperVPowerShellService();
            IHnsService hnsService = new HnsService();
            IStartupService startupService = new StartupService();
            IUpdateService updateService = new GitHubUpdateService();
            IUsbIpService usbIpService = new UsbIpdCliService();
            _themeService = new ThemeService();
            IUiInteropService uiInteropService = new UiInteropService();

            var configPath = ResolveConfigPath();
            TryMigrateLegacyConfig(configPath);
            var configResult = configService.LoadOrCreate(configPath);

            var uiConfig = configResult.Config.Ui;
            _minimizeToTray = uiConfig.MinimizeToTray;
            _themeService.ApplyTheme(uiConfig.Theme);

            if (!startupService.SetStartWithWindows(uiConfig.StartWithWindows, "HyperTool", Environment.ProcessPath ?? string.Empty, out var startupError)
                && !string.IsNullOrWhiteSpace(startupError))
            {
                Log.Warning("Could not apply startup setting: {StartupError}", startupError);
            }

            _mainViewModel = new MainViewModel(
                configResult,
                hyperVService,
                hnsService,
                configService,
                startupService,
                updateService,
                usbIpService,
                uiInteropService);

            _mainWindow = new MainWindow(_themeService, _mainViewModel, showStartupSplash: true);
            AttachMainWindowHandlers(_mainWindow);

            var shouldForceShowFromSecondLaunch = _pendingSingleInstanceShow;

            if (shouldForceShowFromSecondLaunch)
            {
                BringMainWindowToFront();
            }

            _trayService?.Dispose();
            _trayService = null;
            _trayControlCenterService?.Dispose();
            _trayControlCenterService = null;
            TryInitializeTray(_mainWindow, _mainViewModel);

            if (configResult.Config.Ui.StartMinimized && _isTrayFunctional && !shouldForceShowFromSecondLaunch)
            {
                _mainWindow.AppWindow.Hide();
            }
            else
            {
                _mainWindow.Activate();
            }

            Log.Information("Config loaded from {ConfigPath}", configResult.ConfigPath);
            Log.Information("HyperTool started (WinUI 3).");
        }
        catch (Exception ex)
        {
            ShowFatalErrorAndExit(ex, "HyperTool konnte beim Start nicht initialisiert werden.");
        }
    }

    private void AttachMainWindowHandlers(MainWindow window)
    {
        window.AppWindow.Closing += (_, eventArgs) =>
        {
            if (_isThemeWindowReopenInProgress || _isExitRequested)
            {
                return;
            }

            if (_mainViewModel is null)
            {
                return;
            }

            if (!_mainViewModel.TryPromptSaveConfigOnClose())
            {
                eventArgs.Cancel = true;
                return;
            }

            if (!_isTrayFunctional || !_minimizeToTray)
            {
                return;
            }

            eventArgs.Cancel = true;
            window.AppWindow.Hide();
        };

        window.Closed += async (_, _) =>
        {
            if (_isThemeWindowReopenInProgress)
            {
                return;
            }

            await ExecuteUsbShutdownCleanupAsync();

            _trayControlCenterService?.Dispose();
            _trayControlCenterService = null;
            _trayService?.Dispose();
            _trayService = null;
            ShutdownSingleInstanceInfrastructure();
            Log.Information("HyperTool exited.");
            Log.CloseAndFlush();
        };
    }

    public async Task ReopenMainWindowForThemeChangeAsync()
    {
        if (_isThemeWindowReopenInProgress)
        {
            return;
        }

        if (_mainWindow is null || _mainViewModel is null || _themeService is null)
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
                previousSize = new SizeInt32(MainWindow.DefaultWindowWidth, MainWindow.DefaultWindowHeight);
                previousVisible = true;
            }

            _trayControlCenterService?.Dispose();
            _trayControlCenterService = null;
            _trayService?.Dispose();
            _trayService = null;

            _themeService.ApplyTheme(_mainViewModel.UiTheme);

            var nextWindow = new MainWindow(_themeService, _mainViewModel, showStartupSplash: false);
            AttachMainWindowHandlers(nextWindow);
            _mainWindow = nextWindow;

            TryInitializeTray(nextWindow, _mainViewModel);

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

                try
                {
                    previousWindow.Close();
                }
                catch
                {
                }
            }
        }
        finally
        {
            _isThemeWindowReopenInProgress = false;
        }
    }

    private static void EnsureControlResources()
    {
        if (Current is not Application app)
        {
            return;
        }

        var resources = app.Resources;

        EnsureBrushResource(resources, "TabViewScrollButtonBackground", Windows.UI.Color.FromArgb(0x00, 0x00, 0x00, 0x00));
        EnsureBrushResource(resources, "TabViewScrollButtonBackgroundPointerOver", Windows.UI.Color.FromArgb(0x1F, 0x80, 0x80, 0x80));
        EnsureBrushResource(resources, "TabViewScrollButtonBackgroundPressed", Windows.UI.Color.FromArgb(0x33, 0x80, 0x80, 0x80));
        EnsureBrushResource(resources, "TabViewScrollButtonForeground", Colors.WhiteSmoke);
        EnsureBrushResource(resources, "TabViewScrollButtonForegroundPointerOver", Colors.White);
        EnsureBrushResource(resources, "TabViewScrollButtonForegroundPressed", Colors.White);
        EnsureBrushResource(resources, "TabViewScrollButtonForegroundDisabled", Windows.UI.Color.FromArgb(0x8F, 0xB0, 0xB0, 0xB0));

        EnsureBrushResource(resources, "TabViewButtonBackground", Windows.UI.Color.FromArgb(0x00, 0x00, 0x00, 0x00));
        EnsureBrushResource(resources, "TabViewButtonBackgroundPointerOver", Windows.UI.Color.FromArgb(0x1F, 0x90, 0x90, 0x90));
        EnsureBrushResource(resources, "TabViewButtonBackgroundPressed", Windows.UI.Color.FromArgb(0x33, 0x90, 0x90, 0x90));
        EnsureBrushResource(resources, "TabViewButtonForeground", Colors.WhiteSmoke);
        EnsureBrushResource(resources, "TabViewButtonForegroundPointerOver", Colors.White);
        EnsureBrushResource(resources, "TabViewButtonForegroundPressed", Colors.White);
        EnsureBrushResource(resources, "TabViewButtonForegroundDisabled", Windows.UI.Color.FromArgb(0x7F, 0xA5, 0xA5, 0xA5));
        EnsureBrushResource(resources, "TabViewButtonBorderBrush", Windows.UI.Color.FromArgb(0x00, 0x00, 0x00, 0x00));
        EnsureBrushResource(resources, "TabViewButtonBorderBrushPointerOver", Windows.UI.Color.FromArgb(0x3F, 0x90, 0x90, 0x90));
        EnsureBrushResource(resources, "TabViewButtonBorderBrushPressed", Windows.UI.Color.FromArgb(0x55, 0x90, 0x90, 0x90));

        try
        {
            var hasXamlControlsResources = resources.MergedDictionaries.OfType<XamlControlsResources>().Any();
            if (!hasXamlControlsResources)
            {
                resources.MergedDictionaries.Add(new XamlControlsResources());
            }
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "XamlControlsResources registration failed; continuing with fallback resources only.");
        }
    }

    private static void EnsureBrushResource(ResourceDictionary resources, string key, Windows.UI.Color color)
    {
        if (resources.ContainsKey(key))
        {
            return;
        }

        resources[key] = new SolidColorBrush(color);
    }

    private void OnUnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
    {
        if (IsRecoverableUiInteropError(e.Exception))
        {
            Log.Warning(
                e.Exception,
                "Recoverable UI interop error encountered. Continuing application. HResult={HResult}; Stack={Stack}",
                e.Exception.HResult,
                e.Exception.ToString());
            e.Handled = true;
            return;
        }

        ShowFatalErrorAndExit(e.Exception, "Ein unerwarteter UI-Fehler ist aufgetreten.");
        e.Handled = true;
    }

    private static bool IsRecoverableUiInteropError(Exception exception)
    {
        if (exception is InvalidCastException)
        {
            return true;
        }

        if (exception is System.Runtime.InteropServices.COMException comException)
        {
            return comException.HResult == unchecked((int)0x80004005)
                   || comException.HResult == unchecked((int)0x80004002)
                   || comException.HResult == unchecked((int)0x80004003);
        }

        return false;
    }

    private static string ResolveConfigPath()
    {
        try
        {
            var overridePath = Environment.GetEnvironmentVariable("HYPERTOOL_CONFIG_PATH");
            if (!string.IsNullOrWhiteSpace(overridePath))
            {
                return overridePath.Trim();
            }

            var localAppDataConfigPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "HyperTool",
                "HyperTool.config.json");

            var localAppDataDirectory = Path.GetDirectoryName(localAppDataConfigPath);
            if (!string.IsNullOrWhiteSpace(localAppDataDirectory))
            {
                Directory.CreateDirectory(localAppDataDirectory);
            }

            return localAppDataConfigPath;
        }
        catch
        {
            return Path.Combine(AppContext.BaseDirectory, "HyperTool.config.json");
        }
    }

    private static void TryMigrateLegacyConfig(string targetConfigPath)
    {
        try
        {
            var legacyConfigPath = Path.Combine(AppContext.BaseDirectory, "HyperTool.config.json");
            if (string.Equals(legacyConfigPath, targetConfigPath, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            if (!File.Exists(legacyConfigPath) || File.Exists(targetConfigPath))
            {
                return;
            }

            var targetDirectory = Path.GetDirectoryName(targetConfigPath);
            if (!string.IsNullOrWhiteSpace(targetDirectory))
            {
                Directory.CreateDirectory(targetDirectory);
            }

            File.Copy(legacyConfigPath, targetConfigPath, overwrite: false);
            Log.Information("Legacy config migrated from {LegacyConfigPath} to {TargetConfigPath}", legacyConfigPath, targetConfigPath);
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Legacy config migration failed.");
        }
    }

    private void TryInitializeTray(MainWindow mainWindow, MainViewModel mainViewModel)
    {
        _isTrayFunctional = false;

        try
        {
            Action requestExit = () => _ = RunExitSequenceAsync(mainWindow);

            _trayControlCenterService = new TrayControlCenterService(mainWindow.DispatcherQueue);
            _trayControlCenterService.Initialize(
                showMainWindowAction: () =>
                {
                    mainWindow.AppWindow.Show();
                    mainWindow.Activate();
                },
                hideMainWindowAction: () => mainWindow.AppWindow.Hide(),
                isMainWindowVisible: () =>
                {
                    try
                    {
                        return mainWindow.AppWindow?.IsVisible ?? true;
                    }
                    catch
                    {
                        return true;
                    }
                },
                getUiTheme: () => mainViewModel.UiTheme,
                getVms: () => mainViewModel.GetTrayVms(),
                getVmAdapters: vmName => mainViewModel.GetVmNetworkAdaptersForTrayAsync(vmName),
                getSwitches: () => mainViewModel.GetTraySwitches(),
                getUsbDevices: () => mainViewModel.GetUsbDevicesForTray(),
                selectUsbDeviceAction: busId => mainViewModel.SelectUsbDeviceForTrayAsync(busId),
                isTrayMenuEnabled: () => mainViewModel.UiEnableTrayMenu,
                refreshTrayDataAction: () => mainViewModel.RefreshTrayDataAsync(),
                startVmAction: vmName => mainViewModel.StartVmFromTrayAsync(vmName),
                stopVmAction: vmName => mainViewModel.StopVmFromTrayAsync(vmName),
                restartVmAction: vmName => mainViewModel.RestartVmByNameCommand.ExecuteAsync(vmName),
                openConsoleAction: vmName => mainViewModel.OpenConsoleFromTrayAsync(vmName),
                createSnapshotAction: vmName => mainViewModel.CreateSnapshotFromTrayAsync(vmName),
                connectVmToSwitchAction: (vmName, switchName, adapterName) => mainViewModel.ConnectVmSwitchFromTrayAsync(vmName, switchName, adapterName),
                getSelectedUsbDevice: () => mainViewModel.GetSelectedUsbDeviceForTray(),
                refreshUsbDevicesAction: () => mainViewModel.RefreshUsbDevicesFromTrayAsync(),
                shareSelectedUsbAction: () => mainViewModel.ShareSelectedUsbFromTrayAsync(),
                unshareSelectedUsbAction: () => mainViewModel.UnshareSelectedUsbFromTrayAsync(),
                exitAction: requestExit);

            _trayService = new HyperTool.WinUI.Services.TrayService();
            _trayService.Initialize(
                showAction: () =>
                {
                    mainWindow.AppWindow.Show();
                    mainWindow.Activate();
                },
                hideAction: () => mainWindow.AppWindow.Hide(),
                toggleControlCenterAction: () => _trayControlCenterService?.ToggleFull(),
                toggleControlCenterCompactAction: () => _trayControlCenterService?.ToggleCompact(),
                hideControlCenterAction: () => _trayControlCenterService?.Hide(),
                getUiTheme: () => mainViewModel.UiTheme,
                isWindowVisible: () =>
                {
                    try
                    {
                        return mainWindow.AppWindow?.IsVisible ?? true;
                    }
                    catch
                    {
                        return true;
                    }
                },
                isTrayMenuEnabled: () => mainViewModel.UiEnableTrayMenu,
                getVms: () => mainViewModel.GetTrayVms(),
                getSwitches: () => mainViewModel.GetTraySwitches(),
                refreshTrayDataAction: () => mainViewModel.RefreshTrayDataAsync(),
                subscribeTrayStateChanged: handler => mainViewModel.TrayStateChanged += handler,
                unsubscribeTrayStateChanged: handler => mainViewModel.TrayStateChanged -= handler,
                startVmAction: vmName => mainViewModel.StartVmFromTrayAsync(vmName),
                stopVmAction: vmName => mainViewModel.StopVmFromTrayAsync(vmName),
                restartVmAction: vmName => mainViewModel.RestartVmByNameCommand.ExecuteAsync(vmName),
                openConsoleAction: vmName => mainViewModel.OpenConsoleFromTrayAsync(vmName),
                createSnapshotAction: vmName => mainViewModel.CreateSnapshotFromTrayAsync(vmName),
                connectVmToSwitchAction: (vmName, switchName) => mainViewModel.ConnectVmSwitchFromTrayAsync(vmName, switchName),
                disconnectVmSwitchAction: vmName => mainViewModel.DisconnectVmSwitchFromTrayAsync(vmName),
                getSelectedUsbDevice: () => mainViewModel.GetSelectedUsbDeviceForTray(),
                refreshUsbDevicesAction: () => mainViewModel.RefreshUsbDevicesFromTrayAsync(),
                shareSelectedUsbAction: () => mainViewModel.ShareSelectedUsbFromTrayAsync(),
                unshareSelectedUsbAction: () => mainViewModel.UnshareSelectedUsbFromTrayAsync(),
                exitAction: requestExit);

            _isTrayFunctional = true;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Tray initialization failed. App continues without tray icon.");
        }
    }

    private async Task RunExitSequenceAsync(MainWindow mainWindow)
    {
        if (_isExitSequenceRunning)
        {
            return;
        }

        _isExitSequenceRunning = true;
        _isExitRequested = true;

        try
        {
            PointInt32 exitPosition;
            SizeInt32 exitSize;
            try
            {
                exitPosition = mainWindow.AppWindow.Position;
                exitSize = mainWindow.AppWindow.Size;
            }
            catch
            {
                exitPosition = new PointInt32(120, 120);
                exitSize = new SizeInt32(980, 640);
            }

            try
            {
                await mainWindow.PlayExitFadeAsync();
            }
            catch
            {
            }

            try
            {
                mainWindow.AppWindow.Hide();
            }
            catch
            {
            }

            await ExecuteUsbShutdownCleanupAsync();

            _trayControlCenterService?.Dispose();
            _trayControlCenterService = null;
            _trayService?.Dispose();
            _trayService = null;

            var exitWindow = new ExitScreenWindow();
            exitWindow.ConfigureBounds(exitPosition, exitSize);
            exitWindow.Activate();
            await exitWindow.PlayAndCloseAsync();
        }
        catch
        {
        }
        finally
        {
            try
            {
                mainWindow.Close();
            }
            catch
            {
            }

            _isExitSequenceRunning = false;
        }
    }

    private async Task ExecuteUsbShutdownCleanupAsync()
    {
        if (_usbShutdownCleanupDone)
        {
            return;
        }

        _usbShutdownCleanupDone = true;

        if (_mainViewModel is null)
        {
            return;
        }

        try
        {
            await _mainViewModel.UnshareAllSharedUsbOnShutdownAsync();
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "USB cleanup on shutdown failed.");
        }
    }

    private static string InitializeLogging()
    {
        var logDirectoryCandidates = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "HyperTool", "logs"),
            Path.Combine(AppContext.BaseDirectory, "logs"),
            Path.Combine(Path.GetTempPath(), "HyperTool", "logs")
        };

        var logsDirectory = logDirectoryCandidates.FirstOrDefault(IsWritableDirectory)
            ?? throw new InvalidOperationException("Kein beschreibbares Logverzeichnis gefunden.");

        var logFilePath = Path.Combine(logsDirectory, "hypertool-.log");

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.File(
                logFilePath,
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 14)
            .CreateLogger();

        return logFilePath;
    }

    private static bool IsWritableDirectory(string directoryPath)
    {
        try
        {
            Directory.CreateDirectory(directoryPath);

            var probePath = Path.Combine(directoryPath, $".write-test-{Guid.NewGuid():N}.tmp");
            using (File.Create(probePath))
            {
            }

            File.Delete(probePath);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private void RegisterGlobalExceptionHandlers()
    {
        AppDomain.CurrentDomain.UnhandledException += (_, eventArgs) =>
        {
            var exception = eventArgs.ExceptionObject as Exception ?? new Exception("Unknown unhandled exception");
            ShowFatalErrorAndExit(exception, "Ein kritischer Fehler ist aufgetreten.");
        };
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
        catch (Exception ex)
        {
            Log.Warning(ex, "Single-instance initialization failed. Continuing without strict single-instance lock.");
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

                _ = _mainWindow.DispatcherQueue.TryEnqueue(() => BringMainWindowToFront());
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Single-instance activation listener error.");
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

    private void ShowFatalErrorAndExit(Exception exception, string userMessage)
    {
        if (_isFatalErrorShown)
        {
            return;
        }

        _isFatalErrorShown = true;

        try
        {
            Log.Fatal(exception, "Fatal startup/runtime error");
            Log.CloseAndFlush();
            WriteCrashDump(exception);
        }
        catch
        {
        }

        NativeMessageBox.Show(
            $"{userMessage}{Environment.NewLine}{Environment.NewLine}{exception.Message}{Environment.NewLine}{Environment.NewLine}Details:{Environment.NewLine}{exception}",
            "HyperTool - Fataler Fehler",
            NativeMessageBoxButtons.Ok,
            NativeMessageBoxIcon.Error);

        ShutdownSingleInstanceInfrastructure();
        Microsoft.UI.Xaml.Application.Current.Exit();
    }

    private static void WriteCrashDump(Exception exception)
    {
        try
        {
            var crashDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "HyperTool",
                "crash");

            Directory.CreateDirectory(crashDir);
            var crashFile = Path.Combine(crashDir, $"crash-{DateTime.Now:yyyyMMdd-HHmmss}.txt");
            File.WriteAllText(crashFile, exception.ToString());
        }
        catch
        {
        }
    }

    public async Task RestartApplicationAsync(string? reason = null)
    {
        if (_isRestartInProgress)
        {
            return;
        }

        _isRestartInProgress = true;

        try
        {
            var executablePath = Environment.ProcessPath;
            if (string.IsNullOrWhiteSpace(executablePath))
            {
                executablePath = Process.GetCurrentProcess().MainModule?.FileName;
            }

            if (string.IsNullOrWhiteSpace(executablePath))
            {
                Log.Warning("Restart requested but executable path could not be determined. Reason: {Reason}", reason ?? "Unknown");
                return;
            }

            Log.Information("Restarting HyperTool. Reason: {Reason}", reason ?? "Unknown");

            Process.Start(new ProcessStartInfo
            {
                FileName = executablePath,
                UseShellExecute = true,
                WorkingDirectory = AppContext.BaseDirectory
            });

            _isExitRequested = true;

            await Task.Delay(120);

            try
            {
                _mainWindow?.Close();
            }
            catch
            {
            }

            Microsoft.UI.Xaml.Application.Current.Exit();
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to restart HyperTool.");
        }
        finally
        {
            _isRestartInProgress = false;
        }
    }

    private void RunRestartHnsHelperMode()
    {
        try
        {
            InitializeLogging();
        }
        catch
        {
        }

        var (success, message) = ExecuteHnsRestart();

        if (success)
        {
            NativeMessageBox.Show(
                "HNS-Dienst wurde erfolgreich neu gestartet.",
                "HyperTool Elevated Helper",
                NativeMessageBoxButtons.Ok,
                NativeMessageBoxIcon.Information);
            Microsoft.UI.Xaml.Application.Current.Exit();
            return;
        }

        NativeMessageBox.Show(
            message,
            "Fehler beim Neustart des HNS-Dienstes",
            NativeMessageBoxButtons.Ok,
            NativeMessageBoxIcon.Error);
        Microsoft.UI.Xaml.Application.Current.Exit();
    }

    private static (bool Success, string Message) ExecuteHnsRestart()
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoProfile -ExecutionPolicy Bypass -Command \"Restart-Service hns -Force -ErrorAction Stop\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardError = true,
                RedirectStandardOutput = true
            };

            using var process = Process.Start(psi);
            if (process is null)
            {
                return (false, "PowerShell konnte nicht gestartet werden.");
            }

            process.WaitForExit();
            var stdErr = process.StandardError.ReadToEnd();
            var stdOut = process.StandardOutput.ReadToEnd();

            if (process.ExitCode == 0)
            {
                Log.Information("HNS restart succeeded. {Output}", stdOut);
                return (true, "OK");
            }

            var errorText = string.IsNullOrWhiteSpace(stdErr) ? stdOut : stdErr;
            return (false, string.IsNullOrWhiteSpace(errorText) ? $"ExitCode {process.ExitCode}" : errorText.Trim());
        }
        catch (Exception ex)
        {
            Log.Error(ex, "HNS restart helper failed");
            return (false, ex.Message);
        }
    }

    private const int SwShow = 5;
    private const int SwRestore = 9;

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(nint hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(nint hWnd);

    [DllImport("user32.dll")]
    private static extern bool IsIconic(nint hWnd);
}
