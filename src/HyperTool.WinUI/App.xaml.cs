using HyperTool.Services;
using HyperTool.ViewModels;
using HyperTool.WinUI.Helpers;
using HyperTool.WinUI.Services;
using HyperTool.WinUI.Views;
using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.Win32;
using Serilog;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text.Json;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Windows.Graphics;

namespace HyperTool.WinUI;

public sealed partial class App : Application
{
    private const int HostUsbAutoRefreshSeconds = 5;
    private static readonly TimeSpan LogRetentionPeriod = TimeSpan.FromDays(3);
    private const string SingleInstanceMutexName = @"Local\HyperTool.WinUI.SingleInstance";
    private const string SingleInstancePipeName = "HyperTool.WinUI.SingleInstance.Activate";
    private static readonly (string ServiceId, string ElementName)[] RequiredHyperVSocketServices =
    [
        (HyperVSocketUsbTunnelDefaults.ServiceIdString, "HyperTool Hyper-V Socket USB Tunnel"),
        (HyperVSocketUsbTunnelDefaults.DiagnosticsServiceIdString, "HyperTool Hyper-V Socket Diagnostics"),
        (HyperVSocketUsbTunnelDefaults.SharedFolderCatalogServiceIdString, "HyperTool Hyper-V Socket Shared Folder Catalog"),
        (HyperVSocketUsbTunnelDefaults.HostIdentityServiceIdString, "HyperTool Hyper-V Socket Host Identity"),
        (HyperVSocketUsbTunnelDefaults.FileServiceIdString, "HyperTool Hyper-V Socket File Service")
    ];

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
    private HyperVSocketUsbHostTunnel? _usbHostTunnel;
    private HyperVSocketDiagnosticsHostListener? _usbDiagnosticsHostListener;
    private HyperVSocketSharedFolderCatalogHostListener? _sharedFolderCatalogHostListener;
    private HyperVSocketHostIdentityHostListener? _hostIdentityHostListener;
    private HyperVSocketFileHostListener? _fileHostListener;
    private CancellationTokenSource? _usbDiagnosticsCts;
    private Task? _usbDiagnosticsTask;
    private CancellationTokenSource? _usbHostDiscoveryCts;
    private Task? _usbHostDiscoveryTask;
    private CancellationTokenSource? _usbAutoRefreshCts;
    private Task? _usbAutoRefreshTask;
    private CancellationTokenSource? _usbEventRefreshCts;
    private string _lastMissingHyperVSocketServicesLogKey = string.Empty;
    private bool _hyperVSocketRegistrationPromptIssued;
    private bool _sharedFolderFileServiceSocketActive;
    private DateTimeOffset? _sharedFolderFileServiceLastActivityUtc;

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

        var rawLaunchArgs = args.Arguments;
        if (string.IsNullOrWhiteSpace(rawLaunchArgs))
        {
            try
            {
                var commandLineArgs = Environment.GetCommandLineArgs();
                if (commandLineArgs.Length > 1)
                {
                    rawLaunchArgs = string.Join(' ', commandLineArgs.Skip(1));
                }
            }
            catch
            {
            }
        }

        if (rawLaunchArgs?.Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Any(arg => string.Equals(arg, "--restart-hns", StringComparison.OrdinalIgnoreCase)) == true)
        {
            RunRestartHnsHelperMode();
            return;
        }

        if (rawLaunchArgs?.Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Any(arg => string.Equals(arg, "--register-sharedfolder-socket", StringComparison.OrdinalIgnoreCase)) == true)
        {
            RunRegisterSharedFolderSocketHelperMode();
            return;
        }

        if (rawLaunchArgs?.Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Any(arg => string.Equals(arg, "--usbipd-elevated-worker", StringComparison.OrdinalIgnoreCase)) == true)
        {
            if (TryGetLaunchArgumentValue(rawLaunchArgs, "--pipe", out var pipeName)
                && TryGetLaunchArgumentValue(rawLaunchArgs, "--token", out var token))
            {
                RunUsbipdElevatedWorkerMode(pipeName, token);
            }
            else
            {
                Environment.ExitCode = 2;
                Microsoft.UI.Xaml.Application.Current.Exit();
            }

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

            EnsureHyperVSocketServiceRegistrationsAtStartup();

            try
            {
                _usbHostTunnel = new HyperVSocketUsbHostTunnel();
                _usbHostTunnel.Start();
                Log.Information("Hyper-V socket USB host tunnel started.");

                _usbDiagnosticsHostListener = new HyperVSocketDiagnosticsHostListener(ack =>
                {
                    UsbGuestConnectionRegistry.UpdateFromDiagnosticsAck(ack);

                    if (_mainViewModel is not null
                        && !string.IsNullOrWhiteSpace(ack.BusId)
                        && string.Equals(ack.EventType, "usb-disconnected", StringComparison.OrdinalIgnoreCase))
                    {
                        _ = HandleUsbClientDisconnectEventAsync(ack.BusId);
                    }

                    if (_mainViewModel is not null
                        && !string.IsNullOrWhiteSpace(ack.BusId)
                        && (string.Equals(ack.EventType, "usb-connected", StringComparison.OrdinalIgnoreCase)
                            || string.Equals(ack.EventType, "usb-disconnected", StringComparison.OrdinalIgnoreCase)))
                    {
                        TriggerHostUsbRefreshForDiagnosticsEvent();
                    }

                    Log.Information(
                    "Hyper-V socket diagnostics acknowledged. EventType={EventType}; BusId={BusId}; GuestComputerName={GuestComputerName}; HostComputerName={HostComputerName}; GuestHyperVSocketActive={GuestHyperVSocketActive}; GuestRegistryServiceOk={GuestRegistryServiceOk}; GuestSentAtUtc={GuestSentAtUtc}; UsbTunnelActive={UsbTunnelActive}; RegistryServiceOk={RegistryServiceOk}",
                    ack.EventType,
                    ack.BusId,
                    ack.GuestComputerName,
                        Environment.MachineName,
                    ack.HyperVSocketActive,
                    ack.RegistryServiceOk,
                    ack.SentAtUtc,
                        _usbHostTunnel?.IsRunning == true,
                        HyperVSocketUsbHostTunnel.IsServiceRegistered());
                });
                _usbDiagnosticsHostListener.Start();
                Log.Information("Hyper-V socket diagnostics listener started.");
            }
            catch (Exception ex)
            {
                _usbDiagnosticsHostListener?.Dispose();
                _usbDiagnosticsHostListener = null;
                _usbHostTunnel?.Dispose();
                _usbHostTunnel = null;
                Log.Warning(ex, "Hyper-V socket USB host tunnel could not be started. Falling back to network-only USB transport.");
            }

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

            StartSharedFolderCatalogListenerWithRecovery();
            StartHostIdentityListenerWithRecovery();
            StartFileHostListenerWithRecovery();

            _mainWindow = new MainWindow(_themeService, _mainViewModel, showStartupSplash: true);
            AttachMainWindowHandlers(_mainWindow);
            UpdateSharedFolderFileServiceStatusPanel();
            StartUsbDiagnosticsLoop();
            StartUsbHostDiscoveryResponder();
            StartUsbAutoRefreshLoop();

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

            _ = RefreshHostUsbDevicesSafeAsync();

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
            StopUsbDiagnosticsLoop();
            StopUsbHostDiscoveryResponder();
            StopUsbAutoRefreshLoop();
            StopUsbEventRefreshScheduling();
            _usbDiagnosticsHostListener?.Dispose();
            _usbDiagnosticsHostListener = null;
            _usbHostTunnel?.Dispose();
            _usbHostTunnel = null;
            _sharedFolderCatalogHostListener?.Dispose();
            _sharedFolderCatalogHostListener = null;
            _hostIdentityHostListener?.Dispose();
            _hostIdentityHostListener = null;
            _fileHostListener?.Dispose();
            _fileHostListener = null;
            _sharedFolderFileServiceSocketActive = false;
            _sharedFolderFileServiceLastActivityUtc = null;
            UpdateSharedFolderFileServiceStatusPanel();
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
            StopUsbDiagnosticsLoop();
            StopUsbHostDiscoveryResponder();
            StopUsbAutoRefreshLoop();
            StopUsbEventRefreshScheduling();
            _usbDiagnosticsHostListener?.Dispose();
            _usbDiagnosticsHostListener = null;
            _usbHostTunnel?.Dispose();
            _usbHostTunnel = null;
            _sharedFolderCatalogHostListener?.Dispose();
            _sharedFolderCatalogHostListener = null;
            _hostIdentityHostListener?.Dispose();
            _hostIdentityHostListener = null;
            _fileHostListener?.Dispose();
            _fileHostListener = null;
            _sharedFolderFileServiceSocketActive = false;
            _sharedFolderFileServiceLastActivityUtc = null;
            UpdateSharedFolderFileServiceStatusPanel();

            _themeService.ApplyTheme(_mainViewModel.UiTheme);

            var nextWindow = new MainWindow(_themeService, _mainViewModel, showStartupSplash: false);
            await nextWindow.ShowLifecycleGuardAsync("Design wird neu geladen …");
            AttachMainWindowHandlers(nextWindow);
            _mainWindow = nextWindow;

            TryInitializeTray(nextWindow, _mainViewModel);

            try
            {
                _usbHostTunnel = new HyperVSocketUsbHostTunnel();
                _usbHostTunnel.Start();
                Log.Information("Hyper-V socket USB host tunnel restarted after theme change.");

                _usbDiagnosticsHostListener = new HyperVSocketDiagnosticsHostListener(ack =>
                {
                    Log.Information(
                    "Hyper-V socket diagnostics test acknowledged. GuestComputerName={GuestComputerName}; HostComputerName={HostComputerName}; GuestHyperVSocketActive={GuestHyperVSocketActive}; GuestRegistryServiceOk={GuestRegistryServiceOk}; GuestSentAtUtc={GuestSentAtUtc}; UsbTunnelActive={UsbTunnelActive}; RegistryServiceOk={RegistryServiceOk}",
                    ack.GuestComputerName,
                        Environment.MachineName,
                    ack.HyperVSocketActive,
                    ack.RegistryServiceOk,
                    ack.SentAtUtc,
                        _usbHostTunnel?.IsRunning == true,
                        HyperVSocketUsbHostTunnel.IsServiceRegistered());
                });
                _usbDiagnosticsHostListener.Start();
                Log.Information("Hyper-V socket diagnostics listener restarted after theme change.");
            }
            catch (Exception ex)
            {
                _usbDiagnosticsHostListener?.Dispose();
                _usbDiagnosticsHostListener = null;
                _usbHostTunnel?.Dispose();
                _usbHostTunnel = null;
                Log.Warning(ex, "Hyper-V socket services could not be restarted after theme change.");
            }

            StartSharedFolderCatalogListenerWithRecovery(isThemeRestart: true);
            StartHostIdentityListenerWithRecovery(isThemeRestart: true);
            StartFileHostListenerWithRecovery(isThemeRestart: true);

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
                await nextWindow.HideLifecycleGuardAsync();
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

                await nextWindow.HideLifecycleGuardAsync();
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
            var exitAnimationStopwatch = Stopwatch.StartNew();
            const int inlineExitAnimationDurationMs = 2000;

            try
            {
                await mainWindow.ShowLifecycleGuardAsync("Beende HyperTool …");
            }
            catch
            {
            }

            await ExecuteUsbShutdownCleanupAsync();

            _trayControlCenterService?.Dispose();
            _trayControlCenterService = null;
            _trayService?.Dispose();
            _trayService = null;

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

            try
            {
                mainWindow.AppWindow.Hide();
            }
            catch
            {
            }
        }
        catch
        {
        }
        finally
        {
            try
            {
                mainWindow.CloseAuxiliaryWindows();
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

        if (!_mainViewModel.UsbUnshareOnExit)
        {
            Log.Information("USB cleanup on shutdown skipped by setting (UsbUnshareOnExit=false).");
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

    private void StartUsbAutoRefreshLoop()
    {
        StopUsbAutoRefreshLoop();

        var cts = new CancellationTokenSource();
        _usbAutoRefreshCts = cts;
        _usbAutoRefreshTask = Task.Run(() => RunUsbAutoRefreshLoopAsync(cts.Token));
    }

    private void StartUsbHostDiscoveryResponder()
    {
        StopUsbHostDiscoveryResponder();

        var cts = new CancellationTokenSource();
        _usbHostDiscoveryCts = cts;
        _usbHostDiscoveryTask = Task.Run(() => RunUsbHostDiscoveryResponderAsync(cts.Token));
    }

    private void StopUsbHostDiscoveryResponder()
    {
        try
        {
            _usbHostDiscoveryCts?.Cancel();
        }
        catch
        {
        }

        _usbHostDiscoveryCts?.Dispose();
        _usbHostDiscoveryCts = null;
        _usbHostDiscoveryTask = null;
    }

    private async Task RunUsbHostDiscoveryResponderAsync(CancellationToken cancellationToken)
    {
        try
        {
            await UsbHostDiscoveryService.RunHostResponderAsync(
                hostComputerName: Environment.MachineName,
                getHostAddresses: () => UsbHostDiscoveryService.GetLocalIpv4Addresses(),
                cancellationToken: cancellationToken);
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "USB host discovery responder failed.");
        }
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

    private void StopUsbEventRefreshScheduling()
    {
        try
        {
            _usbEventRefreshCts?.Cancel();
        }
        catch
        {
        }

        _usbEventRefreshCts?.Dispose();
        _usbEventRefreshCts = null;
    }

    private void TriggerHostUsbRefreshForDiagnosticsEvent()
    {
        _ = RefreshHostUsbDevicesSafeAsync();

        try
        {
            _usbEventRefreshCts?.Cancel();
        }
        catch
        {
        }

        _usbEventRefreshCts?.Dispose();
        _usbEventRefreshCts = new CancellationTokenSource();
        var refreshToken = _usbEventRefreshCts.Token;

        _ = Task.Run(async () =>
        {
            try
            {
                await Task.Delay(TimeSpan.FromSeconds(2), refreshToken);
                await RefreshHostUsbDevicesSafeAsync();
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Deferred host USB refresh after diagnostics event failed.");
            }
        }, refreshToken);
    }

    private async Task RefreshHostUsbDevicesSafeAsync()
    {
        if (_mainViewModel is null
            || _isExitRequested
            || _isThemeWindowReopenInProgress)
        {
            return;
        }

        if (_mainWindow?.DispatcherQueue is { } queue && !queue.HasThreadAccess)
        {
            var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            if (!queue.TryEnqueue(async () =>
                {
                    try
                    {
                        if (_mainViewModel is null)
                        {
                            completion.TrySetResult(true);
                            return;
                        }

                        await _mainViewModel.RefreshUsbDevicesFromTrayAsync();
                        completion.TrySetResult(true);
                    }
                    catch (Exception ex)
                    {
                        completion.TrySetException(ex);
                    }
                }))
            {
                return;
            }

            await completion.Task;
            return;
        }

        await _mainViewModel.RefreshUsbDevicesFromTrayAsync();
    }

    private async Task HandleUsbClientDisconnectEventAsync(string busId)
    {
        if (_mainViewModel is null
            || _isExitRequested
            || _isThemeWindowReopenInProgress
            || string.IsNullOrWhiteSpace(busId))
        {
            return;
        }

        try
        {
            await _mainViewModel.HandleUsbClientDisconnectedAsync(busId);
        }
        catch (Exception ex)
        {
            Log.Debug(ex, "Automatic host detach on usb-disconnected event failed. BusId={BusId}", busId);
        }
    }

    private async Task RunUsbAutoRefreshLoopAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                if (_mainViewModel is not null
                    && !_isExitRequested
                    && !_isThemeWindowReopenInProgress
                    && !_mainViewModel.IsBusy
                    && _mainViewModel.HasUsbAutoShareConfigured)
                {
                    await _mainViewModel.RefreshUsbDevicesFromTrayAsync();
                }
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "USB auto refresh loop iteration failed.");
            }

            try
            {
                await Task.Delay(TimeSpan.FromSeconds(HostUsbAutoRefreshSeconds), cancellationToken);
            }
            catch (OperationCanceledException)
            {
                break;
            }
        }
    }

    private async Task RunUsbDiagnosticsLoopAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                UpdateHostUsbDiagnostics();
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

    private void UpdateHostUsbDiagnostics()
    {
        if (_mainViewModel is null)
        {
            return;
        }

        var hyperVActiveText = _usbHostTunnel?.IsRunning == true ? "Ja" : "Nein";
        var missingServiceIds = GetMissingHyperVSocketServiceIds();
        var registryText = missingServiceIds.Count == 0 ? "Ja" : "Nein";
        const string fallbackText = "Nein (Host ist Quelle)";

        var missingServicesLogKey = string.Join("|", missingServiceIds.OrderBy(static id => id, StringComparer.OrdinalIgnoreCase));
        if (!string.Equals(_lastMissingHyperVSocketServicesLogKey, missingServicesLogKey, StringComparison.Ordinal))
        {
            _lastMissingHyperVSocketServicesLogKey = missingServicesLogKey;
            if (missingServiceIds.Count > 0)
            {
                Log.Warning("Registry service check: missing Hyper-V socket service IDs: {ServiceIds}", string.Join(", ", missingServiceIds));
            }
            else
            {
                Log.Information("Registry service check: all Hyper-V socket service IDs are present.");
            }
        }

        void apply()
        {
            if (_mainViewModel is null)
            {
                return;
            }

            _mainViewModel.UsbDiagnosticsHyperVSocketText = hyperVActiveText;
            _mainViewModel.UsbDiagnosticsRegistryServiceText = registryText;
            _mainViewModel.UsbDiagnosticsFallbackText = fallbackText;
        }

        if (_mainWindow?.DispatcherQueue is { } queue && !queue.HasThreadAccess)
        {
            queue.TryEnqueue(apply);
            return;
        }

        apply();
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

        CleanupOldLogFiles(logsDirectory, LogRetentionPeriod);

        var logFilePath = Path.Combine(logsDirectory, "hypertool-.log");

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.File(
                logFilePath,
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 14,
                encoding: new UTF8Encoding(encoderShouldEmitUTF8Identifier: true))
            .CreateLogger();

        return logFilePath;
    }

    private static void CleanupOldLogFiles(string directoryPath, TimeSpan maxAge)
    {
        try
        {
            if (!Directory.Exists(directoryPath))
            {
                return;
            }

            var cutoffUtc = DateTime.UtcNow - maxAge;
            foreach (var filePath in Directory.EnumerateFiles(directoryPath, "*", SearchOption.TopDirectoryOnly))
            {
                try
                {
                    if (File.GetLastWriteTimeUtc(filePath) < cutoffUtc)
                    {
                        File.Delete(filePath);
                    }
                }
                catch
                {
                }
            }
        }
        catch
        {
        }
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

    private void RunRegisterSharedFolderSocketHelperMode()
    {
        try
        {
            InitializeLogging();
        }
        catch
        {
        }

        var (success, message) = ExecuteSharedFolderSocketRegistration();
        if (!success)
        {
            Log.Error("Shared-folder socket service registration failed: {Message}", message);
        }

        Environment.ExitCode = success ? 0 : 1;
        Microsoft.UI.Xaml.Application.Current.Exit();
    }

    private void RunUsbipdElevatedWorkerMode(string pipeName, string token)
    {
        try
        {
            InitializeLogging();
        }
        catch
        {
        }

        var exitCode = ExecuteUsbipdElevatedWorker(pipeName, token);
        Environment.ExitCode = exitCode;
        Microsoft.UI.Xaml.Application.Current.Exit();
    }

    private static int ExecuteUsbipdElevatedWorker(string pipeName, string token)
    {
        if (string.IsNullOrWhiteSpace(pipeName) || string.IsNullOrWhiteSpace(token))
        {
            return 2;
        }

        try
        {
            using var pipe = new NamedPipeClientStream(
                ".",
                pipeName,
                PipeDirection.InOut,
                PipeOptions.Asynchronous);

            pipe.Connect(120000);

            using var reader = new StreamReader(pipe, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
            using var writer = new StreamWriter(pipe, Encoding.UTF8, bufferSize: 4096, leaveOpen: true)
            {
                AutoFlush = true
            };

            writer.WriteLine("READY");

            while (true)
            {
                var requestLine = reader.ReadLine();
                if (requestLine is null)
                {
                    break;
                }

                UsbElevatedWorkerRequest? request;
                try
                {
                    request = JsonSerializer.Deserialize<UsbElevatedWorkerRequest>(requestLine);
                }
                catch
                {
                    var invalidJson = new UsbElevatedWorkerResponse
                    {
                        ExitCode = 1,
                        StandardOutput = string.Empty,
                        StandardError = "Ungültiges Request-Format.",
                        ContinueRunning = true
                    };
                    writer.WriteLine(JsonSerializer.Serialize(invalidJson));
                    continue;
                }

                if (request is null || !string.Equals(request.Token, token, StringComparison.Ordinal))
                {
                    var unauthorized = new UsbElevatedWorkerResponse
                    {
                        ExitCode = 5,
                        StandardOutput = string.Empty,
                        StandardError = "Unauthorized.",
                        ContinueRunning = false
                    };
                    writer.WriteLine(JsonSerializer.Serialize(unauthorized));
                    return 5;
                }

                if (string.Equals(request.Command, "shutdown", StringComparison.OrdinalIgnoreCase))
                {
                    var shutdown = new UsbElevatedWorkerResponse
                    {
                        ExitCode = 0,
                        StandardOutput = string.Empty,
                        StandardError = string.Empty,
                        ContinueRunning = false
                    };
                    writer.WriteLine(JsonSerializer.Serialize(shutdown));
                    return 0;
                }

                var commandResult = ExecuteUsbipdWorkerCommand(request.Command, request.Arguments);
                writer.WriteLine(JsonSerializer.Serialize(commandResult));

                if (!commandResult.ContinueRunning)
                {
                    return commandResult.ExitCode;
                }
            }

            return 0;
        }
        catch (Exception ex)
        {
            try
            {
                Log.Error(ex, "USB elevated worker failed.");
            }
            catch
            {
            }

            return 1;
        }
    }

    private static UsbElevatedWorkerResponse ExecuteUsbipdWorkerCommand(string? command, IReadOnlyList<string>? arguments)
    {
        if (!string.Equals(command, "usbipd", StringComparison.OrdinalIgnoreCase))
        {
            return new UsbElevatedWorkerResponse
            {
                ExitCode = 1,
                StandardOutput = string.Empty,
                StandardError = "Unbekannter Befehl.",
                ContinueRunning = true
            };
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = ResolveUsbipdExecutablePath(),
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            CreateNoWindow = true
        };

        if (arguments is not null)
        {
            foreach (var argument in arguments)
            {
                if (string.IsNullOrWhiteSpace(argument))
                {
                    continue;
                }

                startInfo.ArgumentList.Add(argument);
            }
        }

        try
        {
            using var process = Process.Start(startInfo)
                                ?? throw new InvalidOperationException("usbipd konnte nicht gestartet werden.");

            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();
            process.WaitForExit();

            return new UsbElevatedWorkerResponse
            {
                ExitCode = process.ExitCode,
                StandardOutput = stdout,
                StandardError = stderr,
                ContinueRunning = true
            };
        }
        catch (Exception ex)
        {
            return new UsbElevatedWorkerResponse
            {
                ExitCode = 1,
                StandardOutput = string.Empty,
                StandardError = ex.Message,
                ContinueRunning = true
            };
        }
    }

    private static string ResolveUsbipdExecutablePath()
    {
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var candidate = Path.Combine(programFiles, "usbipd-win", "usbipd.exe");
        if (File.Exists(candidate))
        {
            return candidate;
        }

        return "usbipd";
    }

    private sealed class UsbElevatedWorkerRequest
    {
        public string Token { get; init; } = string.Empty;

        public string Command { get; init; } = string.Empty;

        public IReadOnlyList<string> Arguments { get; init; } = [];
    }

    private sealed class UsbElevatedWorkerResponse
    {
        public int ExitCode { get; init; }

        public string StandardOutput { get; init; } = string.Empty;

        public string StandardError { get; init; } = string.Empty;

        public bool ContinueRunning { get; init; } = true;
    }

    private static bool TryGetLaunchArgumentValue(string? rawArguments, string key, out string value)
    {
        value = string.Empty;
        if (string.IsNullOrWhiteSpace(rawArguments) || string.IsNullOrWhiteSpace(key))
        {
            return false;
        }

        var args = rawArguments.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        for (var index = 0; index < args.Length; index++)
        {
            if (!string.Equals(args[index], key, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (index + 1 >= args.Length)
            {
                return false;
            }

            value = args[index + 1];
            return !string.IsNullOrWhiteSpace(value);
        }

        return false;
    }

    private static (bool Success, string Message) ExecuteSharedFolderSocketRegistration()
    {
        try
        {
            const string rootPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices";

            using var rootKey = Registry.LocalMachine.CreateSubKey(rootPath, writable: true);
            if (rootKey is null)
            {
                return (false, "Registry-Pfad für GuestCommunicationServices konnte nicht geöffnet werden.");
            }

            foreach (var (serviceId, elementName) in RequiredHyperVSocketServices)
            {
                using var serviceKey = rootKey.CreateSubKey(serviceId, writable: true);
                if (serviceKey is null)
                {
                    return (false, $"Registry-Eintrag für Hyper-V Socket Service {serviceId} konnte nicht erstellt werden.");
                }

                serviceKey.SetValue("ElementName", elementName, RegistryValueKind.String);
            }

            return (true, "OK");
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }

    private void EnsureHyperVSocketServiceRegistrationsAtStartup()
    {
        var missingServiceIds = GetMissingHyperVSocketServiceIds();
        if (missingServiceIds.Count == 0)
        {
            Log.Information("All required Hyper-V socket registry entries are present.");
            return;
        }

        Log.Warning("Missing Hyper-V socket registry entries detected at startup: {ServiceIds}", string.Join(", ", missingServiceIds));

        if (!TryRegisterSharedFolderSocketServiceElevated(allowPrompt: false))
        {
            Log.Warning("Could not auto-register missing Hyper-V socket registry entries at startup without prompt.");
            return;
        }

        var remainingMissing = GetMissingHyperVSocketServiceIds();
        if (remainingMissing.Count == 0)
        {
            Log.Information("Missing Hyper-V socket registry entries were successfully created by elevated helper.");
            return;
        }

        Log.Warning("Hyper-V socket registry entries are still missing after elevated helper: {ServiceIds}", string.Join(", ", remainingMissing));
    }

    private bool TryRegisterSharedFolderSocketServiceWithoutPrompt()
    {
        if (!IsRunningAsAdministrator())
        {
            return false;
        }

        var (success, message) = ExecuteSharedFolderSocketRegistration();
        if (!success)
        {
            Log.Warning("Silent Hyper-V socket registration failed: {Message}", message);
            return false;
        }

        Log.Information("Silent Hyper-V socket registration completed (process already elevated).");
        return true;
    }

    private static bool IsRunningAsAdministrator()
    {
        try
        {
            using var identity = WindowsIdentity.GetCurrent();
            return new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator);
        }
        catch
        {
            return false;
        }
    }

    private static List<string> GetMissingHyperVSocketServiceIds()
    {
        var missing = new List<string>();
        foreach (var (serviceId, _) in RequiredHyperVSocketServices)
        {
            if (!Guid.TryParse(serviceId, out var parsedServiceId)
                || !HyperVSocketUsbHostTunnel.IsServiceRegistered(parsedServiceId))
            {
                missing.Add(serviceId);
            }
        }

        return missing;
    }

    private void StartSharedFolderCatalogListenerWithRecovery(bool isThemeRestart = false)
    {
        if (_mainViewModel is null)
        {
            return;
        }

        try
        {
            _sharedFolderCatalogHostListener = new HyperVSocketSharedFolderCatalogHostListener(
                () => _mainViewModel.GetHostSharedFoldersSnapshot());
            _sharedFolderCatalogHostListener.Start();
            Log.Information(isThemeRestart
                ? "Hyper-V socket shared-folder catalog listener restarted after theme change."
                : "Hyper-V socket shared-folder catalog listener started.");
            return;
        }
        catch (Exception ex)
        {
            _sharedFolderCatalogHostListener?.Dispose();
            _sharedFolderCatalogHostListener = null;
            Log.Warning(ex, isThemeRestart
                ? "Hyper-V socket shared-folder catalog listener could not be restarted after theme change."
                : "Hyper-V socket shared-folder catalog listener could not be started.");

            if (!IsMissingSharedFolderSocketRegistration(ex))
            {
                return;
            }
        }

        if (!TryRegisterSharedFolderSocketServiceElevated(allowPrompt: false))
        {
            Log.Information("Hyper-V socket shared-folder catalog listener remains disabled until registry entries exist (startup prompt disabled).");
            return;
        }

        try
        {
            _sharedFolderCatalogHostListener = new HyperVSocketSharedFolderCatalogHostListener(
                () => _mainViewModel.GetHostSharedFoldersSnapshot());
            _sharedFolderCatalogHostListener.Start();
            Log.Information("Hyper-V socket shared-folder catalog listener started after elevated registration helper.");
        }
        catch (Exception ex)
        {
            _sharedFolderCatalogHostListener?.Dispose();
            _sharedFolderCatalogHostListener = null;
            Log.Warning(ex, "Hyper-V socket shared-folder catalog listener still unavailable after elevated registration helper.");
        }
    }

    private void StartHostIdentityListenerWithRecovery(bool isThemeRestart = false)
    {
        try
        {
            _hostIdentityHostListener = new HyperVSocketHostIdentityHostListener(
                featureAvailabilityProvider: () => new HyperTool.Models.HostFeatureAvailability
                {
                    UsbSharingEnabled = _mainViewModel?.HostUsbSharingEnabled ?? true,
                    SharedFoldersEnabled = _mainViewModel?.HostSharedFoldersEnabled ?? true
                });
            _hostIdentityHostListener.Start();
            Log.Information(isThemeRestart
                ? "Hyper-V socket host-identity listener restarted after theme change."
                : "Hyper-V socket host-identity listener started.");
            return;
        }
        catch (Exception ex)
        {
            _hostIdentityHostListener?.Dispose();
            _hostIdentityHostListener = null;
            Log.Warning(ex, isThemeRestart
                ? "Hyper-V socket host-identity listener could not be restarted after theme change."
                : "Hyper-V socket host-identity listener could not be started.");

            if (!IsMissingSharedFolderSocketRegistration(ex))
            {
                return;
            }
        }

        if (!TryRegisterSharedFolderSocketServiceElevated(allowPrompt: false))
        {
            Log.Information("Hyper-V socket host-identity listener remains disabled until registry entries exist (startup prompt disabled).");
            return;
        }

        try
        {
            _hostIdentityHostListener = new HyperVSocketHostIdentityHostListener(
                featureAvailabilityProvider: () => new HyperTool.Models.HostFeatureAvailability
                {
                    UsbSharingEnabled = _mainViewModel?.HostUsbSharingEnabled ?? true,
                    SharedFoldersEnabled = _mainViewModel?.HostSharedFoldersEnabled ?? true
                });
            _hostIdentityHostListener.Start();
            Log.Information("Hyper-V socket host-identity listener started after elevated registration helper.");
        }
        catch (Exception ex)
        {
            _hostIdentityHostListener?.Dispose();
            _hostIdentityHostListener = null;
            Log.Warning(ex, "Hyper-V socket host-identity listener still unavailable after elevated registration helper.");
        }
    }

    private void StartFileHostListenerWithRecovery(bool isThemeRestart = false)
    {
        if (_mainViewModel is null)
        {
            return;
        }

        try
        {
            _fileHostListener = new HyperVSocketFileHostListener(
                () => _mainViewModel.GetHostSharedFoldersSnapshot(),
                onRequestServed: OnSharedFolderFileServiceRequestServed);
            _fileHostListener.Start();
            _sharedFolderFileServiceSocketActive = true;
            UpdateSharedFolderFileServiceStatusPanel();
            Log.Information(isThemeRestart
                ? "Hyper-V socket file listener restarted after theme change."
                : "Hyper-V socket file listener started.");
            return;
        }
        catch (Exception ex)
        {
            _fileHostListener?.Dispose();
            _fileHostListener = null;
            _sharedFolderFileServiceSocketActive = false;
            UpdateSharedFolderFileServiceStatusPanel();
            Log.Warning(ex, isThemeRestart
                ? "Hyper-V socket file listener could not be restarted after theme change."
                : "Hyper-V socket file listener could not be started.");

            if (!IsMissingSharedFolderSocketRegistration(ex))
            {
                return;
            }
        }

        if (!TryRegisterSharedFolderSocketServiceElevated(allowPrompt: false))
        {
            _sharedFolderFileServiceSocketActive = false;
            UpdateSharedFolderFileServiceStatusPanel();
            Log.Information("Hyper-V socket file listener remains disabled until registry entries exist (startup prompt disabled).");
            return;
        }

        try
        {
            _fileHostListener = new HyperVSocketFileHostListener(
                () => _mainViewModel.GetHostSharedFoldersSnapshot(),
                onRequestServed: OnSharedFolderFileServiceRequestServed);
            _fileHostListener.Start();
            _sharedFolderFileServiceSocketActive = true;
            UpdateSharedFolderFileServiceStatusPanel();
            Log.Information("Hyper-V socket file listener started after elevated registration helper.");
        }
        catch (Exception ex)
        {
            _fileHostListener?.Dispose();
            _fileHostListener = null;
            _sharedFolderFileServiceSocketActive = false;
            UpdateSharedFolderFileServiceStatusPanel();
            Log.Warning(ex, "Hyper-V socket file listener still unavailable after elevated registration helper.");
        }
    }

    private void OnSharedFolderFileServiceRequestServed(DateTimeOffset servedAtUtc)
    {
        _sharedFolderFileServiceLastActivityUtc = servedAtUtc;
        _sharedFolderFileServiceSocketActive = _fileHostListener?.IsRunning == true;
        UpdateSharedFolderFileServiceStatusPanel();
    }

    private void UpdateSharedFolderFileServiceStatusPanel()
    {
        var socketActive = _sharedFolderFileServiceSocketActive;
        var lastActivityUtc = _sharedFolderFileServiceLastActivityUtc;

        void apply()
        {
            _mainWindow?.UpdateSharedFolderFileServiceStatus(socketActive, lastActivityUtc);
        }

        if (_mainWindow?.DispatcherQueue is { } queue && !queue.HasThreadAccess)
        {
            _ = queue.TryEnqueue(apply);
            return;
        }

        apply();
    }

    private static bool IsMissingSharedFolderSocketRegistration(Exception ex)
    {
        return ex is InvalidOperationException invalidOperation
               && invalidOperation.Message.Contains("nicht registriert", StringComparison.OrdinalIgnoreCase);
    }

    private bool TryRegisterSharedFolderSocketServiceElevated(bool allowPrompt = true)
    {
        if (TryRegisterSharedFolderSocketServiceWithoutPrompt())
        {
            return true;
        }

        if (!allowPrompt)
        {
            Log.Information("Skipping elevated Hyper-V socket registration prompt (prompt disabled for current flow).");
            return false;
        }

        if (_hyperVSocketRegistrationPromptIssued)
        {
            Log.Information("Skipping additional elevated Hyper-V socket registration prompt (already attempted in this app session).");
            return false;
        }

        _hyperVSocketRegistrationPromptIssued = true;

        string? scriptPath = null;

        try
        {
            scriptPath = Path.Combine(Path.GetTempPath(), $"HyperTool.RegisterHyperVSocket.{Guid.NewGuid():N}.ps1");
            File.WriteAllText(scriptPath, BuildElevatedHyperVSocketRegistrationScript(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{scriptPath}\"",
                UseShellExecute = true,
                Verb = "runas"
            });

            if (process is null)
            {
                Log.Warning("Elevated shared-folder registration helper could not be started.");
                return false;
            }

            process.WaitForExit();
            if (process.ExitCode == 0)
            {
                return true;
            }

            Log.Warning("Elevated shared-folder registration helper exited with code {ExitCode}.", process.ExitCode);
            return false;
        }
        catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
        {
            Log.Warning("UAC prompt for shared-folder socket registration was cancelled.");
            return false;
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Failed to start elevated shared-folder registration helper.");
            return false;
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(scriptPath))
            {
                try
                {
                    File.Delete(scriptPath);
                }
                catch
                {
                }
            }
        }
    }

    private static string BuildElevatedHyperVSocketRegistrationScript()
    {
        static string escape(string value) => value.Replace("'", "''", StringComparison.Ordinal);

        var rootPath = @"HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices";
        var sb = new StringBuilder();
        sb.AppendLine("$ErrorActionPreference = 'Stop'");
        sb.AppendLine($"$rootPath = '{escape(rootPath)}'");
        sb.AppendLine("New-Item -Path $rootPath -Force | Out-Null");

        foreach (var (serviceId, elementName) in RequiredHyperVSocketServices)
        {
            sb.AppendLine($"$servicePath = Join-Path $rootPath '{escape(serviceId)}'");
            sb.AppendLine("New-Item -Path $servicePath -Force | Out-Null");
            sb.AppendLine($"Set-ItemProperty -Path $servicePath -Name 'ElementName' -Type String -Value '{escape(elementName)}'");
        }

        return sb.ToString();
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
