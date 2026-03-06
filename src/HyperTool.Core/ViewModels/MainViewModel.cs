using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using HyperTool.Models;
using HyperTool.Services;
using Serilog;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace HyperTool.ViewModels;

public partial class MainViewModel : ViewModelBase
{
    private const string NotConnectedSwitchDisplay = "Nicht verbunden";
    private const int ConfigMenuIndex = 4;
    private const int VkNumLock = 0x90;
    private const uint KeyeventfExtendedKey = 0x0001;
    private const uint KeyeventfKeyUp = 0x0002;
    private static readonly TimeSpan StaleUsbAttachGracePeriod = TimeSpan.FromSeconds(25);

    [ObservableProperty]
    private string _windowTitle = "HyperTool";

    [ObservableProperty]
    private string _statusText = "Bereit";

    [ObservableProperty]
    private bool _isBusy;

    [ObservableProperty]
    private string? _configurationNotice;

    [ObservableProperty]
    private VmDefinition? _selectedVm;

    [ObservableProperty]
    private HyperVSwitchInfo? _selectedSwitch;

    [ObservableProperty]
    private HyperVVmNetworkAdapterInfo? _selectedVmNetworkAdapter;

    [ObservableProperty]
    private HyperVCheckpointInfo? _selectedCheckpoint;

    [ObservableProperty]
    private HyperVCheckpointTreeItem? _selectedCheckpointNode;

    [ObservableProperty]
    private string _newCheckpointName = string.Empty;

    [ObservableProperty]
    private string _newCheckpointDescription = string.Empty;

    [ObservableProperty]
    private string _selectedVmState = "Unbekannt";

    [ObservableProperty]
    private string _selectedVmCurrentSwitch = NotConnectedSwitchDisplay;

    [ObservableProperty]
    private string _busyText = "Bitte warten...";

    [ObservableProperty]
    private int _busyProgressPercent = -1;

    [ObservableProperty]
    private VmDefinition? _selectedVmForConfig;

    [ObservableProperty]
    private bool _selectedVmOpenConsoleWithSessionEdit;

    [ObservableProperty]
    private VmTrayAdapterOption? _selectedVmTrayAdapterOption;

    [ObservableProperty]
    private HyperVVmNetworkAdapterInfo? _selectedVmAdapterForRename;

    [ObservableProperty]
    private string _newVmAdapterName = string.Empty;

    [ObservableProperty]
    private string _vmAdapterRenameValidationMessage = string.Empty;

    [ObservableProperty]
    private VmDefinition? _selectedDefaultVmForConfig;

    [ObservableProperty]
    private string _newVmName = string.Empty;

    [ObservableProperty]
    private string _newVmLabel = string.Empty;

    [ObservableProperty]
    private string _importVmRequestedName = string.Empty;

    [ObservableProperty]
    private string _defaultVmImportDestinationPath = string.Empty;

    [ObservableProperty]
    private VmImportModeOption? _selectedVmImportModeOption;

    [ObservableProperty]
    private bool _hnsEnabled;

    [ObservableProperty]
    private bool _hnsAutoRestartAfterDefaultSwitch;

    [ObservableProperty]
    private bool _hnsAutoRestartAfterAnyConnect;

    [ObservableProperty]
    private string _defaultVmName = string.Empty;

    [ObservableProperty]
    private string _vmConnectComputerName = "";

    [ObservableProperty]
    private string _lastSelectedVmName = string.Empty;

    [ObservableProperty]
    private bool _uiEnableTrayIcon = true;

    [ObservableProperty]
    private bool _uiEnableTrayMenu = true;

    [ObservableProperty]
    private bool _uiStartMinimized;

    [ObservableProperty]
    private bool _uiStartWithWindows;

    [ObservableProperty]
    private bool _uiOpenConsoleAfterVmStart = true;

    [ObservableProperty]
    private bool _uiRestoreNumLockAfterVmStart;

    [ObservableProperty]
    private bool _uiOpenVmConnectWithSessionEdit;

    [DllImport("user32.dll")]
    private static extern short GetKeyState(int virtualKey);

    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    [ObservableProperty]
    private string _uiTheme = "Dark";

    [ObservableProperty]
    private bool _updateCheckOnStartup = true;

    [ObservableProperty]
    private string _githubOwner = "koerby";

    [ObservableProperty]
    private string _githubRepo = "HyperTool";

    [ObservableProperty]
    private string _appVersion = "0.0.0";

    [ObservableProperty]
    private string _updateStatus = "Noch nicht geprüft";

    [ObservableProperty]
    private string _releaseUrl = string.Empty;

    [ObservableProperty]
    private string _installerDownloadUrl = string.Empty;

    [ObservableProperty]
    private string _installerFileName = string.Empty;

    [ObservableProperty]
    private bool _updateInstallAvailable;

    [ObservableProperty]
    private int _selectedMenuIndex;

    [ObservableProperty]
    private bool _isLogExpanded;

    [ObservableProperty]
    private bool _areSwitchesLoaded;

    [ObservableProperty]
    private string _networkSwitchStatusHint = string.Empty;

    [ObservableProperty]
    private string _hostNetworkProfileCategory = "Unknown";

    [ObservableProperty]
    private string _hostNetworkProfileDisplayText = "Host-Netzprofil: Unbekannt";

    [ObservableProperty]
    private bool _hasPendingConfigChanges;

    [ObservableProperty]
    private UsbIpDeviceInfo? _selectedUsbDevice;

    [ObservableProperty]
    private string _usbStatusText = "Noch nicht geladen";

    [ObservableProperty]
    private bool _hostUsbSharingEnabled = true;

    [ObservableProperty]
    private bool _usbAutoDetachOnClientDisconnect = true;

    [ObservableProperty]
    private bool _usbUnshareOnExit = true;

    [ObservableProperty]
    private bool _usbRuntimeAvailable = true;

    [ObservableProperty]
    private string _usbRuntimeHintText = string.Empty;

    [ObservableProperty]
    private bool _hostSharedFoldersEnabled = true;

    [ObservableProperty]
    private string _usbDiagnosticsHyperVSocketText = "Unbekannt";

    [ObservableProperty]
    private string _usbDiagnosticsRegistryServiceText = "Unbekannt";

    [ObservableProperty]
    private string _usbDiagnosticsFallbackText = "Nein";

    [ObservableProperty]
    private bool _selectedUsbDeviceAutoShareEnabled;

    public ObservableCollection<VmDefinition> AvailableVms { get; } = [];

    public ObservableCollection<HyperVSwitchInfo> AvailableSwitches { get; } = [];

    public ObservableCollection<HyperVVmNetworkAdapterInfo> AvailableVmNetworkAdapters { get; } = [];

    public ObservableCollection<VmTrayAdapterOption> AvailableVmTrayAdapterOptions { get; } = [];

    public ObservableCollection<HyperVVmNetworkAdapterInfo> AvailableVmAdaptersForRename { get; } = [];

    public ObservableCollection<HyperVCheckpointInfo> AvailableCheckpoints { get; } = [];

    public ObservableCollection<HyperVCheckpointTreeItem> AvailableCheckpointTree { get; } = [];

    public ObservableCollection<UiNotification> Notifications { get; } = [];

    public ObservableCollection<UsbIpDeviceInfo> UsbDevices { get; } = [];

    public ObservableCollection<HostSharedFolderDefinition> HostSharedFolders { get; } = [];

    public IReadOnlyList<VmImportModeOption> VmImportModeOptions { get; } =
    [
        new() { Key = "copy", Label = "Kopieren (Quellordner + Zielordner)", Description = "Fragt Quellordner und Zielordner ab, kopiert die VM und erzeugt eine neue VM-ID." },
        new() { Key = "register", Label = "Direkt registrieren (nur Quellordner)", Description = "Fragt nur den Quellordner ab und registriert die VM am bestehenden Speicherort." },
        new() { Key = "restore", Label = "Wiederherstellen (nur Quellordner)", Description = "Fragt nur den Quellordner ab und importiert ohne Kopie mit bestehender VM-ID." }
    ];

    public string DefaultSwitchName { get; }

    public string ConfigPath => _configPath;

    public IReadOnlyList<string> AvailableUiThemes { get; } = ["Dark", "Light"];

    public bool HasConfigurationNotice => !string.IsNullOrWhiteSpace(ConfigurationNotice);

    public UiNotification? LatestNotification => Notifications.FirstOrDefault();

    public string LastNotificationText
    {
        get
        {
            var latest = LatestNotification;
            if (latest is null)
            {
                return "Keine Notifications";
            }

            var text = $"[{latest.Timestamp:HH:mm:ss}] [{latest.Level}] {latest.Message}";
            return text.Length <= 120 ? text : $"{text[..117]}...";
        }
    }

    public string LogToggleText => IsLogExpanded ? "▾ Log einklappen" : "▸ Log ausklappen";

    public string SelectedVmDisplayName => SelectedVm?.DisplayLabel ?? "-";

    public bool HasBusyProgress => IsBusy && BusyProgressPercent >= 0;

    public string SelectedVmAdapterSwitchDisplay
    {
        get
        {
            if (SelectedVm is null)
            {
                return "-";
            }

            if (SelectedVmNetworkAdapter is null)
            {
                return SelectedVmCurrentSwitch;
            }

            var adapterName = GetAdapterDisplayName(SelectedVmNetworkAdapter);
            var switchName = NormalizeSwitchDisplayName(SelectedVmNetworkAdapter.SwitchName);
            return $"{adapterName} | {switchName}";
        }
    }

    public IAsyncRelayCommand StartDefaultVmCommand { get; }

    public IAsyncRelayCommand StopDefaultVmCommand { get; }

    public IAsyncRelayCommand CreateCheckpointCommand { get; }

    public IAsyncRelayCommand StartSelectedVmCommand { get; }

    public IAsyncRelayCommand StopSelectedVmCommand { get; }

    public IAsyncRelayCommand TurnOffSelectedVmCommand { get; }

    public IAsyncRelayCommand RestartSelectedVmCommand { get; }

    public IAsyncRelayCommand OpenConsoleCommand { get; }

    public IAsyncRelayCommand ReopenConsoleWithSessionEditCommand { get; }

    public IAsyncRelayCommand ExportSelectedVmCommand { get; }

    public IAsyncRelayCommand ImportVmCommand { get; }

    public IAsyncRelayCommand LoadSwitchesCommand { get; }

    public IAsyncRelayCommand RefreshSwitchesCommand { get; }

    public IAsyncRelayCommand ConnectSelectedSwitchCommand { get; }

    public IAsyncRelayCommand DisconnectSwitchCommand { get; }

    public IAsyncRelayCommand<string> ConnectAdapterToSwitchByKeyCommand { get; }

    public IAsyncRelayCommand<string> DisconnectAdapterByNameCommand { get; }

    public IAsyncRelayCommand RefreshVmStatusCommand { get; }

    public IAsyncRelayCommand LoadCheckpointsCommand { get; }

    public IAsyncRelayCommand ApplyCheckpointCommand { get; }

    public IAsyncRelayCommand DeleteCheckpointCommand { get; }

    public IRelayCommand AddVmCommand { get; }

    public IRelayCommand RemoveVmCommand { get; }

    public IRelayCommand SetDefaultVmCommand { get; }

    public IAsyncRelayCommand RenameVmAdapterCommand { get; }

    public IAsyncRelayCommand SaveConfigCommand { get; }

    public IAsyncRelayCommand ReloadConfigCommand { get; }

    public IAsyncRelayCommand RestartHnsCommand { get; }

    public IAsyncRelayCommand CheckForUpdatesCommand { get; }

    public IAsyncRelayCommand InstallUpdateCommand { get; }

    public IRelayCommand OpenReleasePageCommand { get; }

    public IRelayCommand ToggleLogCommand { get; }

    public IRelayCommand OpenLogFileCommand { get; }

    public IRelayCommand<VmDefinition> SelectVmFromChipCommand { get; }

    public IRelayCommand ClearNotificationsCommand { get; }

    public IRelayCommand CopyNotificationsCommand { get; }

    public IAsyncRelayCommand<string> StartVmByNameCommand { get; }

    public IAsyncRelayCommand<string> StopVmByNameCommand { get; }

    public IAsyncRelayCommand<string> TurnOffVmByNameCommand { get; }

    public IAsyncRelayCommand<string> RestartVmByNameCommand { get; }

    public IAsyncRelayCommand<string> OpenConsoleByNameCommand { get; }

    public IAsyncRelayCommand<string> CreateSnapshotByNameCommand { get; }

    public IAsyncRelayCommand RefreshUsbDevicesCommand { get; }

    public IAsyncRelayCommand BindUsbDeviceCommand { get; }

    public IAsyncRelayCommand UnbindUsbDeviceCommand { get; }

    public IAsyncRelayCommand DetachUsbDeviceCommand { get; }

    private readonly IHyperVService _hyperVService;
    private readonly IHnsService _hnsService;
    private readonly IConfigService _configService;
    private readonly IStartupService _startupService;
    private readonly IUpdateService _updateService;
    private readonly IUsbIpService _usbIpService;
    private readonly IUiInteropService _uiInteropService;
    private readonly string _configPath;

    private readonly CancellationTokenSource _lifetimeCancellation = new();
    private List<string> _trayVmNames = [];
    private readonly Dictionary<string, VmDefinition> _configuredVmDefinitions = new(StringComparer.OrdinalIgnoreCase);
    private static readonly char[] VmAdapterInvalidNameChars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|'];
    private int _selectedVmChangeSuppressionDepth;
    private int _configChangeSuppressionDepth;
    private int _lastSelectedMenuIndex;
    private bool _isHandlingMenuSelectionChange;
    private bool _suppressUsbAutoShareToggleHandling;
    private readonly SemaphoreSlim _usbTrayRefreshGate = new(1, 1);
    private readonly HashSet<string> _usbAutoShareDeviceKeys = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, DateTimeOffset> _usbAttachedWithoutAckSinceUtc = new(StringComparer.OrdinalIgnoreCase);
    private static readonly HttpClient UpdateDownloadClient = new();
    private int _uiNumLockWatcherIntervalSeconds = 30;

    public event EventHandler? TrayStateChanged;

    public bool CanPromptSaveOnClose => HasPendingConfigChanges && !IsBusy;

    public bool HasUsbAutoShareConfigured => _usbAutoShareDeviceKeys.Count > 0;

    public MainViewModel(
        ConfigLoadResult configResult,
        IHyperVService hyperVService,
        IHnsService hnsService,
        IConfigService configService,
        IStartupService startupService,
        IUpdateService updateService,
        IUsbIpService usbIpService,
        IUiInteropService uiInteropService)
    {
        _hyperVService = hyperVService;
        _hnsService = hnsService;
        _configService = configService;
        _startupService = startupService;
        _updateService = updateService;
        _usbIpService = usbIpService;
        _uiInteropService = uiInteropService;
        SelectedVmImportModeOption = VmImportModeOptions.FirstOrDefault();
        _configPath = configResult.ConfigPath;

        _configChangeSuppressionDepth++;

        WindowTitle = "HyperTool";
        DefaultVmName = configResult.Config.DefaultVmName;
        LastSelectedVmName = configResult.Config.LastSelectedVmName;
        DefaultSwitchName = configResult.Config.DefaultSwitchName;
        VmConnectComputerName = NormalizeVmConnectComputerName(configResult.Config.VmConnectComputerName);
        ConfigurationNotice = configResult.Notice;
        HnsEnabled = configResult.Config.Hns.Enabled;
        HnsAutoRestartAfterDefaultSwitch = configResult.Config.Hns.AutoRestartAfterDefaultSwitch;
        HnsAutoRestartAfterAnyConnect = configResult.Config.Hns.AutoRestartAfterAnyConnect;
        UiEnableTrayIcon = true;
        UiEnableTrayMenu = configResult.Config.Ui.EnableTrayMenu;
        UiStartMinimized = configResult.Config.Ui.StartMinimized;
        UiStartWithWindows = configResult.Config.Ui.StartWithWindows;
        UiOpenConsoleAfterVmStart = configResult.Config.Ui.OpenConsoleAfterVmStart;
        UiRestoreNumLockAfterVmStart = configResult.Config.Ui.RestoreNumLockAfterVmStart;
        _uiNumLockWatcherIntervalSeconds = Math.Clamp(configResult.Config.Ui.NumLockWatcherIntervalSeconds, 5, 600);
        UiOpenVmConnectWithSessionEdit = configResult.Config.Ui.OpenVmConnectWithSessionEdit;
        UiTheme = NormalizeUiTheme(configResult.Config.Ui.Theme);
        UpdateCheckOnStartup = configResult.Config.Update.CheckOnStartup;
        GithubOwner = configResult.Config.Update.GitHubOwner;
        GithubRepo = configResult.Config.Update.GitHubRepo;
        HostUsbSharingEnabled = configResult.Config.Usb.Enabled;
        UsbAutoDetachOnClientDisconnect = configResult.Config.Usb.AutoDetachOnClientDisconnect;
        UsbUnshareOnExit = configResult.Config.Usb.UnshareOnExit;
        HostSharedFoldersEnabled = configResult.Config.SharedFolders.Enabled;
        DefaultVmImportDestinationPath = configResult.Config.DefaultVmImportDestinationPath?.Trim() ?? string.Empty;
        _usbAutoShareDeviceKeys.Clear();
        foreach (var key in configResult.Config.Usb.AutoShareDeviceKeys)
        {
            if (!string.IsNullOrWhiteSpace(key))
            {
                _usbAutoShareDeviceKeys.Add(key.Trim());
            }
        }
        ApplyConfiguredSharedFolders(configResult.Config.SharedFolders.HostDefinitions);
        ApplyConfiguredVmDefinitions(configResult.Config.Vms);
        _trayVmNames = NormalizeTrayVmNames(configResult.Config.Ui.TrayVmNames);
        AppVersion = ResolveAppVersion();
        UpdateStatus = "Noch nicht geprüft";

        if (configResult.IsGenerated)
        {
            StatusText = "Beispiel-Konfiguration erzeugt";
        }
        else if (configResult.HasValidationFixes)
        {
            StatusText = "Konfiguration korrigiert geladen";
        }
        else
        {
            StatusText = "Konfiguration geladen";
        }

        StartSelectedVmCommand = new AsyncRelayCommand(StartSelectedVmAsync, CanExecuteStartVmAction);
        StopSelectedVmCommand = new AsyncRelayCommand(StopSelectedVmAsync, CanExecuteStopVmAction);
        TurnOffSelectedVmCommand = new AsyncRelayCommand(TurnOffSelectedVmAsync, CanExecuteStopVmAction);
        RestartSelectedVmCommand = new AsyncRelayCommand(RestartSelectedVmAsync, CanExecuteRestartVmAction);
        OpenConsoleCommand = new AsyncRelayCommand(OpenConsoleAsync, CanExecuteStopVmAction);
        ReopenConsoleWithSessionEditCommand = new AsyncRelayCommand(ReopenConsoleWithSessionEditAsync, CanExecuteVmAction);
        ExportSelectedVmCommand = new AsyncRelayCommand(ExportSelectedVmAsync, () => !IsBusy && SelectedVmForConfig is not null);
        ImportVmCommand = new AsyncRelayCommand(ImportVmAsync, () => !IsBusy);

        LoadSwitchesCommand = new AsyncRelayCommand(RefreshSwitchesAsync, () => !IsBusy);
        RefreshSwitchesCommand = new AsyncRelayCommand(RefreshSwitchesAsync, () => !IsBusy);
        ConnectSelectedSwitchCommand = new AsyncRelayCommand(ConnectSelectedSwitchAsync, () => !IsBusy && SelectedVm is not null && SelectedVmNetworkAdapter is not null && SelectedSwitch is not null && AreSwitchesLoaded);
        DisconnectSwitchCommand = new AsyncRelayCommand(DisconnectSwitchAsync, () => !IsBusy && SelectedVm is not null && SelectedVmNetworkAdapter is not null);
        ConnectAdapterToSwitchByKeyCommand = new AsyncRelayCommand<string>(ConnectAdapterToSwitchByKeyAsync, _ => !IsBusy && SelectedVm is not null && AreSwitchesLoaded);
        DisconnectAdapterByNameCommand = new AsyncRelayCommand<string>(DisconnectAdapterByNameAsync, _ => !IsBusy && SelectedVm is not null);
        RefreshVmStatusCommand = new AsyncRelayCommand(RefreshRuntimeDataAsync, () => !IsBusy);

        LoadCheckpointsCommand = new AsyncRelayCommand(LoadCheckpointsAsync, () => SelectedVm is not null);
        ApplyCheckpointCommand = new AsyncRelayCommand(ApplyCheckpointAsync, () => !IsBusy && SelectedVm is not null && SelectedCheckpoint is not null);
        DeleteCheckpointCommand = new AsyncRelayCommand(DeleteCheckpointAsync, () => !IsBusy && SelectedVm is not null && SelectedCheckpoint is not null);

        AddVmCommand = new RelayCommand(AddVm);
        RemoveVmCommand = new RelayCommand(RemoveSelectedVm);
        SetDefaultVmCommand = new RelayCommand(SetDefaultVmFromSelection);
        RenameVmAdapterCommand = new AsyncRelayCommand(RenameVmAdapterAsync, CanExecuteRenameVmAdapter);
        SaveConfigCommand = new AsyncRelayCommand(SaveConfigAsync, () => !IsBusy);
        ReloadConfigCommand = new AsyncRelayCommand(ReloadConfigAsync, () => !IsBusy);
        RestartHnsCommand = new AsyncRelayCommand(RestartHnsAsync, () => !IsBusy);
        CheckForUpdatesCommand = new AsyncRelayCommand(CheckForUpdatesAsync, () => !IsBusy);
        InstallUpdateCommand = new AsyncRelayCommand(InstallUpdateAsync, () => !IsBusy && UpdateInstallAvailable && !string.IsNullOrWhiteSpace(InstallerDownloadUrl));
        OpenReleasePageCommand = new RelayCommand(OpenReleasePage);
        ToggleLogCommand = new RelayCommand(ToggleLog);
        OpenLogFileCommand = new RelayCommand(OpenLogFile);
        SelectVmFromChipCommand = new RelayCommand<VmDefinition>(SelectVmFromChip);
        ClearNotificationsCommand = new RelayCommand(ClearNotifications);
        CopyNotificationsCommand = new RelayCommand(CopyNotificationsToClipboard);

        StartVmByNameCommand = new AsyncRelayCommand<string>(StartVmByNameAsync, _ => !IsBusy);
        StopVmByNameCommand = new AsyncRelayCommand<string>(StopVmByNameAsync, _ => !IsBusy);
        TurnOffVmByNameCommand = new AsyncRelayCommand<string>(TurnOffVmByNameAsync, _ => !IsBusy);
        RestartVmByNameCommand = new AsyncRelayCommand<string>(RestartVmByNameAsync, _ => !IsBusy);
        OpenConsoleByNameCommand = new AsyncRelayCommand<string>(OpenConsoleByNameAsync, _ => !IsBusy);
        CreateSnapshotByNameCommand = new AsyncRelayCommand<string>(CreateSnapshotByNameAsync, _ => !IsBusy);

        RefreshUsbDevicesCommand = new AsyncRelayCommand(RefreshUsbDevicesAsync, () => !IsBusy && UsbRuntimeAvailable && HostUsbSharingEnabled);
        BindUsbDeviceCommand = new AsyncRelayCommand(BindSelectedUsbDeviceAsync, CanExecuteBindUsbAction);
        UnbindUsbDeviceCommand = new AsyncRelayCommand(UnbindSelectedUsbDeviceAsync, CanExecuteUnbindUsbAction);
        DetachUsbDeviceCommand = new AsyncRelayCommand(DetachSelectedUsbDeviceAsync, CanExecuteDetachUsbAction);

        StartDefaultVmCommand = new AsyncRelayCommand(StartDefaultVmAsync, () => !IsBusy);
        StopDefaultVmCommand = new AsyncRelayCommand(StopDefaultVmAsync, () => !IsBusy);
        CreateCheckpointCommand = new AsyncRelayCommand(CreateCheckpointAsync, CanExecuteVmAction);

        SelectedVmForConfig = SelectedVm;
        SelectedDefaultVmForConfig = SelectedVm;

        IsLogExpanded = false;
        Notifications.CollectionChanged += OnNotificationsChanged;

        _configChangeSuppressionDepth--;
        HasPendingConfigChanges = false;
        _lastSelectedMenuIndex = SelectedMenuIndex;

        _ = InitializeAsync();
        _ = RunNumLockWatcherAsync();
    }

    partial void OnIsLogExpandedChanged(bool value)
    {
        OnPropertyChanged(nameof(LogToggleText));
    }

    partial void OnAreSwitchesLoadedChanged(bool value)
    {
        ConnectSelectedSwitchCommand.NotifyCanExecuteChanged();
    }

    partial void OnConfigurationNoticeChanged(string? value)
    {
        OnPropertyChanged(nameof(HasConfigurationNotice));
    }

    partial void OnIsBusyChanged(bool value)
    {
        StartSelectedVmCommand.NotifyCanExecuteChanged();
        StopSelectedVmCommand.NotifyCanExecuteChanged();
        TurnOffSelectedVmCommand.NotifyCanExecuteChanged();
        RestartSelectedVmCommand.NotifyCanExecuteChanged();
        OpenConsoleCommand.NotifyCanExecuteChanged();
        ReopenConsoleWithSessionEditCommand.NotifyCanExecuteChanged();
        ExportSelectedVmCommand.NotifyCanExecuteChanged();
        ImportVmCommand.NotifyCanExecuteChanged();
        LoadSwitchesCommand.NotifyCanExecuteChanged();
        RefreshSwitchesCommand.NotifyCanExecuteChanged();
        ConnectSelectedSwitchCommand.NotifyCanExecuteChanged();
        DisconnectSwitchCommand.NotifyCanExecuteChanged();
        ConnectAdapterToSwitchByKeyCommand.NotifyCanExecuteChanged();
        DisconnectAdapterByNameCommand.NotifyCanExecuteChanged();
        RefreshVmStatusCommand.NotifyCanExecuteChanged();
        CreateCheckpointCommand.NotifyCanExecuteChanged();
        LoadCheckpointsCommand.NotifyCanExecuteChanged();
        ApplyCheckpointCommand.NotifyCanExecuteChanged();
        DeleteCheckpointCommand.NotifyCanExecuteChanged();
        LoadCheckpointsCommand.NotifyCanExecuteChanged();
        StartDefaultVmCommand.NotifyCanExecuteChanged();
        StopDefaultVmCommand.NotifyCanExecuteChanged();
        SaveConfigCommand.NotifyCanExecuteChanged();
        ReloadConfigCommand.NotifyCanExecuteChanged();
        RestartHnsCommand.NotifyCanExecuteChanged();
        CheckForUpdatesCommand.NotifyCanExecuteChanged();
        InstallUpdateCommand.NotifyCanExecuteChanged();
        StartVmByNameCommand.NotifyCanExecuteChanged();
        StopVmByNameCommand.NotifyCanExecuteChanged();
        TurnOffVmByNameCommand.NotifyCanExecuteChanged();
        RestartVmByNameCommand.NotifyCanExecuteChanged();
        OpenConsoleByNameCommand.NotifyCanExecuteChanged();
        CreateSnapshotByNameCommand.NotifyCanExecuteChanged();
        RenameVmAdapterCommand.NotifyCanExecuteChanged();
        RefreshUsbDevicesCommand.NotifyCanExecuteChanged();
        BindUsbDeviceCommand.NotifyCanExecuteChanged();
        UnbindUsbDeviceCommand.NotifyCanExecuteChanged();
        DetachUsbDeviceCommand.NotifyCanExecuteChanged();
        OnPropertyChanged(nameof(HasBusyProgress));
    }

    partial void OnBusyProgressPercentChanged(int value)
    {
        OnPropertyChanged(nameof(HasBusyProgress));
    }

    partial void OnSelectedMenuIndexChanged(int value)
    {
        if (_isHandlingMenuSelectionChange)
        {
            _lastSelectedMenuIndex = value;

            if (value == 0)
            {
                _ = HandleNetworkTabActivatedAsync();
            }

            return;
        }

        if (ShouldPromptSaveWhenLeavingConfig(value) && !TryPromptSaveConfigChanges())
        {
            _isHandlingMenuSelectionChange = true;
            try
            {
                SelectedMenuIndex = ConfigMenuIndex;
            }
            finally
            {
                _isHandlingMenuSelectionChange = false;
            }

            _lastSelectedMenuIndex = ConfigMenuIndex;
            return;
        }

        _lastSelectedMenuIndex = value;

        if (value == 0)
        {
            _ = HandleNetworkTabActivatedAsync();
        }
        else if (value == 1)
        {
            _ = LoadUsbDevicesAsync(showNotification: false);
        }
    }

    partial void OnHnsAutoRestartAfterDefaultSwitchChanged(bool value) => MarkConfigDirty();

    partial void OnHnsAutoRestartAfterAnyConnectChanged(bool value) => MarkConfigDirty();

    partial void OnUiEnableTrayMenuChanged(bool value)
    {
        NotifyTrayStateChanged();
        MarkConfigDirty();
    }

    partial void OnUiStartMinimizedChanged(bool value) => MarkConfigDirty();

    partial void OnUiStartWithWindowsChanged(bool value) => MarkConfigDirty();

    partial void OnUiOpenConsoleAfterVmStartChanged(bool value) => MarkConfigDirty();

    partial void OnUiRestoreNumLockAfterVmStartChanged(bool value) => MarkConfigDirty();

    partial void OnUiOpenVmConnectWithSessionEditChanged(bool value) => MarkConfigDirty();

    partial void OnDefaultVmImportDestinationPathChanged(string value)
    {
        var normalized = value?.Trim() ?? string.Empty;
        if (!string.Equals(value, normalized, StringComparison.Ordinal))
        {
            DefaultVmImportDestinationPath = normalized;
            return;
        }

        MarkConfigDirty();
    }

    partial void OnUiThemeChanged(string value)
    {
        MarkConfigDirty();
    }

    partial void OnVmConnectComputerNameChanged(string value) => MarkConfigDirty();

    partial void OnUpdateCheckOnStartupChanged(bool value) => MarkConfigDirty();

    partial void OnGithubOwnerChanged(string value) => MarkConfigDirty();

    partial void OnGithubRepoChanged(string value) => MarkConfigDirty();

    partial void OnHostUsbSharingEnabledChanged(bool value)
    {
        RefreshUsbDevicesCommand?.NotifyCanExecuteChanged();
        BindUsbDeviceCommand?.NotifyCanExecuteChanged();
        UnbindUsbDeviceCommand?.NotifyCanExecuteChanged();
        DetachUsbDeviceCommand?.NotifyCanExecuteChanged();

        if (!value)
        {
            UsbDevices.Clear();
            SelectedUsbDevice = null;
            UsbStatusText = "USB Share ist global im Host deaktiviert.";
            UsbRuntimeHintText = string.Empty;
        }

        MarkConfigDirty();
        NotifyTrayStateChanged();
    }

    partial void OnUsbAutoDetachOnClientDisconnectChanged(bool value) => MarkConfigDirty();

    partial void OnUsbUnshareOnExitChanged(bool value) => MarkConfigDirty();

    partial void OnHostSharedFoldersEnabledChanged(bool value)
    {
        MarkConfigDirty();
    }

    public async Task SetHostUsbSharingEnabledAsync(bool enabled)
    {
        if (HostUsbSharingEnabled == enabled)
        {
            return;
        }

        if (!enabled)
        {
            if (!UsbRuntimeAvailable)
            {
                HostUsbSharingEnabled = false;
                return;
            }

            try
            {
                await UnshareAllSharedUsbAsync(
                    timeout: TimeSpan.FromSeconds(20),
                    successMessage: "Beim Deaktivieren wurden {0} USB-Freigabe(n) entfernt.",
                    failedMessage: "Beim Deaktivieren konnten {0} USB-Freigabe(n) nicht entfernt werden.",
                    logContext: "host-usb-disable");
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "USB-Freigaben konnten beim Deaktivieren nicht vollständig entfernt werden.");
                AddNotification("USB-Freigaben konnten nicht vollständig entfernt werden (usbipd ggf. nicht verfügbar).", "Warning");
            }
            finally
            {
                try
                {
                    await _usbIpService.ShutdownElevatedSessionAsync(CancellationToken.None);
                }
                catch (Exception ex)
                {
                    Log.Debug(ex, "Elevated USB session could not be closed after disabling host USB sharing.");
                }
            }
        }

        HostUsbSharingEnabled = enabled;
    }

    public Task SetHostSharedFoldersEnabledAsync(bool enabled)
    {
        if (HostSharedFoldersEnabled != enabled)
        {
            HostSharedFoldersEnabled = enabled;
        }

        return Task.CompletedTask;
    }

    partial void OnSelectedVmChanged(VmDefinition? value)
    {
        ConnectSelectedSwitchCommand.NotifyCanExecuteChanged();
        CreateCheckpointCommand.NotifyCanExecuteChanged();
        ReopenConsoleWithSessionEditCommand.NotifyCanExecuteChanged();
        SelectedVmForConfig = value;
        OnPropertyChanged(nameof(SelectedVmDisplayName));
        OnPropertyChanged(nameof(SelectedVmAdapterSwitchDisplay));
        NotifyTrayStateChanged();

        if (value is null)
        {
            SelectedVmState = "Unbekannt";
            SelectedVmCurrentSwitch = NotConnectedSwitchDisplay;
            AvailableVmNetworkAdapters.Clear();
            SelectedVmNetworkAdapter = null;
            SelectedSwitch = null;
            NetworkSwitchStatusHint = "Keine VM ausgewählt.";
            AvailableCheckpoints.Clear();
            AvailableCheckpointTree.Clear();
            SelectedCheckpoint = null;
            SelectedCheckpointNode = null;
            return;
        }

        if (_selectedVmChangeSuppressionDepth > 0)
        {
            return;
        }

        LastSelectedVmName = value.Name;
        SelectedVmState = string.IsNullOrWhiteSpace(value.RuntimeState) ? "Unbekannt" : value.RuntimeState;
        SelectedVmCurrentSwitch = NormalizeSwitchDisplayName(value.RuntimeSwitchName);
        SyncSelectedSwitchWithCurrentVm(showNotificationOnMissingSwitch: false);
        _ = PersistSelectedVmAsync(value.Name);

        _ = EnsureSelectedVmNetworkSelectionAsync(showNotificationOnMissingSwitch: false);
        _ = RefreshSelectedVmStatusAfterSelectionAsync(value.Name);
        _ = LoadCheckpointsAsync();
    }

    private async Task RefreshSelectedVmStatusAfterSelectionAsync(string selectedVmName)
    {
        for (var attempt = 0; attempt < 6; attempt++)
        {
            if (SelectedVm is null
                || !string.Equals(SelectedVm.Name, selectedVmName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            if (!IsBusy)
            {
                await RefreshVmStatusAsync();
                return;
            }

            await Task.Delay(200);
        }
    }

    private async Task HandleNetworkTabActivatedAsync()
    {
        await EnsureSelectedVmNetworkSelectionAsync(showNotificationOnMissingSwitch: false);

        if (!AreSwitchesLoaded)
        {
            await RefreshSwitchesAsync();
            return;
        }

        SyncSelectedSwitchWithCurrentVm(showNotificationOnMissingSwitch: false);
    }

    partial void OnUpdateInstallAvailableChanged(bool value)
    {
        InstallUpdateCommand.NotifyCanExecuteChanged();
    }

    partial void OnInstallerDownloadUrlChanged(string value)
    {
        InstallUpdateCommand.NotifyCanExecuteChanged();
    }

    partial void OnSelectedVmStateChanged(string value)
    {
        StartSelectedVmCommand.NotifyCanExecuteChanged();
        StopSelectedVmCommand.NotifyCanExecuteChanged();
        TurnOffSelectedVmCommand.NotifyCanExecuteChanged();
        RestartSelectedVmCommand.NotifyCanExecuteChanged();
        OpenConsoleCommand.NotifyCanExecuteChanged();
        ReopenConsoleWithSessionEditCommand.NotifyCanExecuteChanged();
    }

    partial void OnSelectedVmCurrentSwitchChanged(string value)
    {
        OnPropertyChanged(nameof(SelectedVmAdapterSwitchDisplay));
    }

    partial void OnSelectedSwitchChanged(HyperVSwitchInfo? value)
    {
        ConnectSelectedSwitchCommand.NotifyCanExecuteChanged();
    }

    partial void OnSelectedVmNetworkAdapterChanged(HyperVVmNetworkAdapterInfo? value)
    {
        ConnectSelectedSwitchCommand.NotifyCanExecuteChanged();
        DisconnectSwitchCommand.NotifyCanExecuteChanged();
        OnPropertyChanged(nameof(SelectedVmAdapterSwitchDisplay));

        if (SelectedVm is null || value is null)
        {
            if (SelectedVm is not null && AvailableVmNetworkAdapters.Count == 0)
            {
                SelectedVmCurrentSwitch = NotConnectedSwitchDisplay;
                NetworkSwitchStatusHint = "Keine VM-Netzwerkkarten gefunden.";
                SelectedSwitch = null;
            }

            return;
        }

        SelectedVmCurrentSwitch = NormalizeSwitchDisplayName(value.SwitchName);

        var selectedVmEntry = AvailableVms.FirstOrDefault(vm =>
            string.Equals(vm.Name, SelectedVm.Name, StringComparison.OrdinalIgnoreCase));

        if (selectedVmEntry is not null)
        {
            selectedVmEntry.RuntimeSwitchName = SelectedVmCurrentSwitch;
        }

        SyncSelectedSwitchWithCurrentVm(showNotificationOnMissingSwitch: false);
    }

    partial void OnSelectedCheckpointChanged(HyperVCheckpointInfo? value)
    {
        ApplyCheckpointCommand.NotifyCanExecuteChanged();
        DeleteCheckpointCommand.NotifyCanExecuteChanged();
    }

    partial void OnSelectedCheckpointNodeChanged(HyperVCheckpointTreeItem? value)
    {
        SelectedCheckpoint = value?.Checkpoint;
    }

    partial void OnSelectedUsbDeviceChanged(UsbIpDeviceInfo? value)
    {
        BindUsbDeviceCommand.NotifyCanExecuteChanged();
        UnbindUsbDeviceCommand.NotifyCanExecuteChanged();
        DetachUsbDeviceCommand.NotifyCanExecuteChanged();
        _suppressUsbAutoShareToggleHandling = true;
        try
        {
            SelectedUsbDeviceAutoShareEnabled = value is not null && _usbAutoShareDeviceKeys.Contains(BuildUsbAutoShareKey(value));
        }
        finally
        {
            _suppressUsbAutoShareToggleHandling = false;
        }

        NotifyTrayStateChanged();
    }

    partial void OnSelectedUsbDeviceAutoShareEnabledChanged(bool value)
    {
        if (_suppressUsbAutoShareToggleHandling || SelectedUsbDevice is null)
        {
            return;
        }

        var key = BuildUsbAutoShareKey(SelectedUsbDevice);
        if (string.IsNullOrWhiteSpace(key))
        {
            AddNotification("Auto-Share konnte für dieses USB-Gerät nicht gespeichert werden (kein stabiler Schlüssel).", "Warning");
            return;
        }

        var changed = value
            ? _usbAutoShareDeviceKeys.Add(key)
            : _usbAutoShareDeviceKeys.Remove(key);

        if (!changed)
        {
            return;
        }

        PersistUsbAutoShareConfig();
        AddNotification(value
                ? $"Auto-Share für USB-Gerät '{SelectedUsbDevice.Description}' aktiviert."
                : $"Auto-Share für USB-Gerät '{SelectedUsbDevice.Description}' deaktiviert.",
            "Info");

        if (value
            && !IsBusy
            && !SelectedUsbDevice.IsShared
            && !string.IsNullOrWhiteSpace(SelectedUsbDevice.BusId))
        {
            _ = BindSelectedUsbDeviceAsync();
        }
    }

    partial void OnSelectedVmForConfigChanged(VmDefinition? value)
    {
        RemoveVmCommand.NotifyCanExecuteChanged();
        ExportSelectedVmCommand.NotifyCanExecuteChanged();
        RenameVmAdapterCommand.NotifyCanExecuteChanged();
        UpdateVmAdapterRenameValidationState();

        _configChangeSuppressionDepth++;
        try
        {
            SelectedVmOpenConsoleWithSessionEdit = value?.OpenConsoleWithSessionEdit ?? false;
        }
        finally
        {
            _configChangeSuppressionDepth--;
        }

        _ = LoadVmAdaptersForConfigAsync(value);
    }

    partial void OnSelectedVmOpenConsoleWithSessionEditChanged(bool value)
    {
        if (SelectedVmForConfig is null)
        {
            return;
        }

        SelectedVmForConfig.OpenConsoleWithSessionEdit = value;
        MarkConfigDirty();
    }

    partial void OnSelectedVmTrayAdapterOptionChanged(VmTrayAdapterOption? value)
    {
        if (SelectedVmForConfig is null || value is null)
        {
            return;
        }

        SelectedVmForConfig.TrayAdapterName = value.AdapterName?.Trim() ?? string.Empty;
        NotifyTrayStateChanged();
        MarkConfigDirty();
    }

    partial void OnSelectedVmAdapterForRenameChanged(HyperVVmNetworkAdapterInfo? value)
    {
        UpdateVmAdapterRenameValidationState();
        RenameVmAdapterCommand.NotifyCanExecuteChanged();
    }

    partial void OnNewVmAdapterNameChanged(string value)
    {
        UpdateVmAdapterRenameValidationState();
        RenameVmAdapterCommand.NotifyCanExecuteChanged();
    }

    private bool CanExecuteVmAction() => !IsBusy && SelectedVm is not null;

    private bool CanExecuteBindUsbAction()
    {
         return !IsBusy
                         && HostUsbSharingEnabled
             && UsbRuntimeAvailable
               && SelectedUsbDevice is not null
               && !string.IsNullOrWhiteSpace(SelectedUsbDevice.BusId)
               && !SelectedUsbDevice.IsShared;
    }

    private bool CanExecuteUnbindUsbAction()
    {
         return !IsBusy
                         && HostUsbSharingEnabled
             && UsbRuntimeAvailable
               && SelectedUsbDevice is not null
               && !string.IsNullOrWhiteSpace(SelectedUsbDevice.BusId)
               && SelectedUsbDevice.IsShared;
    }

    private bool CanExecuteDetachUsbAction()
    {
         return !IsBusy
                         && HostUsbSharingEnabled
             && UsbRuntimeAvailable
               && SelectedUsbDevice is not null
               && !string.IsNullOrWhiteSpace(SelectedUsbDevice.BusId)
               && SelectedUsbDevice.IsAttached;
    }

    private bool CanExecuteStartVmAction() => CanExecuteVmAction() && !IsRunningState(SelectedVmState);


    partial void OnUsbRuntimeAvailableChanged(bool value)
    {
        RefreshUsbDevicesCommand.NotifyCanExecuteChanged();
        BindUsbDeviceCommand.NotifyCanExecuteChanged();
        UnbindUsbDeviceCommand.NotifyCanExecuteChanged();
        DetachUsbDeviceCommand.NotifyCanExecuteChanged();

        if (!value && HostUsbSharingEnabled)
        {
            HostUsbSharingEnabled = false;
        }
    }
    private bool CanExecuteStopVmAction() => CanExecuteVmAction() && IsRunningState(SelectedVmState);

    private bool CanExecuteRestartVmAction() => CanExecuteVmAction() && IsRunningState(SelectedVmState);

    private static bool IsRunningState(string? state)
    {
        if (string.IsNullOrWhiteSpace(state))
        {
            return false;
        }

        return state.Contains("Running", StringComparison.OrdinalIgnoreCase)
               || state.Contains("Läuft", StringComparison.OrdinalIgnoreCase);
    }

    private async Task InitializeAsync()
    {
        await RefreshUsbRuntimeAvailabilityAsync();
        await LoadVmsFromHyperVWithRetryAsync();
        await RefreshSwitchesAsync();
        await RefreshHostNetworkProfileAsync();
        await RefreshVmStatusAsync();
        await LoadCheckpointsAsync();

        if (UpdateCheckOnStartup)
        {
            await CheckForUpdatesAsync();
        }
    }

    public async Task RefreshUsbRuntimeAvailabilityAsync()
    {
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(4));
            await _usbIpService.GetDevicesAsync(cts.Token);

            UsbRuntimeAvailable = true;
            UsbRuntimeHintText = string.Empty;
        }
        catch (Exception ex) when (IsUsbRuntimeMissing(ex.Message))
        {
            UsbRuntimeAvailable = false;
            UsbRuntimeHintText = "USB-Funktion deaktiviert: usbipd-win ist nicht installiert. Quelle: https://github.com/dorssel/usbipd-win";
            UsbStatusText = BuildUsbUnavailableStatus(ex.Message);

            if (HostUsbSharingEnabled)
            {
                HostUsbSharingEnabled = false;
            }
        }
        catch
        {
        }
    }

    private async Task LoadVmsFromHyperVWithRetryAsync()
    {
        var retryDelays = new[] { 300, 700, 1500 };
        Exception? lastException = null;

        for (var attempt = 0; attempt <= retryDelays.Length; attempt++)
        {
            try
            {
                await LoadVmsFromHyperVAsync();

                if (StatusText.Equals("Keine Berechtigung", StringComparison.OrdinalIgnoreCase))
                {
                    lastException = new UnauthorizedAccessException("Keine Berechtigung für Hyper-V.");
                    break;
                }

                if (AvailableVms.Count > 0)
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                lastException = ex;
            }

            if (attempt < retryDelays.Length)
            {
                await Task.Delay(retryDelays[attempt]);
            }
        }

        if (lastException is not null)
        {
            var message = lastException is UnauthorizedAccessException
                ? "Hyper-V Zugriff verweigert. Bitte HyperTool als Administrator starten oder Benutzerrechte für Hyper-V setzen."
                : $"Hyper-V scheint nicht verfügbar: {lastException.Message}";

            AddNotification(message, "Warning");
            StatusText = "Hyper-V nicht verfügbar";
            return;
        }

        AddNotification("Keine Hyper-V VMs gefunden. Bitte Hyper-V aktivieren/installieren.", "Warning");
        StatusText = "Keine Hyper-V VMs gefunden";
    }

    private async Task LoadVmsFromHyperVAsync()
    {
        await ExecuteBusyActionAsync("Hyper-V VMs werden geladen...", async token =>
        {
            var vms = await _hyperVService.GetVmsAsync(token);
            if (vms.Count == 0)
            {
                AvailableVms.Clear();
                SelectedVm = null;
                SelectedVmForConfig = null;
                SelectedDefaultVmForConfig = null;
                SelectedVmState = "Unbekannt";
                SelectedVmCurrentSwitch = NotConnectedSwitchDisplay;
                return;
            }

            var existingLabels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var existingTrayAdapters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var existingSessionEditPreference = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

            foreach (var configured in _configuredVmDefinitions.Values)
            {
                if (!string.IsNullOrWhiteSpace(configured.Name))
                {
                    existingLabels[configured.Name] = configured.Label;
                    existingTrayAdapters[configured.Name] = configured.TrayAdapterName;
                    existingSessionEditPreference[configured.Name] = configured.OpenConsoleWithSessionEdit;
                }
            }

            foreach (var vm in AvailableVms)
            {
                existingLabels[vm.Name] = vm.Label;
                existingTrayAdapters[vm.Name] = vm.TrayAdapterName;
                existingSessionEditPreference[vm.Name] = vm.OpenConsoleWithSessionEdit;
            }

            var orderedVms = vms.OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase).ToList();

            AvailableVms.Clear();
            foreach (var vmInfo in orderedVms)
            {
                var label = existingLabels.TryGetValue(vmInfo.Name, out var existingLabel) && !string.IsNullOrWhiteSpace(existingLabel)
                    ? existingLabel
                    : vmInfo.Name;

                AvailableVms.Add(new VmDefinition
                {
                    Name = vmInfo.Name,
                    Label = label,
                    RuntimeState = vmInfo.State,
                    RuntimeSwitchName = NormalizeSwitchDisplayName(vmInfo.CurrentSwitchName),
                    TrayAdapterName = existingTrayAdapters.TryGetValue(vmInfo.Name, out var trayAdapterName) ? trayAdapterName : string.Empty,
                    OpenConsoleWithSessionEdit = existingSessionEditPreference.TryGetValue(vmInfo.Name, out var openWithSessionEdit) && openWithSessionEdit
                });
            }

            if (AvailableVms.Count > 0 && !AvailableVms.Any(vm => string.Equals(vm.Name, DefaultVmName, StringComparison.OrdinalIgnoreCase)))
            {
                DefaultVmName = AvailableVms[0].Name;
            }

            var preferredVmName = !string.IsNullOrWhiteSpace(LastSelectedVmName)
                ? LastSelectedVmName
                : DefaultVmName;

            SetSelectedVmInternal(AvailableVms.FirstOrDefault(vm => string.Equals(vm.Name, preferredVmName, StringComparison.OrdinalIgnoreCase))
                                  ?? AvailableVms.FirstOrDefault());
            SelectedVmForConfig = SelectedVm;
            SelectedDefaultVmForConfig = AvailableVms.FirstOrDefault(vm => string.Equals(vm.Name, DefaultVmName, StringComparison.OrdinalIgnoreCase))
                                       ?? SelectedVm;
            NotifyTrayStateChanged();

            AddNotification($"{AvailableVms.Count} Hyper-V VM(s) automatisch geladen.", "Info");
        }, showNotificationOnErrorOnly: true);
    }

    private async Task StartSelectedVmAsync()
    {
        await ExecuteBusyActionAsync("VM wird gestartet...", async token =>
        {
            await _hyperVService.StartVmAsync(SelectedVm!.Name, token);
            AddNotification($"VM '{SelectedVm.Name}' gestartet.", "Success");

            if (UiOpenConsoleAfterVmStart)
            {
                await _hyperVService.OpenVmConnectAsync(SelectedVm.Name, VmConnectComputerName, ShouldOpenConsoleWithSessionEdit(SelectedVm.Name), token);
                AddNotification($"Konsole für '{SelectedVm.Name}' geöffnet.", "Info");
            }
        });
        await RefreshVmStatusAsync();
    }

    private async Task StopSelectedVmAsync()
    {
        await ExecuteBusyActionAsync("VM wird gestoppt...", async token =>
        {
            await _hyperVService.StopVmGracefulAsync(SelectedVm!.Name, token);
            AddNotification($"VM '{SelectedVm.Name}' graceful gestoppt.", "Success");
        });
        await RefreshVmStatusAsync();
    }

    private async Task TurnOffSelectedVmAsync()
    {
        await ExecuteBusyActionAsync("VM wird hart ausgeschaltet...", async token =>
        {
            await _hyperVService.TurnOffVmAsync(SelectedVm!.Name, token);
            AddNotification($"VM '{SelectedVm.Name}' hart ausgeschaltet.", "Warning");
        });
        await RefreshVmStatusAsync();
    }

    private async Task RestartSelectedVmAsync()
    {
        await ExecuteBusyActionAsync("VM wird neu gestartet...", async token =>
        {
            await _hyperVService.RestartVmAsync(SelectedVm!.Name, token);
            AddNotification($"VM '{SelectedVm.Name}' neu gestartet.", "Success");
        });
        await RefreshVmStatusAsync();
    }

    private async Task OpenConsoleAsync()
    {
        await ExecuteBusyActionAsync("vmconnect wird geöffnet...", async token =>
        {
            await _hyperVService.OpenVmConnectAsync(SelectedVm!.Name, VmConnectComputerName, ShouldOpenConsoleWithSessionEdit(SelectedVm!.Name), token);
            AddNotification($"Konsole für '{SelectedVm.Name}' geöffnet.", "Info");
        });
    }

    private async Task ReopenConsoleWithSessionEditAsync()
    {
        if (SelectedVm is null)
        {
            return;
        }

        var vmName = SelectedVm.Name;

        await ExecuteBusyActionAsync("Konsole wird mit Sitzungsbearbeitung neu aufgebaut...", async token =>
        {
            await _hyperVService.ReopenVmConnectWithSessionEditAsync(vmName, VmConnectComputerName, token);
            AddNotification($"Konsole für '{vmName}' wurde mit Sitzungsbearbeitung neu aufgebaut.", "Info");
        });
    }

    private async Task ExportSelectedVmAsync()
    {
        if (SelectedVmForConfig is null)
        {
            return;
        }

        var vmName = SelectedVmForConfig.Name;

        try
        {
            var selectedFolder = PickFolderPath($"Zielordner für Backup-Export von '{vmName}' auswählen");
            if (string.IsNullOrWhiteSpace(selectedFolder))
            {
                AddNotification("VM-Export abgebrochen.", "Info");
                return;
            }

            var exportPath = Path.Combine(selectedFolder, $"{vmName}-{DateTime.Now:yyyyMMdd-HHmmss}");

            var spaceCheck = await _hyperVService.CheckExportDiskSpaceAsync(vmName, exportPath, _lifetimeCancellation.Token);
            if (!spaceCheck.HasEnoughSpace)
            {
                AddNotification(
                    $"Zu wenig Speicherplatz auf {spaceCheck.TargetDrive}: benötigt {FormatByteSize(spaceCheck.RequiredBytes)}, verfügbar {FormatByteSize(spaceCheck.AvailableBytes)}.",
                    "Error");
                return;
            }

            Directory.CreateDirectory(exportPath);

            var hasReliableProgress = false;
            var progress = new Progress<int>(percent =>
            {
                var clampedPercent = Math.Clamp(percent, 0, 100);

                if (!hasReliableProgress)
                {
                    if (clampedPercent < 2)
                    {
                        BusyProgressPercent = -1;
                        BusyText = $"VM '{vmName}' wird exportiert... (Fortschritt wird von Hyper-V ermittelt)";
                        return;
                    }

                    hasReliableProgress = true;
                }

                BusyProgressPercent = clampedPercent;
                BusyText = $"VM '{vmName}' wird exportiert... {clampedPercent}%";
            });

            await ExecuteBusyActionAsync($"VM '{vmName}' wird exportiert...", async token =>
            {
                BusyProgressPercent = -1;
                BusyText = $"VM '{vmName}' wird exportiert... (Fortschritt wird von Hyper-V ermittelt)";
                await _hyperVService.ExportVmAsync(vmName, exportPath, progress, token);
                AddNotification($"VM '{vmName}' exportiert nach: {exportPath}", "Success");
            });
        }
        catch (Exception ex)
        {
            Log.Error(ex, "VM export failed for {VmName}.", vmName);
            AddNotification($"VM-Export fehlgeschlagen: {ex.Message}", "Error");
        }
    }

    private async Task ImportVmAsync()
    {
        var importPath = PickFolderPath("Schritt 1/2: Ordner der zu importierenden Hyper-V VM auswählen");
        if (string.IsNullOrWhiteSpace(importPath))
        {
            AddNotification("VM-Import abgebrochen.", "Info");
            return;
        }

        try
        {
            var hasVmConfig = Directory.EnumerateFiles(importPath, "*.vmcx", SearchOption.AllDirectories).Any()
                              || Directory.EnumerateFiles(importPath, "*.xml", SearchOption.AllDirectories).Any();

            if (!hasVmConfig)
            {
                AddNotification("Im gewählten Quellordner wurde keine Hyper-V Konfigurationsdatei (.vmcx/.xml) gefunden.", "Error");
                return;
            }
        }
        catch (Exception ex)
        {
            AddNotification($"Quellordner konnte nicht geprüft werden: {ex.Message}", "Error");
            return;
        }

        var importMode = (SelectedVmImportModeOption?.Key ?? "copy").Trim().ToLowerInvariant();
        var isCopyMode = string.Equals(importMode, "copy", StringComparison.OrdinalIgnoreCase);
        var destinationPath = (DefaultVmImportDestinationPath ?? string.Empty).Trim();

        if (isCopyMode && string.IsNullOrWhiteSpace(destinationPath))
        {
            destinationPath = PickFolderPath("Schritt 2/2: Zielordner für den VM-Import (Kopieren) auswählen")?.Trim() ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(destinationPath))
            {
                DefaultVmImportDestinationPath = destinationPath;
                AddNotification("Import-Zielordner für künftige VM-Imports gesetzt.", "Info");
            }
        }

        if (isCopyMode && string.IsNullOrWhiteSpace(destinationPath))
        {
            AddNotification("VM-Import abgebrochen (kein Zielordner ausgewählt).", "Info");
            return;
        }

        var progress = new Progress<int>(percent =>
        {
            var clampedPercent = Math.Clamp(percent, 0, 100);

            BusyProgressPercent = -1;
            BusyText = clampedPercent >= 100
                ? "VM-Import wird abgeschlossen..."
                : "VM wird importiert... (Fortschritt wird von Hyper-V ermittelt)";
        });

        var requestedVmName = string.IsNullOrWhiteSpace(ImportVmRequestedName) ? null : ImportVmRequestedName.Trim();
        var requestedFolderName = requestedVmName;

        AddNotification($"VM-Import gestartet · Modus: {SelectedVmImportModeOption?.Label ?? importMode}", "Info");

        await ExecuteBusyActionAsync("VM wird importiert...", async token =>
        {
            BusyProgressPercent = -1;
            BusyText = "VM wird importiert... (Fortschritt wird von Hyper-V ermittelt)";

            var importResult = await _hyperVService.ImportVmAsync(
                importPath,
                destinationPath,
                requestedVmName,
                requestedFolderName,
                importMode,
                progress,
                token);

            var modeLabel = SelectedVmImportModeOption?.Label ?? importResult.ImportMode;
            var destinationInfo = string.IsNullOrWhiteSpace(importResult.DestinationFolderPath)
                ? (isCopyMode ? destinationPath : "aktueller Speicherort (keine Kopie)")
                : importResult.DestinationFolderPath;

            AddNotification($"VM '{importResult.VmName}' importiert ({modeLabel}) · Ziel: {destinationInfo}", "Success");

            if (importResult.RenamedDueToConflict
                && !string.Equals(importResult.OriginalName, importResult.VmName, StringComparison.OrdinalIgnoreCase))
            {
                AddNotification($"Namenskonflikt erkannt: '{importResult.OriginalName}' wurde automatisch zu '{importResult.VmName}' umbenannt.", "Warning");
            }
        });

        await LoadVmsFromHyperVAsync();
        await RefreshVmStatusAsync();
    }

    private async Task RefreshSwitchesAsync()
    {
        AreSwitchesLoaded = false;

        await ExecuteBusyActionAsync("Switches werden geladen...", async token =>
        {
            var switches = await _hyperVService.GetVmSwitchesAsync(token);

            AvailableSwitches.Clear();
            foreach (var vmSwitch in switches.OrderBy(item => item.Name))
            {
                AvailableSwitches.Add(vmSwitch);
            }

            AreSwitchesLoaded = true;
            SyncSelectedSwitchWithCurrentVm(showNotificationOnMissingSwitch: true);
            NotifyTrayStateChanged();

            AddNotification($"{AvailableSwitches.Count} Switch(es) geladen.", "Info");
        });
    }

    private async Task RefreshRuntimeDataAsync()
    {
        await LoadVmsFromHyperVWithRetryAsync();
        await RefreshSwitchesAsync();
        await RefreshHostNetworkProfileAsync();
        await RefreshVmStatusAsync();
    }

    private async Task RefreshHostNetworkProfileAsync()
    {
        try
        {
            var profileCategory = await _hyperVService.GetHostNetworkProfileCategoryAsync(_lifetimeCancellation.Token);
            ApplyHostNetworkProfileCategory(profileCategory);
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Host-Netzprofil konnte nicht ermittelt werden.");
            ApplyHostNetworkProfileCategory("Unknown");
        }
    }

    private void ApplyHostNetworkProfileCategory(string? category)
    {
        HostNetworkProfileCategory = NormalizeHostNetworkProfileCategory(category);

        HostNetworkProfileDisplayText = HostNetworkProfileCategory switch
        {
            "Public" => "Host-Netzprofil: Öffentlich (Default Switch oft problematisch)",
            "DomainAuthenticated" => "Host-Netzprofil: Domäne",
            "Private" => "Host-Netzprofil: Privat",
            _ => "Host-Netzprofil: Unbekannt"
        };
    }

    private static string NormalizeHostNetworkProfileCategory(string? category)
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

    private async Task RefreshUsbDevicesAsync()
    {
        if (!HostUsbSharingEnabled)
        {
            UsbDevices.Clear();
            SelectedUsbDevice = null;
            UsbStatusText = "USB Share ist global im Host deaktiviert.";
            AddNotification(UsbStatusText, "Info");
            return;
        }

        if (!UsbRuntimeAvailable)
        {
            UsbStatusText = string.IsNullOrWhiteSpace(UsbRuntimeHintText)
                ? "USB ist deaktiviert. usbipd-win ist nicht installiert."
                : UsbRuntimeHintText;
            AddNotification(UsbStatusText, "Warning");
            return;
        }

        await LoadUsbDevicesAsync(showNotification: true);
    }

    private async Task LoadUsbDevicesAsync(bool showNotification, bool applyAutoShare = true, bool useBusyIndicator = true)
    {
        if (!HostUsbSharingEnabled)
        {
            UsbDevices.Clear();
            SelectedUsbDevice = null;
            UsbStatusText = "USB Share ist global im Host deaktiviert.";
            UsbRuntimeHintText = string.Empty;

            if (showNotification)
            {
                AddNotification(UsbStatusText, "Info");
            }

            NotifyTrayStateChanged();
            return;
        }

        var previouslySelectedBusId = SelectedUsbDevice?.BusId;
        UsbStatusText = "USB-Geräte werden geladen...";

        async Task loadAction(CancellationToken token)
        {
            IReadOnlyList<UsbIpDeviceInfo> devices;
            try
            {
                devices = await _usbIpService.GetDevicesAsync(token);

                var canApplyAutoShareNow = applyAutoShare;

                var autoShareCandidates = canApplyAutoShareNow
                    ? devices
                        .Where(device =>
                            !string.IsNullOrWhiteSpace(device.BusId)
                            && !device.IsShared
                            && !device.IsAttached
                            && _usbAutoShareDeviceKeys.Contains(BuildUsbAutoShareKey(device)))
                        .ToList()
                    : [];

                if (autoShareCandidates.Count > 0)
                {
                    var autoShareApplied = 0;
                    var autoShareFailed = 0;

                    foreach (var candidate in autoShareCandidates)
                    {
                        try
                        {
                            await _usbIpService.BindAsync(candidate.BusId, force: false, token);
                            autoShareApplied++;
                        }
                        catch (Exception ex)
                        {
                            autoShareFailed++;
                            Log.Warning(ex, "Auto-Share für USB-Gerät {BusId} fehlgeschlagen.", candidate.BusId);
                        }
                    }

                    devices = await _usbIpService.GetDevicesAsync(token);

                    if (autoShareApplied > 0)
                    {
                        AddNotification($"Auto-Share: {autoShareApplied} USB-Gerät(e) automatisch freigegeben.", "Info");
                    }

                    if (autoShareFailed > 0)
                    {
                        AddNotification($"Auto-Share: {autoShareFailed} USB-Gerät(e) konnten nicht freigegeben werden.", "Warning");
                    }
                }

                if (UsbAutoDetachOnClientDisconnect)
                {
                    var detachedStaleCount = await TryDetachStaleAttachedDevicesWithoutGuestAckAsync(devices, token);
                    if (detachedStaleCount > 0)
                    {
                        devices = await _usbIpService.GetDevicesAsync(token);
                    }
                }
                else
                {
                    _usbAttachedWithoutAckSinceUtc.Clear();
                }
            }
            catch (Exception ex)
            {
                UsbDevices.Clear();
                SelectedUsbDevice = null;
                UsbStatusText = BuildUsbUnavailableStatus(ex.Message);

                if (IsUsbRuntimeMissing(ex.Message))
                {
                    UsbRuntimeAvailable = false;
                    UsbRuntimeHintText = "USB-Funktion deaktiviert: usbipd-win ist nicht installiert. Quelle: https://github.com/dorssel/usbipd-win";
                }

                NotifyTrayStateChanged();

                if (showNotification)
                {
                    AddNotification(UsbStatusText, "Warning");
                }

                return;
            }

            UsbDevices.Clear();
            foreach (var device in devices)
            {
                if (device.IsAttached
                    && !string.IsNullOrWhiteSpace(device.BusId)
                    && UsbGuestConnectionRegistry.TryGetGuestComputerName(device.BusId, out var guestComputerName))
                {
                    device.AttachedGuestComputerName = guestComputerName;
                }

                UsbDevices.Add(device);
            }

            SelectedUsbDevice = UsbDevices.FirstOrDefault(device =>
                                    !string.IsNullOrWhiteSpace(previouslySelectedBusId)
                                    && string.Equals(device.BusId, previouslySelectedBusId, StringComparison.OrdinalIgnoreCase))
                                ?? UsbDevices.FirstOrDefault();

            UsbStatusText = UsbDevices.Count == 0
                ? "Keine USB-Geräte gefunden."
                : $"{UsbDevices.Count} USB-Gerät(e) geladen.";

            UsbRuntimeAvailable = true;
            UsbRuntimeHintText = string.Empty;

            NotifyTrayStateChanged();

            if (showNotification)
            {
                AddNotification(UsbStatusText, UsbDevices.Count == 0 ? "Warning" : "Info");
            }
        }

        if (useBusyIndicator)
        {
            await ExecuteBusyActionAsync("USB-Geräte werden geladen...", loadAction, showNotificationOnErrorOnly: true);
        }
        else
        {
            await loadAction(_lifetimeCancellation.Token);
        }
    }

    private async Task<int> TryDetachStaleAttachedDevicesWithoutGuestAckAsync(IReadOnlyList<UsbIpDeviceInfo> devices, CancellationToken token)
    {
        if (devices.Count == 0)
        {
            _usbAttachedWithoutAckSinceUtc.Clear();
            return 0;
        }

        var now = DateTimeOffset.UtcNow;
        var activeAttachedBusIds = devices
            .Where(device => device.IsAttached && !string.IsNullOrWhiteSpace(device.BusId))
            .Select(device => device.BusId.Trim())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var obsoleteKeys = _usbAttachedWithoutAckSinceUtc.Keys
            .Where(key => !activeAttachedBusIds.Contains(key))
            .ToList();

        foreach (var obsoleteKey in obsoleteKeys)
        {
            _usbAttachedWithoutAckSinceUtc.Remove(obsoleteKey);
        }

        var detachedCount = 0;

        foreach (var device in devices)
        {
            if (!device.IsAttached || string.IsNullOrWhiteSpace(device.BusId))
            {
                continue;
            }

            var busId = device.BusId.Trim();

            if (!string.Equals(device.ClientIpAddress?.Trim(), HyperVSocketUsbTunnelDefaults.LoopbackAddress, StringComparison.OrdinalIgnoreCase))
            {
                _usbAttachedWithoutAckSinceUtc.Remove(busId);
                continue;
            }

            if (UsbGuestConnectionRegistry.TryGetGuestComputerName(busId, out _))
            {
                _usbAttachedWithoutAckSinceUtc.Remove(busId);
                continue;
            }

            if (!_usbAttachedWithoutAckSinceUtc.TryGetValue(busId, out var firstSeenUtc))
            {
                _usbAttachedWithoutAckSinceUtc[busId] = now;
                continue;
            }

            if ((now - firstSeenUtc) < StaleUsbAttachGracePeriod)
            {
                continue;
            }

            try
            {
                await _usbIpService.DetachAsync(busId, token);
                detachedCount++;
                _usbAttachedWithoutAckSinceUtc.Remove(busId);
                Log.Information("Stale attached USB device detached on host (missing guest ack). BusId={BusId}", busId);
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Stale attached USB auto-detach failed. BusId={BusId}", busId);
            }
        }

        return detachedCount;
    }

    private static string BuildUsbUnavailableStatus(string? details)
    {
        if (string.IsNullOrWhiteSpace(details))
        {
            return "USB nicht verfügbar (usbipd fehlt oder Dienst läuft nicht).";
        }

        if (details.Contains("nicht installiert", StringComparison.OrdinalIgnoreCase)
            || details.Contains("nicht gefunden", StringComparison.OrdinalIgnoreCase)
            || details.Contains("winget", StringComparison.OrdinalIgnoreCase))
        {
            return "USB nicht verfügbar: usbipd-win fehlt. Optional im Setup installieren oder manuell von https://github.com/dorssel/usbipd-win";
        }

        if (details.Contains("Dienst", StringComparison.OrdinalIgnoreCase)
            || details.Contains("service", StringComparison.OrdinalIgnoreCase))
        {
            return "USB nicht verfügbar: usbipd-Dienst läuft nicht.";
        }

        return "USB nicht verfügbar. Details im Log/Benachrichtigung prüfen.";
    }

    private static bool IsUsbRuntimeMissing(string? details)
    {
        if (string.IsNullOrWhiteSpace(details))
        {
            return false;
        }

        return details.Contains("nicht installiert", StringComparison.OrdinalIgnoreCase)
               || details.Contains("nicht gefunden", StringComparison.OrdinalIgnoreCase)
               || details.Contains("usbipd-win", StringComparison.OrdinalIgnoreCase)
               || details.Contains("usbipd.exe", StringComparison.OrdinalIgnoreCase)
               || details.Contains("winget", StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildUsbAutoShareKey(UsbIpDeviceInfo device)
    {
        if (!string.IsNullOrWhiteSpace(device.InstanceId))
        {
            return "instance:" + device.InstanceId.Trim();
        }

        if (!string.IsNullOrWhiteSpace(device.HardwareId))
        {
            return "hardware:" + device.HardwareId.Trim();
        }

        if (!string.IsNullOrWhiteSpace(device.Description))
        {
            return "description:" + device.Description.Trim();
        }

        return string.Empty;
    }

    private void PersistUsbAutoShareConfig()
    {
        try
        {
            var configResult = _configService.LoadOrCreate(_configPath);
            configResult.Config.Usb.AutoShareDeviceKeys = _usbAutoShareDeviceKeys
                .OrderBy(key => key, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (!_configService.TrySave(_configPath, configResult.Config, out var errorMessage))
            {
                AddNotification($"USB Auto-Share-Konfiguration konnte nicht gespeichert werden: {errorMessage}", "Warning");
            }
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "USB Auto-Share-Konfiguration konnte nicht gespeichert werden.");
            AddNotification("USB Auto-Share-Konfiguration konnte nicht gespeichert werden.", "Warning");
        }
    }

    private async Task BindSelectedUsbDeviceAsync()
    {
        if (SelectedUsbDevice is null || string.IsNullOrWhiteSpace(SelectedUsbDevice.BusId))
        {
            return;
        }

        var busId = SelectedUsbDevice.BusId;
        if (!IsProcessElevated())
        {
            AddNotification("UAC wird angefordert für USB Share...", "Info");
        }

        await ExecuteBusyActionAsync($"USB-Gerät {busId} wird freigegeben...", async token =>
        {
            await _usbIpService.BindAsync(busId, force: false, token);
            AddNotification($"USB-Gerät '{busId}' wurde freigegeben.", "Success");
        });

        await LoadUsbDevicesAsync(showNotification: false);

        try
        {
            await Task.Delay(TimeSpan.FromSeconds(2), _lifetimeCancellation.Token);
        }
        catch (OperationCanceledException)
        {
            return;
        }

        await LoadUsbDevicesAsync(showNotification: false, useBusyIndicator: false);
    }

    private async Task UnbindSelectedUsbDeviceAsync()
    {
        if (SelectedUsbDevice is null || string.IsNullOrWhiteSpace(SelectedUsbDevice.BusId))
        {
            return;
        }

        var busId = SelectedUsbDevice.BusId;
        if (!IsProcessElevated())
        {
            AddNotification("UAC wird angefordert für USB Unshare...", "Info");
        }

        await ExecuteBusyActionAsync($"USB-Freigabe für {busId} wird entfernt...", async token =>
        {
            await _usbIpService.UnbindAsync(busId, token);
            AddNotification($"USB-Freigabe für '{busId}' wurde entfernt.", "Success");
        });

        await LoadUsbDevicesAsync(showNotification: false, applyAutoShare: false);

        try
        {
            await Task.Delay(TimeSpan.FromSeconds(2), _lifetimeCancellation.Token);
        }
        catch (OperationCanceledException)
        {
            return;
        }

        await LoadUsbDevicesAsync(showNotification: false, applyAutoShare: false, useBusyIndicator: false);
    }

    private async Task DetachSelectedUsbDeviceAsync()
    {
        if (SelectedUsbDevice is null || string.IsNullOrWhiteSpace(SelectedUsbDevice.BusId))
        {
            return;
        }

        var busId = SelectedUsbDevice.BusId;
        if (!IsProcessElevated())
        {
            AddNotification("UAC wird angefordert für USB Detach...", "Info");
        }

        await ExecuteBusyActionAsync($"USB-Gerät {busId} wird getrennt...", async token =>
        {
            await _usbIpService.DetachAsync(busId, token);
            AddNotification($"USB-Gerät '{busId}' wurde getrennt.", "Success");
        });

        await LoadUsbDevicesAsync(showNotification: false);
    }

    private async Task EnsureSelectedVmNetworkSelectionAsync(bool showNotificationOnMissingSwitch)
    {
        if (SelectedVm is null)
        {
            AvailableVmNetworkAdapters.Clear();
            SelectedVmNetworkAdapter = null;
            SelectedSwitch = null;
            NetworkSwitchStatusHint = "Keine VM ausgewählt.";
            return;
        }

        try
        {
            var selectedVmName = SelectedVm.Name;
            var previouslySelectedAdapterName = SelectedVmNetworkAdapter?.Name;
            var adapters = await _hyperVService.GetVmNetworkAdaptersAsync(selectedVmName, _lifetimeCancellation.Token);

            if (SelectedVm is null || !string.Equals(SelectedVm.Name, selectedVmName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            AvailableVmNetworkAdapters.Clear();
            foreach (var adapter in adapters.OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase))
            {
                AvailableVmNetworkAdapters.Add(adapter);
            }

            if (AvailableVmNetworkAdapters.Count == 0)
            {
                SelectedVmNetworkAdapter = null;
                SelectedVmCurrentSwitch = NotConnectedSwitchDisplay;
                SelectedSwitch = null;
                NetworkSwitchStatusHint = "Keine VM-Netzwerkkarten gefunden.";
                ConnectSelectedSwitchCommand.NotifyCanExecuteChanged();
                DisconnectSwitchCommand.NotifyCanExecuteChanged();
                return;
            }

            SelectedVmNetworkAdapter = AvailableVmNetworkAdapters.FirstOrDefault(item =>
                                          string.Equals(item.Name, previouslySelectedAdapterName, StringComparison.OrdinalIgnoreCase))
                                      ?? AvailableVmNetworkAdapters.FirstOrDefault();

            SelectedVmCurrentSwitch = NormalizeSwitchDisplayName(SelectedVmNetworkAdapter?.SwitchName);

            var selectedVmEntry = AvailableVms.FirstOrDefault(vm =>
                string.Equals(vm.Name, selectedVmName, StringComparison.OrdinalIgnoreCase));

            if (selectedVmEntry is not null)
            {
                selectedVmEntry.RuntimeSwitchName = SelectedVmCurrentSwitch;
            }

            SyncSelectedSwitchWithCurrentVm(showNotificationOnMissingSwitch);
            NotifyTrayStateChanged();
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Netzwerkkarten für VM {VmName} konnten nicht gelesen werden.", SelectedVm.Name);
        }
    }

    private void SyncSelectedSwitchWithCurrentVm(bool showNotificationOnMissingSwitch)
    {
        if (!AreSwitchesLoaded)
        {
            SelectedSwitch = null;
            NetworkSwitchStatusHint = "Switch-Liste noch nicht geladen.";
            ConnectSelectedSwitchCommand.NotifyCanExecuteChanged();
            return;
        }

        if (SelectedVmNetworkAdapter is null)
        {
            SelectedSwitch = null;
            NetworkSwitchStatusHint = "Bitte VM-Netzwerkkarte auswählen.";
            ConnectSelectedSwitchCommand.NotifyCanExecuteChanged();
            return;
        }

        if (AvailableSwitches.Count == 0)
        {
            SelectedSwitch = null;
            NetworkSwitchStatusHint = "Keine Switches verfügbar.";
            ConnectSelectedSwitchCommand.NotifyCanExecuteChanged();
            return;
        }

        if (IsNotConnectedSwitchDisplay(SelectedVmCurrentSwitch))
        {
            SelectedSwitch = null;
            NetworkSwitchStatusHint = $"{GetAdapterDisplayName(SelectedVmNetworkAdapter)}: {NotConnectedSwitchDisplay}";
            ConnectSelectedSwitchCommand.NotifyCanExecuteChanged();
            return;
        }

        var matchingSwitch = AvailableSwitches.FirstOrDefault(item =>
            string.Equals(item.Name, SelectedVmCurrentSwitch, StringComparison.OrdinalIgnoreCase));

        if (matchingSwitch is null)
        {
            SelectedSwitch = null;
            NetworkSwitchStatusHint = $"Aktiver Switch '{SelectedVmCurrentSwitch}' für '{GetAdapterDisplayName(SelectedVmNetworkAdapter)}' ist nicht in der aktuellen Liste.";
            if (showNotificationOnMissingSwitch)
            {
                AddNotification(NetworkSwitchStatusHint, "Warning");
            }

            ConnectSelectedSwitchCommand.NotifyCanExecuteChanged();
            return;
        }

        SelectedSwitch = matchingSwitch;
        NetworkSwitchStatusHint = $"Aktiver Switch ({GetAdapterDisplayName(SelectedVmNetworkAdapter)}): {matchingSwitch.Name}";
        ConnectSelectedSwitchCommand.NotifyCanExecuteChanged();
    }

    private static string NormalizeSwitchDisplayName(string? switchName)
    {
        return string.IsNullOrWhiteSpace(switchName) ? NotConnectedSwitchDisplay : switchName.Trim();
    }

    private static bool IsNotConnectedSwitchDisplay(string? switchName)
    {
        return string.IsNullOrWhiteSpace(switchName)
               || string.Equals(switchName, "-", StringComparison.Ordinal)
               || string.Equals(switchName, NotConnectedSwitchDisplay, StringComparison.OrdinalIgnoreCase);
    }

    private static string GetAdapterDisplayName(HyperVVmNetworkAdapterInfo? adapter)
    {
        if (adapter is null)
        {
            return "Network Adapter";
        }

        return string.IsNullOrWhiteSpace(adapter.Name) ? "Network Adapter" : adapter.Name;
    }

    private async Task ConnectSelectedSwitchAsync()
    {
        if (SelectedSwitch is null || SelectedVm is null || SelectedVmNetworkAdapter is null)
        {
            return;
        }

        await ExecuteBusyActionAsync("VM-Netzwerk wird verbunden...", async token =>
        {
            await _hyperVService.ConnectVmNetworkAdapterAsync(SelectedVm.Name, SelectedSwitch.Name, SelectedVmNetworkAdapter.Name, token);
            AddNotification($"'{SelectedVm.Name}' Adapter '{GetAdapterDisplayName(SelectedVmNetworkAdapter)}' mit '{SelectedSwitch.Name}' verbunden.", "Success");

            if (ShouldAutoRestartHnsAfterConnect(SelectedSwitch.Name))
            {
                var hnsResult = await _hnsService.RestartHnsElevatedAsync(token);
                AddNotification(
                    hnsResult.Success ? hnsResult.Message : $"HNS Neustart fehlgeschlagen: {hnsResult.Message}",
                    hnsResult.Success ? "Success" : "Error");
            }
        });
        await RefreshVmStatusAsync();
    }

    private async Task DisconnectSwitchAsync()
    {
        if (SelectedVm is null || SelectedVmNetworkAdapter is null)
        {
            return;
        }

        await ExecuteBusyActionAsync("VM-Netzwerk wird getrennt...", async token =>
        {
            await _hyperVService.DisconnectVmNetworkAdapterAsync(SelectedVm.Name, SelectedVmNetworkAdapter.Name, token);
            AddNotification($"Netzwerkkarte '{GetAdapterDisplayName(SelectedVmNetworkAdapter)}' von '{SelectedVm.Name}' getrennt.", "Warning");
        });
        await RefreshVmStatusAsync();
    }

    private async Task ConnectAdapterToSwitchByKeyAsync(string? adapterSwitchKey)
    {
        if (SelectedVm is null || string.IsNullOrWhiteSpace(adapterSwitchKey))
        {
            return;
        }

        var separatorIndex = adapterSwitchKey.IndexOf("|||", StringComparison.Ordinal);
        if (separatorIndex <= 0 || separatorIndex >= adapterSwitchKey.Length - 3)
        {
            return;
        }

        var adapterName = adapterSwitchKey[..separatorIndex].Trim();
        var switchName = adapterSwitchKey[(separatorIndex + 3)..].Trim();
        if (string.IsNullOrWhiteSpace(adapterName) || string.IsNullOrWhiteSpace(switchName))
        {
            return;
        }

        var targetAdapter = AvailableVmNetworkAdapters.FirstOrDefault(item => string.Equals(item.Name, adapterName, StringComparison.OrdinalIgnoreCase));
        var targetSwitch = AvailableSwitches.FirstOrDefault(item => string.Equals(item.Name, switchName, StringComparison.OrdinalIgnoreCase));
        if (targetAdapter is null || targetSwitch is null)
        {
            return;
        }

        SelectedVmNetworkAdapter = targetAdapter;
        SelectedSwitch = targetSwitch;
        await ConnectSelectedSwitchAsync();
    }

    private async Task DisconnectAdapterByNameAsync(string? adapterName)
    {
        if (SelectedVm is null || string.IsNullOrWhiteSpace(adapterName))
        {
            return;
        }

        var targetAdapter = AvailableVmNetworkAdapters.FirstOrDefault(item => string.Equals(item.Name, adapterName, StringComparison.OrdinalIgnoreCase));
        if (targetAdapter is null)
        {
            return;
        }

        SelectedVmNetworkAdapter = targetAdapter;
        await DisconnectSwitchAsync();
    }

    private bool ShouldAutoRestartHnsAfterConnect(string connectedSwitch)
    {
        if (HnsAutoRestartAfterAnyConnect)
        {
            return true;
        }

        return HnsAutoRestartAfterDefaultSwitch
               && string.Equals(connectedSwitch, DefaultSwitchName, StringComparison.OrdinalIgnoreCase);
    }

    private async Task StartDefaultVmAsync()
    {
        var targetVm = AvailableVms.FirstOrDefault(vm => string.Equals(vm.Name, DefaultVmName, StringComparison.OrdinalIgnoreCase))?.Name
                       ?? SelectedVm?.Name
                       ?? AvailableVms.FirstOrDefault()?.Name;
        if (string.IsNullOrWhiteSpace(targetVm))
        {
            AddNotification("Keine VM zum Starten gefunden.", "Error");
            return;
        }

        await ExecuteBusyActionAsync("Default VM wird gestartet...", async token =>
        {
            await _hyperVService.StartVmAsync(targetVm, token);
            AddNotification($"'{targetVm}' gestartet.", "Success");

            if (UiOpenConsoleAfterVmStart)
            {
                await _hyperVService.OpenVmConnectAsync(targetVm, VmConnectComputerName, ShouldOpenConsoleWithSessionEdit(targetVm), token);
                AddNotification($"Konsole für '{targetVm}' geöffnet.", "Info");
            }
        });
        await RefreshVmStatusAsync();
    }

    private async Task StopDefaultVmAsync()
    {
        var targetVm = AvailableVms.FirstOrDefault(vm => string.Equals(vm.Name, DefaultVmName, StringComparison.OrdinalIgnoreCase))?.Name
                       ?? SelectedVm?.Name
                       ?? AvailableVms.FirstOrDefault()?.Name;
        if (string.IsNullOrWhiteSpace(targetVm))
        {
            AddNotification("Keine VM zum Stoppen gefunden.", "Error");
            return;
        }

        await ExecuteBusyActionAsync("Default VM wird gestoppt...", async token =>
        {
            await _hyperVService.StopVmGracefulAsync(targetVm, token);
            AddNotification($"'{targetVm}' gestoppt.", "Success");
        });
        await RefreshVmStatusAsync();
    }

    private async Task RefreshVmStatusAsync()
    {
        if (SelectedVm is null)
        {
            return;
        }

        await ExecuteBusyActionAsync("VM-Status wird aktualisiert...", async token =>
        {
            var vms = await _hyperVService.GetVmsAsync(token);
            UpdateVmRuntimeStates(vms);

            var vmInfo = vms.FirstOrDefault(item => string.Equals(item.Name, SelectedVm.Name, StringComparison.OrdinalIgnoreCase));
            SelectedVmState = vmInfo?.State ?? "Unbekannt";
            SelectedVmCurrentSwitch = NormalizeSwitchDisplayName(vmInfo?.CurrentSwitchName);
            NotifyTrayStateChanged();
            StatusText = vmInfo is null ? "VM nicht gefunden" : $"{vmInfo.Name}: {vmInfo.State}";
        }, showNotificationOnErrorOnly: true);

        await EnsureSelectedVmNetworkSelectionAsync(showNotificationOnMissingSwitch: false);
    }

    private void UpdateVmRuntimeStates(IReadOnlyList<HyperVVmInfo> runtimeVms)
    {
        if (runtimeVms.Count == 0 || AvailableVms.Count == 0)
        {
            return;
        }

        var labelsByName = AvailableVms
            .GroupBy(vm => vm.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First().Label, StringComparer.OrdinalIgnoreCase);

        var trayAdapterByName = AvailableVms
            .GroupBy(vm => vm.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First().TrayAdapterName, StringComparer.OrdinalIgnoreCase);

        var sessionEditByName = AvailableVms
            .GroupBy(vm => vm.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First().OpenConsoleWithSessionEdit, StringComparer.OrdinalIgnoreCase);

        var selectedName = SelectedVm?.Name;
        var defaultName = DefaultVmName;

        var rebuilt = runtimeVms
            .OrderBy(vm => vm.Name, StringComparer.OrdinalIgnoreCase)
            .Select(vm => new VmDefinition
            {
                Name = vm.Name,
                Label = labelsByName.TryGetValue(vm.Name, out var label) && !string.IsNullOrWhiteSpace(label) ? label : vm.Name,
                RuntimeState = vm.State,
                RuntimeSwitchName = NormalizeSwitchDisplayName(vm.CurrentSwitchName),
                TrayAdapterName = trayAdapterByName.TryGetValue(vm.Name, out var trayAdapter) ? trayAdapter : string.Empty,
                OpenConsoleWithSessionEdit = sessionEditByName.TryGetValue(vm.Name, out var openWithSessionEdit) && openWithSessionEdit
            })
            .ToList();

        AvailableVms.Clear();
        foreach (var vm in rebuilt)
        {
            AvailableVms.Add(vm);
        }

        SetSelectedVmInternal(AvailableVms.FirstOrDefault(vm => string.Equals(vm.Name, selectedName, StringComparison.OrdinalIgnoreCase))
                      ?? AvailableVms.FirstOrDefault());
        SelectedVmForConfig = SelectedVm;
        SelectedDefaultVmForConfig = AvailableVms.FirstOrDefault(vm => string.Equals(vm.Name, defaultName, StringComparison.OrdinalIgnoreCase))
                                   ?? SelectedVm;
        NotifyTrayStateChanged();
    }

    private async Task LoadCheckpointsAsync()
    {
        if (SelectedVm is null)
        {
            return;
        }

        var vmName = SelectedVm.Name;

        try
        {
            var checkpoints = await _hyperVService.GetCheckpointsAsync(vmName, _lifetimeCancellation.Token);

            if (SelectedVm is null
                || !string.Equals(SelectedVm.Name, vmName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            AvailableCheckpoints.Clear();
            foreach (var checkpoint in checkpoints.OrderByDescending(item => item.Created))
            {
                AvailableCheckpoints.Add(checkpoint);
            }

            RebuildCheckpointTree(checkpoints);

            var newestCheckpoint = checkpoints
                .OrderByDescending(item => item.Created)
                .ThenByDescending(item => item.Name, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();

            if (newestCheckpoint is null)
            {
                SelectedCheckpointNode = null;
                SelectedCheckpoint = null;
            }
            else
            {
                SelectedCheckpointNode = FindCheckpointNodeById(newestCheckpoint.Id);
                SelectedCheckpoint = newestCheckpoint;
            }

            AddNotification($"{AvailableCheckpoints.Count} Checkpoint(s) für '{vmName}' geladen.", "Info");
        }
        catch (UnauthorizedAccessException ex)
        {
            Log.Warning(ex, "Checkpoint laden fehlgeschlagen (Berechtigung) für VM {VmName}", vmName);
            AddNotification(ex.Message, "Warning");
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Checkpoint laden fehlgeschlagen für VM {VmName}", vmName);
            AddNotification($"Fehler beim Laden der Checkpoints: {ex.Message}", "Error");
        }
    }

    private async Task CreateCheckpointAsync()
    {
        if (SelectedVm is null)
        {
            return;
        }

        var checkpointName = string.IsNullOrWhiteSpace(NewCheckpointName)
            ? $"checkpoint-{DateTime.Now:yyyyMMdd-HHmmss}"
            : NewCheckpointName.Trim();

        await ExecuteBusyActionAsync("Checkpoint wird erstellt...", async token =>
        {
            await _hyperVService.CreateCheckpointAsync(
                SelectedVm.Name,
                checkpointName,
                string.IsNullOrWhiteSpace(NewCheckpointDescription) ? null : NewCheckpointDescription.Trim(),
                token);

            AddNotification($"Checkpoint '{checkpointName}' für '{SelectedVm.Name}' erstellt.", "Success");
        });

        NewCheckpointName = string.Empty;
        NewCheckpointDescription = string.Empty;
        await LoadCheckpointsAsync();
    }

    private async Task ApplyCheckpointAsync()
    {
        if (SelectedVm is null || SelectedCheckpoint is null)
        {
            return;
        }

        await ExecuteBusyActionAsync("Checkpoint wird wiederhergestellt...", async token =>
        {
            await _hyperVService.ApplyCheckpointAsync(SelectedVm.Name, SelectedCheckpoint.Name, SelectedCheckpoint.Id, token);
            AddNotification($"Checkpoint '{SelectedCheckpoint.Name}' auf '{SelectedVm.Name}' wiederhergestellt.", "Warning");
        });

        await RefreshVmStatusAsync();
        await LoadCheckpointsAsync();
    }

    private async Task DeleteCheckpointAsync()
    {
        if (SelectedVm is null || SelectedCheckpoint is null)
        {
            return;
        }

        var checkpointName = SelectedCheckpoint.Name;
        var checkpointId = SelectedCheckpoint.Id;
        await ExecuteBusyActionAsync("Checkpoint wird gelöscht...", async token =>
        {
            await _hyperVService.RemoveCheckpointAsync(SelectedVm.Name, checkpointName, checkpointId, token);
            AddNotification($"Checkpoint '{checkpointName}' von '{SelectedVm.Name}' gelöscht.", "Warning");
        });

        await LoadCheckpointsAsync();
    }

    private void RebuildCheckpointTree(IReadOnlyList<HyperVCheckpointInfo> checkpoints)
    {
        AvailableCheckpointTree.Clear();
        if (checkpoints.Count == 0)
        {
            return;
        }

        var latestId = checkpoints
            .OrderByDescending(item => item.Created)
            .ThenByDescending(item => item.Name, StringComparer.OrdinalIgnoreCase)
            .Select(item => item.Id)
            .FirstOrDefault() ?? string.Empty;

        var nodesById = new Dictionary<string, HyperVCheckpointTreeItem>(StringComparer.OrdinalIgnoreCase);
        var allNodes = checkpoints
            .Select(checkpoint => new HyperVCheckpointTreeItem
            {
                Checkpoint = checkpoint,
                IsLatest = !string.IsNullOrWhiteSpace(checkpoint.Id)
                           && string.Equals(checkpoint.Id, latestId, StringComparison.OrdinalIgnoreCase)
            })
            .ToList();

        foreach (var node in allNodes)
        {
            if (!string.IsNullOrWhiteSpace(node.Checkpoint.Id)
                && !nodesById.ContainsKey(node.Checkpoint.Id))
            {
                nodesById.Add(node.Checkpoint.Id, node);
            }
        }

        var rootNodes = new List<HyperVCheckpointTreeItem>();
        foreach (var node in allNodes)
        {
            var parentId = node.Checkpoint.ParentId;
            if (!string.IsNullOrWhiteSpace(parentId)
                && nodesById.TryGetValue(parentId, out var parentNode)
                && !ReferenceEquals(parentNode, node))
            {
                parentNode.Children.Add(node);
            }
            else
            {
                rootNodes.Add(node);
            }
        }

        foreach (var root in rootNodes
                     .OrderBy(item => item.Created)
                     .ThenBy(item => item.Name, StringComparer.OrdinalIgnoreCase))
        {
            SortCheckpointTreeRecursively(root);
            AvailableCheckpointTree.Add(root);
        }
    }

    private static void SortCheckpointTreeRecursively(HyperVCheckpointTreeItem node)
    {
        if (node.Children.Count == 0)
        {
            return;
        }

        var orderedChildren = node.Children
            .OrderBy(item => item.Created)
            .ThenBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        node.Children.Clear();
        foreach (var child in orderedChildren)
        {
            SortCheckpointTreeRecursively(child);
            node.Children.Add(child);
        }
    }

    private HyperVCheckpointTreeItem? FindCheckpointNodeById(string checkpointId)
    {
        if (string.IsNullOrWhiteSpace(checkpointId))
        {
            return null;
        }

        foreach (var root in AvailableCheckpointTree)
        {
            var found = FindCheckpointNodeByIdRecursive(root, checkpointId);
            if (found is not null)
            {
                return found;
            }
        }

        return null;
    }

    private static HyperVCheckpointTreeItem? FindCheckpointNodeByIdRecursive(HyperVCheckpointTreeItem node, string checkpointId)
    {
        if (string.Equals(node.Checkpoint.Id, checkpointId, StringComparison.OrdinalIgnoreCase))
        {
            return node;
        }

        foreach (var child in node.Children)
        {
            var found = FindCheckpointNodeByIdRecursive(child, checkpointId);
            if (found is not null)
            {
                return found;
            }
        }

        return null;
    }

    private void AddVm()
    {
        var vmName = NewVmName.Trim();
        if (string.IsNullOrWhiteSpace(vmName))
        {
            AddNotification("VM-Name darf nicht leer sein.", "Error");
            return;
        }

        if (AvailableVms.Any(vm => string.Equals(vm.Name, vmName, StringComparison.OrdinalIgnoreCase)))
        {
            AddNotification("VM existiert bereits in der Konfiguration.", "Warning");
            return;
        }

        var vmLabel = string.IsNullOrWhiteSpace(NewVmLabel) ? vmName : NewVmLabel.Trim();
        var vm = new VmDefinition
        {
            Name = vmName,
            Label = vmLabel
        };

        AvailableVms.Add(vm);
        SelectedVmForConfig = vm;
        SelectedDefaultVmForConfig ??= vm;

        NewVmName = string.Empty;
        NewVmLabel = string.Empty;
        MarkConfigDirty();
        AddNotification($"VM '{vmName}' zur Konfiguration hinzugefügt.", "Success");
    }

    private void RemoveSelectedVm()
    {
        if (SelectedVmForConfig is null)
        {
            AddNotification("Keine VM zum Entfernen ausgewählt.", "Warning");
            return;
        }

        var vmToRemove = SelectedVmForConfig;
        AvailableVms.Remove(vmToRemove);

        if (SelectedVm == vmToRemove)
        {
            SelectedVm = AvailableVms.FirstOrDefault();
        }

        if (SelectedDefaultVmForConfig == vmToRemove)
        {
            SelectedDefaultVmForConfig = AvailableVms.FirstOrDefault();
            DefaultVmName = SelectedDefaultVmForConfig?.Name ?? string.Empty;
        }

        SelectedVmForConfig = AvailableVms.FirstOrDefault();
        MarkConfigDirty();
        AddNotification($"VM '{vmToRemove.Name}' aus Konfiguration entfernt.", "Warning");
    }

    private void SetDefaultVmFromSelection()
    {
        var targetVm = SelectedVmForConfig ?? SelectedDefaultVmForConfig;
        if (targetVm is null)
        {
            AddNotification("Keine Default-VM ausgewählt.", "Warning");
            return;
        }

        DefaultVmName = targetVm.Name;
        SelectedDefaultVmForConfig = targetVm;
        MarkConfigDirty();
        AddNotification($"Default VM gesetzt: '{DefaultVmName}'.", "Info");
    }

    private async Task LoadVmAdaptersForConfigAsync(VmDefinition? vm)
    {
        AvailableVmTrayAdapterOptions.Clear();
        AvailableVmAdaptersForRename.Clear();
        SelectedVmTrayAdapterOption = null;
        SelectedVmAdapterForRename = null;
        NewVmAdapterName = string.Empty;

        if (vm is null || string.IsNullOrWhiteSpace(vm.Name))
        {
            return;
        }

        try
        {
            var adapters = await _hyperVService.GetVmNetworkAdaptersAsync(vm.Name, _lifetimeCancellation.Token);

            if (SelectedVmForConfig is null || !string.Equals(SelectedVmForConfig.Name, vm.Name, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            var orderedAdapters = adapters
                .OrderBy(item => item.DisplayName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var option in BuildTrayAdapterOptions(orderedAdapters))
            {
                AvailableVmTrayAdapterOptions.Add(option);
            }

            foreach (var adapter in orderedAdapters)
            {
                AvailableVmAdaptersForRename.Add(adapter);
            }

            var initialTrayAdapterOption = AvailableVmTrayAdapterOptions.FirstOrDefault(option =>
                                               string.Equals(option.AdapterName, vm.TrayAdapterName, StringComparison.OrdinalIgnoreCase))
                                           ?? AvailableVmTrayAdapterOptions.FirstOrDefault();

            var initialAdapterForRename = AvailableVmAdaptersForRename.FirstOrDefault(option =>
                string.Equals(option.Name, vm.TrayAdapterName, StringComparison.OrdinalIgnoreCase))
                ?? AvailableVmAdaptersForRename.FirstOrDefault();

            _configChangeSuppressionDepth++;
            try
            {
                SelectedVmTrayAdapterOption = initialTrayAdapterOption;
                SelectedVmAdapterForRename = initialAdapterForRename;
            }
            finally
            {
                _configChangeSuppressionDepth--;
            }
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Adapterliste für Config-VM {VmName} konnte nicht geladen werden.", vm.Name);
        }

        UpdateVmAdapterRenameValidationState();
        RenameVmAdapterCommand.NotifyCanExecuteChanged();
    }

    private bool CanExecuteRenameVmAdapter() => !IsBusy && string.IsNullOrWhiteSpace(ValidateVmAdapterRenameInput());

    private void UpdateVmAdapterRenameValidationState()
    {
        VmAdapterRenameValidationMessage = ValidateVmAdapterRenameInput();
    }

    private string ValidateVmAdapterRenameInput()
    {
        if (SelectedVmForConfig is null)
        {
            return "";
        }

        if (SelectedVmAdapterForRename is null)
        {
            return "Bitte zuerst einen Adapter auswählen.";
        }

        var oldName = SelectedVmAdapterForRename.Name?.Trim() ?? string.Empty;
        var newName = NewVmAdapterName?.Trim() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(newName))
        {
            return "Neuer Adaptername darf nicht leer sein.";
        }

        if (string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase))
        {
            return "Neuer Name ist identisch mit dem aktuellen Namen.";
        }

        if (newName.IndexOfAny(VmAdapterInvalidNameChars) >= 0)
        {
            return "Der Name enthält ungültige Zeichen: \\ / : * ? \" < > |";
        }

        if (AvailableVmAdaptersForRename.Any(adapter =>
                !string.Equals(adapter.Name, oldName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(adapter.Name, newName, StringComparison.OrdinalIgnoreCase)))
        {
            return "Ein Adapter mit diesem Namen existiert bereits auf der VM.";
        }

        return "";
    }

    private static IReadOnlyList<VmTrayAdapterOption> BuildTrayAdapterOptions(IReadOnlyList<HyperVVmNetworkAdapterInfo> adapters)
    {
        var options = new List<VmTrayAdapterOption>
        {
            new()
            {
                AdapterName = string.Empty,
                DisplayName = "Alle Adapter (Standard)"
            }
        };

        options.AddRange(adapters.Select(adapter => new VmTrayAdapterOption
        {
            AdapterName = adapter.Name,
            DisplayName = adapter.DisplayName
        }));

        return options;
    }

    private async Task RenameVmAdapterAsync()
    {
        if (SelectedVmForConfig is null || SelectedVmAdapterForRename is null)
        {
            return;
        }

        var validationMessage = ValidateVmAdapterRenameInput();
        if (!string.IsNullOrWhiteSpace(validationMessage))
        {
            AddNotification($"Adapter umbenennen abgebrochen: {validationMessage}", "Warning");
            return;
        }

        var vmName = SelectedVmForConfig.Name;
        var oldName = SelectedVmAdapterForRename.Name?.Trim() ?? string.Empty;
        var newName = NewVmAdapterName?.Trim() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName))
        {
            return;
        }

        await ExecuteBusyActionAsync($"Adapter '{oldName}' wird umbenannt...", async token =>
        {
            await _hyperVService.RenameVmNetworkAdapterAsync(vmName, oldName, newName, token);

            if (string.Equals(SelectedVmForConfig.TrayAdapterName, oldName, StringComparison.OrdinalIgnoreCase))
            {
                SelectedVmForConfig.TrayAdapterName = newName;
            }

            var liveVm = AvailableVms.FirstOrDefault(vm => string.Equals(vm.Name, vmName, StringComparison.OrdinalIgnoreCase));
            if (liveVm is not null && string.Equals(liveVm.TrayAdapterName, oldName, StringComparison.OrdinalIgnoreCase))
            {
                liveVm.TrayAdapterName = newName;
            }

            AddNotification($"VM-Adapter '{oldName}' wurde in '{newName}' umbenannt.", "Success");
        });

        NewVmAdapterName = string.Empty;
        await LoadVmAdaptersForConfigAsync(SelectedVmForConfig);
        await RefreshVmStatusAsync();
        NotifyTrayStateChanged();
    }

    private async Task SaveConfigAsync()
    {
        await ExecuteBusyActionAsync("Konfiguration wird gespeichert...", _ =>
        {
            var config = new HyperToolConfig
            {
                DefaultVmName = DefaultVmName,
                LastSelectedVmName = SelectedVm?.Name ?? LastSelectedVmName,
                Vms = AvailableVms
                    .Select(vm => new VmDefinition
                    {
                        Name = vm.Name,
                        Label = string.IsNullOrWhiteSpace(vm.Label) ? vm.Name : vm.Label,
                        TrayAdapterName = vm.TrayAdapterName?.Trim() ?? string.Empty,
                        OpenConsoleWithSessionEdit = vm.OpenConsoleWithSessionEdit
                    })
                    .OrderBy(vm => vm.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList(),
                DefaultSwitchName = DefaultSwitchName,
                VmConnectComputerName = NormalizeVmConnectComputerName(VmConnectComputerName),
                Hns = new HnsSettings
                {
                    Enabled = HnsEnabled,
                    AutoRestartAfterDefaultSwitch = HnsAutoRestartAfterDefaultSwitch,
                    AutoRestartAfterAnyConnect = HnsAutoRestartAfterAnyConnect
                },
                Ui = new UiSettings
                {
                    WindowTitle = "HyperTool",
                    Theme = NormalizeUiTheme(UiTheme),
                    StartMinimized = UiStartMinimized,
                    MinimizeToTray = true,
                    EnableTrayIcon = true,
                    EnableTrayMenu = UiEnableTrayMenu,
                    StartWithWindows = UiStartWithWindows,
                    OpenConsoleAfterVmStart = UiOpenConsoleAfterVmStart,
                    RestoreNumLockAfterVmStart = UiRestoreNumLockAfterVmStart,
                    NumLockWatcherIntervalSeconds = Math.Clamp(_uiNumLockWatcherIntervalSeconds, 5, 600),
                    OpenVmConnectWithSessionEdit = UiOpenVmConnectWithSessionEdit,
                    TrayVmNames = [.. _trayVmNames]
                },
                Update = new UpdateSettings
                {
                    CheckOnStartup = UpdateCheckOnStartup,
                    GitHubOwner = GithubOwner,
                    GitHubRepo = GithubRepo
                },
                Usb = new UsbSettings
                {
                    Enabled = HostUsbSharingEnabled,
                    AutoDetachOnClientDisconnect = UsbAutoDetachOnClientDisconnect,
                    UnshareOnExit = UsbUnshareOnExit,
                    AutoShareDeviceKeys = _usbAutoShareDeviceKeys
                        .OrderBy(key => key, StringComparer.OrdinalIgnoreCase)
                        .ToList()
                },
                SharedFolders = new SharedFolderSettings
                {
                    Enabled = HostSharedFoldersEnabled,
                    HostDefinitions = HostSharedFolders
                        .Select(CloneSharedFolderDefinition)
                        .OrderBy(folder => folder.Label, StringComparer.OrdinalIgnoreCase)
                        .ThenBy(folder => folder.ShareName, StringComparer.OrdinalIgnoreCase)
                        .ToList()
                },
                DefaultVmImportDestinationPath = DefaultVmImportDestinationPath.Trim()
            };

            if (_configService.TrySave(_configPath, config, out var errorMessage))
            {
                ApplyConfiguredVmDefinitions(config.Vms);

                var executablePath = Environment.ProcessPath ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(executablePath)
                    && !_startupService.SetStartWithWindows(UiStartWithWindows, "HyperTool", executablePath, out var startupError))
                {
                    AddNotification($"Autostart konnte nicht gesetzt werden: {startupError}", "Warning");
                }

                AddNotification("Konfiguration gespeichert.", "Success");
                MarkConfigClean();
            }
            else
            {
                AddNotification($"Konfiguration nicht gespeichert: {errorMessage}", "Error");
            }

            return Task.CompletedTask;
        });
    }

    private async Task RestartHnsAsync()
    {
        await ExecuteBusyActionAsync("HNS wird mit UAC neu gestartet...", async token =>
        {
            var result = await _hnsService.RestartHnsElevatedAsync(token);
            if (result.Success)
            {
                AddNotification(result.Message, "Success");
                return;
            }

            AddNotification($"HNS Neustart fehlgeschlagen: {result.Message}", "Error");
        });
    }

    public IReadOnlyList<VmDefinition> GetTrayVms()
    {
        IEnumerable<VmDefinition> trayVms = AvailableVms;

        if (_trayVmNames.Count > 0)
        {
            var allowedVmNames = new HashSet<string>(_trayVmNames, StringComparer.OrdinalIgnoreCase);
            trayVms = trayVms.Where(vm => allowedVmNames.Contains(vm.Name));
        }

        return trayVms
            .OrderByDescending(vm => string.Equals(vm.Name, DefaultVmName, StringComparison.OrdinalIgnoreCase))
            .ThenBy(vm => vm.DisplayLabel, StringComparer.OrdinalIgnoreCase)
            .Select(vm => new VmDefinition
            {
                Name = vm.Name,
                Label = vm.Label,
                TrayAdapterName = vm.TrayAdapterName,
                OpenConsoleWithSessionEdit = vm.OpenConsoleWithSessionEdit,
                RuntimeState = vm.RuntimeState,
                RuntimeSwitchName = vm.RuntimeSwitchName
            })
            .ToList();
    }

    public IReadOnlyList<HostSharedFolderDefinition> GetHostSharedFoldersSnapshot()
    {
        if (!HostSharedFoldersEnabled)
        {
            return [];
        }

        return HostSharedFolders
            .Select(CloneSharedFolderDefinition)
            .ToList();
    }

    public void UpsertHostSharedFolderDefinition(HostSharedFolderDefinition definition)
    {
        ArgumentNullException.ThrowIfNull(definition);

        var normalized = NormalizeSharedFolderDefinition(definition);
        if (string.IsNullOrWhiteSpace(normalized.Id))
        {
            normalized.Id = Guid.NewGuid().ToString("N");
        }

        var existingIndex = -1;
        for (var index = 0; index < HostSharedFolders.Count; index++)
        {
            if (!string.Equals(HostSharedFolders[index].Id, normalized.Id, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            existingIndex = index;
            break;
        }

        if (existingIndex >= 0)
        {
            HostSharedFolders[existingIndex] = normalized;
        }
        else
        {
            HostSharedFolders.Add(normalized);
        }

        MarkConfigDirty();
    }

    public bool RemoveHostSharedFolderDefinition(string id)
    {
        var normalizedId = (id ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalizedId))
        {
            return false;
        }

        for (var index = 0; index < HostSharedFolders.Count; index++)
        {
            if (!string.Equals(HostSharedFolders[index].Id, normalizedId, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            HostSharedFolders.RemoveAt(index);
            MarkConfigDirty();
            return true;
        }

        return false;
    }

    public IReadOnlyList<HyperVSwitchInfo> GetTraySwitches()
    {
        return AvailableSwitches
            .Select(vmSwitch => new HyperVSwitchInfo
            {
                Name = vmSwitch.Name,
                SwitchType = vmSwitch.SwitchType
            })
            .ToList();
    }

    public UsbIpDeviceInfo? GetSelectedUsbDeviceForTray()
    {
        if (SelectedUsbDevice is null)
        {
            return null;
        }

        return new UsbIpDeviceInfo
        {
            BusId = SelectedUsbDevice.BusId,
            Description = SelectedUsbDevice.Description,
            HardwareId = SelectedUsbDevice.HardwareId,
            InstanceId = SelectedUsbDevice.InstanceId,
            PersistedGuid = SelectedUsbDevice.PersistedGuid,
            ClientIpAddress = SelectedUsbDevice.ClientIpAddress,
            AttachedGuestComputerName = SelectedUsbDevice.AttachedGuestComputerName
        };
    }

    public IReadOnlyList<UsbIpDeviceInfo> GetUsbDevicesForTray()
    {
        if (!HostUsbSharingEnabled)
        {
            return [];
        }

        return UsbDevices
            .Select(device => new UsbIpDeviceInfo
            {
                BusId = device.BusId,
                Description = device.Description,
                HardwareId = device.HardwareId,
                InstanceId = device.InstanceId,
                PersistedGuid = device.PersistedGuid,
                ClientIpAddress = device.ClientIpAddress,
                AttachedGuestComputerName = device.AttachedGuestComputerName
            })
            .ToList();
    }

    public Task SelectUsbDeviceForTrayAsync(string selectionKey)
    {
        if (string.IsNullOrWhiteSpace(selectionKey))
        {
            return Task.CompletedTask;
        }

        var key = selectionKey.Trim();

        var target = UsbDevices.FirstOrDefault(device =>
            string.Equals(BuildUsbSelectionKey(device), key, StringComparison.OrdinalIgnoreCase));

        if (target is not null)
        {
            SelectedUsbDevice = target;
        }

        return Task.CompletedTask;
    }

    private static string BuildUsbSelectionKey(UsbIpDeviceInfo device)
    {
        if (!string.IsNullOrWhiteSpace(device.BusId))
        {
            return "busid:" + device.BusId.Trim();
        }

        if (!string.IsNullOrWhiteSpace(device.PersistedGuid))
        {
            return "guid:" + device.PersistedGuid.Trim();
        }

        return "instance:" + (device.InstanceId?.Trim() ?? string.Empty);
    }

    public async Task RefreshUsbDevicesFromTrayAsync()
    {
        if (!HostUsbSharingEnabled)
        {
            return;
        }

        var gateEntered = false;
        try
        {
            await _usbTrayRefreshGate.WaitAsync(_lifetimeCancellation.Token);
            gateEntered = true;
            await LoadUsbDevicesAsync(showNotification: false, applyAutoShare: true, useBusyIndicator: false);
        }
        catch (OperationCanceledException) when (_lifetimeCancellation.IsCancellationRequested)
        {
        }
        finally
        {
            if (gateEntered)
            {
                _usbTrayRefreshGate.Release();
            }
        }
    }

    public async Task HandleUsbClientDisconnectedAsync(string busId)
    {
        var normalizedBusId = (busId ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalizedBusId)
            || !HostUsbSharingEnabled
            || !UsbAutoDetachOnClientDisconnect)
        {
            return;
        }

        var gateEntered = false;
        try
        {
            await _usbTrayRefreshGate.WaitAsync(_lifetimeCancellation.Token);
            gateEntered = true;

            using var cts = CancellationTokenSource.CreateLinkedTokenSource(_lifetimeCancellation.Token);
            cts.CancelAfter(TimeSpan.FromSeconds(12));

            await _usbIpService.DetachAsync(normalizedBusId, cts.Token);
            Log.Information("USB detach executed after client disconnect. BusId={BusId}", normalizedBusId);
        }
        catch (OperationCanceledException) when (_lifetimeCancellation.IsCancellationRequested)
        {
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Automatic USB detach after client disconnect failed. BusId={BusId}", normalizedBusId);
        }
        finally
        {
            if (gateEntered)
            {
                _usbTrayRefreshGate.Release();
            }
        }
    }

    public async Task ShareSelectedUsbFromTrayAsync()
    {
        if (!HostUsbSharingEnabled)
        {
            AddNotification("USB Share ist global im Host deaktiviert.", "Info");
            return;
        }

        if (SelectedUsbDevice is null || string.IsNullOrWhiteSpace(SelectedUsbDevice.BusId))
        {
            AddNotification("Kein USB-Gerät für Share ausgewählt.", "Warning");
            return;
        }

        await BindSelectedUsbDeviceAsync();
    }

    public async Task UnshareSelectedUsbFromTrayAsync()
    {
        if (!HostUsbSharingEnabled)
        {
            AddNotification("USB Share ist global im Host deaktiviert.", "Info");
            return;
        }

        if (SelectedUsbDevice is null || string.IsNullOrWhiteSpace(SelectedUsbDevice.BusId))
        {
            AddNotification("Kein USB-Gerät für Unshare ausgewählt.", "Warning");
            return;
        }

        await UnbindSelectedUsbDeviceAsync();
    }

    public async Task RefreshTrayDataAsync()
    {
        if (IsBusy)
        {
            return;
        }

        try
        {
            var token = _lifetimeCancellation.Token;
            var runtimeVms = await _hyperVService.GetVmsAsync(token);
            UpdateVmRuntimeStates(runtimeVms);

            var switches = await _hyperVService.GetVmSwitchesAsync(token);
            AvailableSwitches.Clear();
            foreach (var vmSwitch in switches.OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase))
            {
                AvailableSwitches.Add(vmSwitch);
            }

            AreSwitchesLoaded = true;
            SyncSelectedSwitchWithCurrentVm(showNotificationOnMissingSwitch: false);
            NotifyTrayStateChanged();
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Tray runtime refresh failed.");
        }
    }

    public Task ReloadTrayDataAsync()
    {
        return ReloadConfigAsync();
    }

    public async Task StartVmFromTrayAsync(string vmName)
    {
        await ExecuteBusyActionAsync($"VM '{vmName}' wird gestartet...", async token =>
        {
            await _hyperVService.StartVmAsync(vmName, token);
            AddNotification($"VM '{vmName}' gestartet.", "Success");

            if (UiOpenConsoleAfterVmStart)
            {
                await _hyperVService.OpenVmConnectAsync(vmName, VmConnectComputerName, ShouldOpenConsoleWithSessionEdit(vmName), token);
                AddNotification($"Konsole für '{vmName}' geöffnet.", "Info");
            }
        });
        await RefreshVmStatusByNameAsync(vmName);
    }

    public async Task StopVmFromTrayAsync(string vmName)
    {
        await ExecuteBusyActionAsync($"VM '{vmName}' wird gestoppt...", async token =>
        {
            await _hyperVService.StopVmGracefulAsync(vmName, token);
            AddNotification($"VM '{vmName}' graceful gestoppt.", "Success");
        });
        await RefreshVmStatusByNameAsync(vmName);
    }

    public async Task OpenConsoleFromTrayAsync(string vmName)
    {
        await ExecuteBusyActionAsync($"Konsole für '{vmName}' wird geöffnet...", async token =>
        {
            await _hyperVService.OpenVmConnectAsync(vmName, VmConnectComputerName, ShouldOpenConsoleWithSessionEdit(vmName), token);
            AddNotification($"Konsole für '{vmName}' geöffnet.", "Info");
        });
    }

    public async Task CreateSnapshotFromTrayAsync(string vmName)
    {
        var checkpointName = $"checkpoint-{DateTime.Now:yyyyMMdd-HHmmss}";

        await ExecuteBusyActionAsync($"Checkpoint für '{vmName}' wird erstellt...", async token =>
        {
            await _hyperVService.CreateCheckpointAsync(vmName, checkpointName, null, token);
            AddNotification($"Checkpoint '{checkpointName}' für '{vmName}' erstellt.", "Success");
        });
    }

    public async Task<IReadOnlyList<HyperVVmNetworkAdapterInfo>> GetVmNetworkAdaptersForTrayAsync(string vmName)
    {
        if (string.IsNullOrWhiteSpace(vmName))
        {
            return [];
        }

        try
        {
            var adapters = await _hyperVService.GetVmNetworkAdaptersAsync(vmName, CancellationToken.None);
            return adapters
                .Select(adapter => new HyperVVmNetworkAdapterInfo
                {
                    Name = adapter.Name,
                    SwitchName = adapter.SwitchName,
                    MacAddress = adapter.MacAddress
                })
                .OrderBy(adapter => adapter.DisplayName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Konnte Netzwerkadapter für Tray-Control-Center nicht laden: {VmName}", vmName);
            return [];
        }
    }

    public Task ConnectVmSwitchFromTrayAsync(string vmName, string switchName)
    {
        return ConnectVmSwitchFromTrayAsync(vmName, switchName, null);
    }

    public async Task ConnectVmSwitchFromTrayAsync(string vmName, string switchName, string? adapterName)
    {
        await ExecuteBusyActionAsync($"'{vmName}' wird mit '{switchName}' verbunden...", async token =>
        {
            string? trayAdapterName = null;
            var vmConfig = AvailableVms.FirstOrDefault(vm => string.Equals(vm.Name, vmName, StringComparison.OrdinalIgnoreCase));
            var configuredAdapterName = vmConfig?.TrayAdapterName?.Trim();
            var requestedAdapterName = adapterName?.Trim();

            IReadOnlyList<HyperVVmNetworkAdapterInfo>? adapters = null;

            if (!string.IsNullOrWhiteSpace(requestedAdapterName)
                || !string.IsNullOrWhiteSpace(configuredAdapterName))
            {
                adapters = await _hyperVService.GetVmNetworkAdaptersAsync(vmName, token);
            }

            if (!string.IsNullOrWhiteSpace(requestedAdapterName))
            {
                var requestedExists = adapters?.Any(adapter =>
                    string.Equals(adapter.Name, requestedAdapterName, StringComparison.OrdinalIgnoreCase)) == true;

                if (requestedExists)
                {
                    trayAdapterName = requestedAdapterName;
                }
                else
                {
                    AddNotification($"Gewählter Adapter '{requestedAdapterName}' für '{vmName}' wurde nicht gefunden. Fallback auf Standard-Verhalten.", "Warning");
                }
            }

            if (string.IsNullOrWhiteSpace(trayAdapterName)
                && !string.IsNullOrWhiteSpace(configuredAdapterName))
            {
                var exists = adapters?.Any(adapter => string.Equals(adapter.Name, configuredAdapterName, StringComparison.OrdinalIgnoreCase)) == true;

                if (exists)
                {
                    trayAdapterName = configuredAdapterName;
                }
                else
                {
                    AddNotification($"Konfigurierter Tray-Adapter '{configuredAdapterName}' für '{vmName}' wurde nicht gefunden. Fallback auf Standard-Verhalten.", "Warning");
                }
            }

            await _hyperVService.ConnectVmNetworkAdapterAsync(vmName, switchName, trayAdapterName, token);

            if (string.IsNullOrWhiteSpace(trayAdapterName))
            {
                AddNotification($"'{vmName}' mit '{switchName}' verbunden.", "Success");
            }
            else
            {
                AddNotification($"'{vmName}' Adapter '{trayAdapterName}' mit '{switchName}' verbunden.", "Success");
            }

            if (ShouldAutoRestartHnsAfterConnect(switchName))
            {
                var hnsResult = await _hnsService.RestartHnsElevatedAsync(token);
                AddNotification(
                    hnsResult.Success ? hnsResult.Message : $"HNS Neustart fehlgeschlagen: {hnsResult.Message}",
                    hnsResult.Success ? "Success" : "Error");
            }
        });

        await RefreshVmStatusByNameAsync(vmName);
    }

    public async Task DisconnectVmSwitchFromTrayAsync(string vmName)
    {
        await ExecuteBusyActionAsync($"'{vmName}' wird vom Switch getrennt...", async token =>
        {
            string? trayAdapterName = null;
            var vmConfig = AvailableVms.FirstOrDefault(vm => string.Equals(vm.Name, vmName, StringComparison.OrdinalIgnoreCase));
            var configuredAdapterName = vmConfig?.TrayAdapterName?.Trim();

            if (!string.IsNullOrWhiteSpace(configuredAdapterName))
            {
                var adapters = await _hyperVService.GetVmNetworkAdaptersAsync(vmName, token);
                var exists = adapters.Any(adapter => string.Equals(adapter.Name, configuredAdapterName, StringComparison.OrdinalIgnoreCase));

                if (exists)
                {
                    trayAdapterName = configuredAdapterName;
                }
                else
                {
                    AddNotification($"Konfigurierter Tray-Adapter '{configuredAdapterName}' für '{vmName}' wurde nicht gefunden. Fallback auf Standard-Verhalten.", "Warning");
                }
            }

            await _hyperVService.DisconnectVmNetworkAdapterAsync(vmName, trayAdapterName, token);

            if (string.IsNullOrWhiteSpace(trayAdapterName))
            {
                AddNotification($"'{vmName}' vom Switch getrennt.", "Success");
            }
            else
            {
                AddNotification($"'{vmName}' Adapter '{trayAdapterName}' vom Switch getrennt.", "Success");
            }
        });

        await RefreshVmStatusByNameAsync(vmName);
    }

    public async Task<IReadOnlyList<HostNetworkAdapterInfo>> GetHostNetworkAdaptersWithUplinkAsync()
    {
        if (IsBusy)
        {
            AddNotification("Bitte warten, ein anderer Vorgang läuft noch.", "Info");
            return [];
        }

        var adapters = Array.Empty<HostNetworkAdapterInfo>();

        await ExecuteBusyActionAsync("Host-Netzwerkdaten werden geladen...", async token =>
        {
            var result = await _hyperVService.GetHostNetworkAdaptersWithUplinkAsync(token);
            adapters = result
                .OrderByDescending(item => !string.IsNullOrWhiteSpace(item.Gateway))
                .ThenBy(item => item.AdapterName, StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }, showNotificationOnErrorOnly: true);

        if (adapters.Length == 0)
        {
            AddNotification("Keine Host-Netzwerkkarten gefunden.", "Warning");
            return adapters;
        }

        AddNotification($"{adapters.Length} Host-Netzwerkkarte(n) geladen.", "Info");
        return adapters;
    }

    public async Task<bool> SetHostNetworkProfileCategoryAsync(string? adapterName, string category)
    {
        if (string.IsNullOrWhiteSpace(category))
        {
            AddNotification("Netzprofil konnte nicht gesetzt werden: Kategorie fehlt.", "Warning");
            return false;
        }

        var normalizedCategory = category.Trim() switch
        {
            "Private" => "Private",
            "Public" => "Public",
            _ => string.Empty
        };

        if (string.IsNullOrWhiteSpace(normalizedCategory))
        {
            AddNotification("Netzprofil konnte nicht gesetzt werden: nur Private oder Public sind erlaubt.", "Warning");
            return false;
        }

        var updated = false;
        await ExecuteBusyActionAsync("Host-Netzprofil wird geändert...", async token =>
        {
            await _hyperVService.SetHostNetworkProfileCategoryAsync(adapterName ?? string.Empty, normalizedCategory, token);
            await RefreshHostNetworkProfileAsync();

            var targetText = string.IsNullOrWhiteSpace(adapterName)
                ? "aktive Host-Netzprofile"
                : $"Netzprofil von '{adapterName}'";

            var categoryText = normalizedCategory == "Private" ? "Privat" : "Öffentlich";
            AddNotification($"{targetText} auf '{categoryText}' gesetzt.", "Success");
            updated = true;
        }, showNotificationOnErrorOnly: true);

        return updated;
    }

    private async Task RefreshVmStatusByNameAsync(string vmName)
    {
        var currentSelectedVm = SelectedVm;

        SetSelectedVmInternal(AvailableVms.FirstOrDefault(vm =>
                                string.Equals(vm.Name, vmName, StringComparison.OrdinalIgnoreCase))
                            ?? currentSelectedVm);

        await RefreshVmStatusAsync();
    }

    private void SetSelectedVmInternal(VmDefinition? vm)
    {
        _selectedVmChangeSuppressionDepth++;
        try
        {
            SelectedVm = vm;
        }
        finally
        {
            _selectedVmChangeSuppressionDepth--;
        }
    }

    private async Task ReloadConfigAsync()
    {
        await ExecuteBusyActionAsync("Konfiguration wird neu geladen...", _ =>
        {
            var configResult = _configService.LoadOrCreate(_configPath);
            var config = configResult.Config;
            var previousSelectionName = SelectedVm?.Name;

            _configChangeSuppressionDepth++;
            try
            {
                WindowTitle = "HyperTool";
                ConfigurationNotice = configResult.Notice;
                HnsEnabled = config.Hns.Enabled;
                HnsAutoRestartAfterDefaultSwitch = config.Hns.AutoRestartAfterDefaultSwitch;
                HnsAutoRestartAfterAnyConnect = config.Hns.AutoRestartAfterAnyConnect;
                DefaultVmName = config.DefaultVmName;
                LastSelectedVmName = config.LastSelectedVmName;
                VmConnectComputerName = NormalizeVmConnectComputerName(config.VmConnectComputerName);
                UiEnableTrayIcon = true;
                UiEnableTrayMenu = config.Ui.EnableTrayMenu;
                UiStartMinimized = config.Ui.StartMinimized;
                UiStartWithWindows = config.Ui.StartWithWindows;
                UiOpenConsoleAfterVmStart = config.Ui.OpenConsoleAfterVmStart;
                UiRestoreNumLockAfterVmStart = config.Ui.RestoreNumLockAfterVmStart;
                _uiNumLockWatcherIntervalSeconds = Math.Clamp(config.Ui.NumLockWatcherIntervalSeconds, 5, 600);
                UiOpenVmConnectWithSessionEdit = config.Ui.OpenVmConnectWithSessionEdit;
                UiTheme = NormalizeUiTheme(config.Ui.Theme);
                ApplyConfiguredVmDefinitions(config.Vms);
                _trayVmNames = NormalizeTrayVmNames(config.Ui.TrayVmNames);
                UpdateCheckOnStartup = config.Update.CheckOnStartup;
                GithubOwner = config.Update.GitHubOwner;
                GithubRepo = config.Update.GitHubRepo;
                DefaultVmImportDestinationPath = config.DefaultVmImportDestinationPath?.Trim() ?? string.Empty;
                _usbAutoShareDeviceKeys.Clear();
                foreach (var key in config.Usb.AutoShareDeviceKeys)
                {
                    if (!string.IsNullOrWhiteSpace(key))
                    {
                        _usbAutoShareDeviceKeys.Add(key.Trim());
                    }
                }
                ApplyConfiguredSharedFolders(config.SharedFolders.HostDefinitions);

                _suppressUsbAutoShareToggleHandling = true;
                try
                {
                    SelectedUsbDeviceAutoShareEnabled = SelectedUsbDevice is not null
                        && _usbAutoShareDeviceKeys.Contains(BuildUsbAutoShareKey(SelectedUsbDevice));
                }
                finally
                {
                    _suppressUsbAutoShareToggleHandling = false;
                }
            }
            finally
            {
                _configChangeSuppressionDepth--;
            }

            if (string.IsNullOrWhiteSpace(LastSelectedVmName) && !string.IsNullOrWhiteSpace(previousSelectionName))
            {
                LastSelectedVmName = previousSelectionName;
            }

            AddNotification("Konfiguration neu geladen.", "Info");
            return Task.CompletedTask;
        });

        await LoadVmsFromHyperVAsync();
        await RefreshSwitchesAsync();
        NotifyTrayStateChanged();
        MarkConfigClean();
    }

    private void ApplyConfiguredVmDefinitions(IEnumerable<VmDefinition>? configuredVms)
    {
        _configuredVmDefinitions.Clear();

        if (configuredVms is null)
        {
            return;
        }

        foreach (var vm in configuredVms)
        {
            if (vm is null || string.IsNullOrWhiteSpace(vm.Name))
            {
                continue;
            }

            var vmName = vm.Name.Trim();
            _configuredVmDefinitions[vmName] = new VmDefinition
            {
                Name = vmName,
                Label = string.IsNullOrWhiteSpace(vm.Label) ? vmName : vm.Label.Trim(),
                TrayAdapterName = vm.TrayAdapterName?.Trim() ?? string.Empty,
                OpenConsoleWithSessionEdit = vm.OpenConsoleWithSessionEdit
            };
        }
    }

    private void ApplyConfiguredSharedFolders(IEnumerable<HostSharedFolderDefinition>? sharedFolders)
    {
        HostSharedFolders.Clear();

        if (sharedFolders is null)
        {
            return;
        }

        foreach (var folder in sharedFolders)
        {
            if (folder is null)
            {
                continue;
            }

            var normalized = NormalizeSharedFolderDefinition(folder);
            if (string.IsNullOrWhiteSpace(normalized.LocalPath)
                || string.IsNullOrWhiteSpace(normalized.ShareName))
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(normalized.Id))
            {
                normalized.Id = Guid.NewGuid().ToString("N");
            }

            HostSharedFolders.Add(normalized);
        }
    }

    private static HostSharedFolderDefinition NormalizeSharedFolderDefinition(HostSharedFolderDefinition definition)
    {
        return new HostSharedFolderDefinition
        {
            Id = (definition.Id ?? string.Empty).Trim(),
            Label = (definition.Label ?? string.Empty).Trim(),
            LocalPath = (definition.LocalPath ?? string.Empty).Trim(),
            ShareName = (definition.ShareName ?? string.Empty).Trim(),
            Enabled = definition.Enabled,
            ReadOnly = definition.ReadOnly
        };
    }

    private static HostSharedFolderDefinition CloneSharedFolderDefinition(HostSharedFolderDefinition definition)
    {
        return new HostSharedFolderDefinition
        {
            Id = definition.Id,
            Label = definition.Label,
            LocalPath = definition.LocalPath,
            ShareName = definition.ShareName,
            Enabled = definition.Enabled,
            ReadOnly = definition.ReadOnly
        };
    }

    private static List<string> NormalizeTrayVmNames(IEnumerable<string>? vmNames)
    {
        if (vmNames is null)
        {
            return [];
        }

        return vmNames
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Select(name => name.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public async Task UnshareAllSharedUsbOnShutdownAsync()
    {
        if (IsBusy)
        {
            try
            {
                await _usbIpService.ShutdownElevatedSessionAsync(CancellationToken.None);
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Elevated USB session could not be closed during shutdown while busy.");
            }

            return;
        }

        try
        {
            await UnshareAllSharedUsbAsync(
                timeout: TimeSpan.FromSeconds(20),
                successMessage: "Beim Beenden wurden {0} USB-Freigabe(n) entfernt.",
                failedMessage: "Beim Beenden konnten {0} USB-Freigabe(n) nicht entfernt werden.",
                logContext: "shutdown");
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "USB-Freigaben konnten beim Beenden nicht vollständig entfernt werden.");
        }
        finally
        {
            try
            {
                await _usbIpService.ShutdownElevatedSessionAsync(CancellationToken.None);
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "Elevated USB session could not be closed during shutdown cleanup.");
            }
        }
    }

    private async Task UnshareAllSharedUsbAsync(TimeSpan timeout, string successMessage, string failedMessage, string logContext)
    {
        using var cts = new CancellationTokenSource(timeout);
        IReadOnlyList<UsbIpDeviceInfo> devices;
        try
        {
            devices = await _usbIpService.GetDevicesAsync(cts.Token);
        }
        catch (Exception ex) when (IsUsbRuntimeMissing(ex.Message))
        {
            Log.Information("USB-Unshare übersprungen ({Context}): usbipd-win nicht verfügbar.", logContext);
            UsbRuntimeAvailable = false;
            UsbRuntimeHintText = "USB-Funktion deaktiviert: usbipd-win ist nicht installiert. Quelle: https://github.com/dorssel/usbipd-win";
            return;
        }

        var sharedDevices = devices.Where(device => device.IsShared).ToList();

        if (sharedDevices.Count == 0)
        {
            return;
        }

        var released = 0;
        var failed = 0;

        foreach (var device in sharedDevices)
        {
            try
            {
                if (device.IsAttached && !string.IsNullOrWhiteSpace(device.BusId))
                {
                    await _usbIpService.DetachAsync(device.BusId, cts.Token);
                }

                if (!string.IsNullOrWhiteSpace(device.BusId))
                {
                    await _usbIpService.UnbindAsync(device.BusId, cts.Token);
                }
                else if (!string.IsNullOrWhiteSpace(device.PersistedGuid))
                {
                    await _usbIpService.UnbindByPersistedGuidAsync(device.PersistedGuid, cts.Token);
                }
                else
                {
                    failed++;
                    continue;
                }

                released++;
            }
            catch (Exception ex)
            {
                failed++;
                Log.Warning(ex,
                    "Freigeben von USB-Gerät fehlgeschlagen ({Context}) (BusId={BusId}, Guid={Guid}).",
                    logContext,
                    device.BusId,
                    device.PersistedGuid);
            }
        }

        if (released > 0)
        {
            AddNotification(string.Format(successMessage, released), "Info");
        }

        if (failed > 0)
        {
            AddNotification(string.Format(failedMessage, failed), "Warning");
        }
    }

    private static string NormalizeVmConnectComputerName(string? computerName)
    {
        var normalized = computerName?.Trim() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(normalized)
            || string.Equals(normalized, "localhost", StringComparison.OrdinalIgnoreCase))
        {
            return Environment.MachineName;
        }

        return normalized;
    }

    private bool ShouldOpenConsoleWithSessionEdit(string vmName)
    {
        var vm = AvailableVms.FirstOrDefault(item => string.Equals(item.Name, vmName, StringComparison.OrdinalIgnoreCase))
                 ?? _configuredVmDefinitions.Values.FirstOrDefault(item => string.Equals(item.Name, vmName, StringComparison.OrdinalIgnoreCase));

        if (vm is not null)
        {
            return vm.OpenConsoleWithSessionEdit;
        }

        return UiOpenVmConnectWithSessionEdit;
    }

    public bool TryPromptSaveConfigOnClose()
    {
        return TryPromptSaveConfigChanges();
    }

    private bool ShouldPromptSaveWhenLeavingConfig(int newMenuIndex)
    {
        return _lastSelectedMenuIndex == ConfigMenuIndex
               && newMenuIndex != ConfigMenuIndex
               && HasPendingConfigChanges
               && !IsBusy;
    }

    private bool TryPromptSaveConfigChanges()
    {
        if (!CanPromptSaveOnClose)
        {
            return true;
        }

        var result = _uiInteropService.ShowUnsavedConfigPrompt();

        if (result == UnsavedConfigPromptResult.Cancel)
        {
            return false;
        }

        if (result == UnsavedConfigPromptResult.No)
        {
            ReloadConfigSnapshotWithoutRuntimeRefresh();
            return true;
        }

        SaveConfigCommand.ExecuteAsync(null).GetAwaiter().GetResult();
        return !HasPendingConfigChanges;
    }

    private void MarkConfigDirty()
    {
        if (_configChangeSuppressionDepth > 0)
        {
            return;
        }

        HasPendingConfigChanges = true;
    }

    private void ReloadConfigSnapshotWithoutRuntimeRefresh()
    {
        var configResult = _configService.LoadOrCreate(_configPath);
        var config = configResult.Config;
        var previousSelectionName = SelectedVm?.Name;

        _configChangeSuppressionDepth++;
        try
        {
            WindowTitle = "HyperTool";
            ConfigurationNotice = configResult.Notice;
            HnsEnabled = config.Hns.Enabled;
            HnsAutoRestartAfterDefaultSwitch = config.Hns.AutoRestartAfterDefaultSwitch;
            HnsAutoRestartAfterAnyConnect = config.Hns.AutoRestartAfterAnyConnect;
            DefaultVmName = config.DefaultVmName;
            LastSelectedVmName = config.LastSelectedVmName;
            VmConnectComputerName = NormalizeVmConnectComputerName(config.VmConnectComputerName);
            UiEnableTrayIcon = true;
            UiEnableTrayMenu = config.Ui.EnableTrayMenu;
            UiStartMinimized = config.Ui.StartMinimized;
            UiStartWithWindows = config.Ui.StartWithWindows;
            UiOpenConsoleAfterVmStart = config.Ui.OpenConsoleAfterVmStart;
            UiOpenVmConnectWithSessionEdit = config.Ui.OpenVmConnectWithSessionEdit;
            UiTheme = NormalizeUiTheme(config.Ui.Theme);
            ApplyConfiguredVmDefinitions(config.Vms);
            _trayVmNames = NormalizeTrayVmNames(config.Ui.TrayVmNames);
            UpdateCheckOnStartup = config.Update.CheckOnStartup;
            GithubOwner = config.Update.GitHubOwner;
            GithubRepo = config.Update.GitHubRepo;
            HostUsbSharingEnabled = config.Usb.Enabled;
            UsbAutoDetachOnClientDisconnect = config.Usb.AutoDetachOnClientDisconnect;
            UsbUnshareOnExit = config.Usb.UnshareOnExit;
            HostSharedFoldersEnabled = config.SharedFolders.Enabled;

            _usbAutoShareDeviceKeys.Clear();
            foreach (var key in config.Usb.AutoShareDeviceKeys)
            {
                if (!string.IsNullOrWhiteSpace(key))
                {
                    _usbAutoShareDeviceKeys.Add(key.Trim());
                }
            }

            ApplyConfiguredSharedFolders(config.SharedFolders.HostDefinitions);

            _suppressUsbAutoShareToggleHandling = true;
            try
            {
                SelectedUsbDeviceAutoShareEnabled = SelectedUsbDevice is not null
                    && _usbAutoShareDeviceKeys.Contains(BuildUsbAutoShareKey(SelectedUsbDevice));
            }
            finally
            {
                _suppressUsbAutoShareToggleHandling = false;
            }
        }
        finally
        {
            _configChangeSuppressionDepth--;
        }

        if (string.IsNullOrWhiteSpace(LastSelectedVmName) && !string.IsNullOrWhiteSpace(previousSelectionName))
        {
            LastSelectedVmName = previousSelectionName;
        }

        MarkConfigClean();
        NotifyTrayStateChanged();
        AddNotification("Konfiguration neu geladen.", "Info");
    }

    private void MarkConfigClean()
    {
        HasPendingConfigChanges = false;
    }

    private static string NormalizeUiTheme(string? theme)
    {
        if (string.Equals(theme, "Light", StringComparison.OrdinalIgnoreCase)
            || string.Equals(theme, "Bright", StringComparison.OrdinalIgnoreCase))
        {
            return "Light";
        }

        return "Dark";
    }

    private async Task PersistSelectedVmAsync(string vmName)
    {
        try
        {
            var configResult = _configService.LoadOrCreate(_configPath);
            var config = configResult.Config;
            config.LastSelectedVmName = vmName;

            _ = _configService.TrySave(_configPath, config, out _);
            await Task.CompletedTask;
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Selected VM could not be persisted.");
        }
    }

    private async Task StartVmByNameAsync(string? vmName)
    {
        if (string.IsNullOrWhiteSpace(vmName))
        {
            return;
        }

        await ExecuteBusyActionAsync($"VM '{vmName}' wird gestartet...", async token =>
        {
            await _hyperVService.StartVmAsync(vmName, token);
            AddNotification($"VM '{vmName}' gestartet.", "Success");

            if (UiOpenConsoleAfterVmStart)
            {
                await _hyperVService.OpenVmConnectAsync(vmName, VmConnectComputerName, ShouldOpenConsoleWithSessionEdit(vmName), token);
                AddNotification($"Konsole für '{vmName}' geöffnet.", "Info");
            }
        });

        await RefreshVmStatusByNameAsync(vmName);
    }

    private async Task RunNumLockWatcherAsync()
    {
        var token = _lifetimeCancellation.Token;
        while (!token.IsCancellationRequested)
        {
            try
            {
                if (UiRestoreNumLockAfterVmStart && !IsNumLockEnabled())
                {
                    ToggleNumLockKey();
                }

                var intervalSeconds = Math.Clamp(_uiNumLockWatcherIntervalSeconds, 5, 600);
                await Task.Delay(TimeSpan.FromSeconds(intervalSeconds), token);
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested)
            {
                return;
            }
            catch (Exception ex)
            {
                Log.Debug(ex, "NumLock watcher iteration failed.");

                try
                {
                    await Task.Delay(TimeSpan.FromSeconds(5), token);
                }
                catch (OperationCanceledException) when (token.IsCancellationRequested)
                {
                    return;
                }
            }
        }
    }

    private static bool IsNumLockEnabled()
    {
        return (GetKeyState(VkNumLock) & 0x1) != 0;
    }

    private static void ToggleNumLockKey()
    {
        keybd_event((byte)VkNumLock, 0x45, KeyeventfExtendedKey, UIntPtr.Zero);
        keybd_event((byte)VkNumLock, 0x45, KeyeventfExtendedKey | KeyeventfKeyUp, UIntPtr.Zero);
    }

    private async Task StopVmByNameAsync(string? vmName)
    {
        if (string.IsNullOrWhiteSpace(vmName))
        {
            return;
        }

        await ExecuteBusyActionAsync($"VM '{vmName}' wird gestoppt...", async token =>
        {
            await _hyperVService.StopVmGracefulAsync(vmName, token);
            AddNotification($"VM '{vmName}' gestoppt.", "Success");
        });

        await RefreshVmStatusByNameAsync(vmName);
    }

    private async Task TurnOffVmByNameAsync(string? vmName)
    {
        if (string.IsNullOrWhiteSpace(vmName))
        {
            return;
        }

        await ExecuteBusyActionAsync($"VM '{vmName}' wird hart ausgeschaltet...", async token =>
        {
            await _hyperVService.TurnOffVmAsync(vmName, token);
            AddNotification($"VM '{vmName}' hart ausgeschaltet.", "Warning");
        });

        await RefreshVmStatusByNameAsync(vmName);
    }

    private async Task RestartVmByNameAsync(string? vmName)
    {
        if (string.IsNullOrWhiteSpace(vmName))
        {
            return;
        }

        await ExecuteBusyActionAsync($"VM '{vmName}' wird neu gestartet...", async token =>
        {
            await _hyperVService.RestartVmAsync(vmName, token);
            AddNotification($"VM '{vmName}' neu gestartet.", "Success");
        });

        await RefreshVmStatusByNameAsync(vmName);
    }

    private async Task OpenConsoleByNameAsync(string? vmName)
    {
        if (string.IsNullOrWhiteSpace(vmName))
        {
            return;
        }

        await ExecuteBusyActionAsync($"Konsole für '{vmName}' wird geöffnet...", async token =>
        {
            await _hyperVService.OpenVmConnectAsync(vmName, VmConnectComputerName, ShouldOpenConsoleWithSessionEdit(vmName), token);
            AddNotification($"Konsole für '{vmName}' geöffnet.", "Info");
        });
    }

    private async Task CreateSnapshotByNameAsync(string? vmName)
    {
        if (string.IsNullOrWhiteSpace(vmName))
        {
            return;
        }

        var checkpointName = $"checkpoint-{DateTime.Now:yyyyMMdd-HHmmss}";
        await ExecuteBusyActionAsync($"Snapshot für '{vmName}' wird erstellt...", async token =>
        {
            await _hyperVService.CreateCheckpointAsync(vmName, checkpointName, null, token);
            AddNotification($"Checkpoint '{checkpointName}' für '{vmName}' erstellt.", "Success");
        });
    }

    private void ClearNotifications()
    {
        Notifications.Clear();
    }

    private void CopyNotificationsToClipboard()
    {
        var lines = Notifications
            .Select(entry => $"[{entry.Timestamp:yyyy-MM-dd HH:mm:ss}] [{entry.Level}] {entry.Message}")
            .ToArray();

        var text = lines.Length == 0 ? "Keine Notifications vorhanden." : string.Join(Environment.NewLine, lines);
        _uiInteropService.SetClipboardText(text);
        AddNotification("Notifications in Zwischenablage kopiert.", "Info");
    }

    private async Task CheckForUpdatesAsync()
    {
        await ExecuteBusyActionAsync("Prüfe GitHub-Version...", async token =>
        {
            var result = await _updateService.CheckForUpdateAsync(
                GithubOwner,
                GithubRepo,
                AppVersion,
                token,
                "HyperTool-Setup");

            UpdateStatus = result.Message;
            ReleaseUrl = result.ReleaseUrl ?? string.Empty;
            InstallerDownloadUrl = result.InstallerDownloadUrl ?? string.Empty;
            InstallerFileName = result.InstallerFileName ?? string.Empty;
            UpdateInstallAvailable = result.HasUpdate && !string.IsNullOrWhiteSpace(InstallerDownloadUrl);

            if (!result.Success)
            {
                AddNotification(result.Message, "Warning");
                return;
            }

            if (result.HasUpdate && !UpdateInstallAvailable)
            {
                AddNotification("Update gefunden, aber kein Installer-Asset im Release erkannt. Bitte Release-Seite öffnen.", "Warning");
            }

            AddNotification(result.Message, result.HasUpdate ? "Success" : "Info");
        });
    }

    private async Task InstallUpdateAsync()
    {
        if (string.IsNullOrWhiteSpace(InstallerDownloadUrl))
        {
            AddNotification("Kein Installer-Download verfügbar.", "Warning");
            return;
        }

        if (HasPendingConfigChanges)
        {
            AddNotification("Ungespeicherte Einstellungen werden vor dem Update gespeichert...", "Info");
            await SaveConfigAsync();

            if (HasPendingConfigChanges)
            {
                AddNotification("Update abgebrochen: Konfiguration konnte nicht sicher gespeichert werden.", "Error");
                return;
            }
        }

        await ExecuteBusyActionAsync("Update wird heruntergeladen...", async token =>
        {
            var targetDirectory = Path.Combine(Path.GetTempPath(), "HyperTool", "updates");
            Directory.CreateDirectory(targetDirectory);

            var fileName = ResolveInstallerFileName(InstallerDownloadUrl, InstallerFileName);
            var installerPath = Path.Combine(targetDirectory, fileName);

            using var response = await UpdateDownloadClient.GetAsync(InstallerDownloadUrl, token);
            response.EnsureSuccessStatusCode();

            await using (var fileStream = new FileStream(installerPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                await response.Content.CopyToAsync(fileStream, token);
            }

            AddNotification($"Installer heruntergeladen: {installerPath}", "Success");

            Process.Start(new ProcessStartInfo
            {
                FileName = installerPath,
                UseShellExecute = true
            });

            AddNotification("Installer gestartet. HyperTool wird beendet.", "Info");
            _uiInteropService.ShutdownApplication();
        });
    }

    private static string ResolveInstallerFileName(string downloadUrl, string? fileName)
    {
        if (!string.IsNullOrWhiteSpace(fileName))
        {
            return fileName.Trim();
        }

        if (Uri.TryCreate(downloadUrl, UriKind.Absolute, out var uri))
        {
            var name = Path.GetFileName(uri.LocalPath);
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name;
            }
        }

        return "HyperTool-Setup.exe";
    }

    private void OpenReleasePage()
    {
        if (string.IsNullOrWhiteSpace(ReleaseUrl))
        {
            AddNotification("Keine Release-URL verfügbar.", "Warning");
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = ReleaseUrl,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            AddNotification($"Release-Seite konnte nicht geöffnet werden: {ex.Message}", "Error");
        }
    }

    private void ToggleLog()
    {
        IsLogExpanded = !IsLogExpanded;
    }

    private void OpenLogFile()
    {
        var logDirectoryPath = ResolveLogDirectoryPath();

        try
        {
            Directory.CreateDirectory(logDirectoryPath);

            Process.Start(new ProcessStartInfo
            {
                FileName = logDirectoryPath,
                UseShellExecute = true
            });

            AddNotification($"Log-Ordner geöffnet: {logDirectoryPath}", "Info");
        }
        catch (Exception ex)
        {
            AddNotification($"Log-Ordner konnte nicht geöffnet werden: {ex.Message}", "Error");
        }
    }

    private void SelectVmFromChip(VmDefinition? vm)
    {
        if (vm is null)
        {
            return;
        }

        var selected = AvailableVms.FirstOrDefault(item => string.Equals(item.Name, vm.Name, StringComparison.OrdinalIgnoreCase));
        if (selected is not null)
        {
            SelectedVm = selected;
        }
    }

    private void OnNotificationsChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        OnPropertyChanged(nameof(LatestNotification));
        OnPropertyChanged(nameof(LastNotificationText));
    }

    private void NotifyTrayStateChanged()
    {
        TrayStateChanged?.Invoke(this, EventArgs.Empty);
    }

    public void PublishNotification(string message, string level = "Info")
    {
        AddNotification(message, level);
    }

    private static string ResolveAppVersion()
    {
        var informationalVersion = Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;

        if (!string.IsNullOrWhiteSpace(informationalVersion))
        {
            return FormatVersionForDisplay(informationalVersion);
        }

        var assemblyVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "0.0.0";
        return FormatVersionForDisplay(assemblyVersion);
    }

    private static string FormatVersionForDisplay(string version)
    {
        if (string.IsNullOrWhiteSpace(version))
        {
            return "0.0.0";
        }

        var raw = version.Trim();
        var plusIndex = raw.IndexOf('+');
        if (plusIndex >= 0)
        {
            raw = raw[..plusIndex];
        }

        raw = raw.TrimStart('v', 'V');

        if (Version.TryParse(raw, out var parsed))
        {
            if (parsed.Build > 0)
            {
                if (parsed.Revision > 0)
                {
                    return $"{parsed.Major}.{parsed.Minor}.{parsed.Build}.{parsed.Revision}";
                }

                return $"{parsed.Major}.{parsed.Minor}.{parsed.Build}";
            }

            return $"{parsed.Major}.{parsed.Minor}";
        }

        return raw;
    }

    private static string FormatByteSize(long bytes)
    {
        if (bytes <= 0)
        {
            return "0 B";
        }

        string[] units = ["B", "KB", "MB", "GB", "TB"];
        double value = bytes;
        var unitIndex = 0;
        while (value >= 1024 && unitIndex < units.Length - 1)
        {
            value /= 1024;
            unitIndex++;
        }

        return $"{value:0.##} {units[unitIndex]}";
    }

    private static bool IsProcessElevated()
    {
        try
        {
            using var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
        catch
        {
            return false;
        }
    }

    private async Task ExecuteBusyActionAsync(string busyText, Func<CancellationToken, Task> action, bool showNotificationOnErrorOnly = false)
    {
        if (IsBusy)
        {
            return;
        }

        IsBusy = true;
        BusyText = busyText;
        BusyProgressPercent = -1;

        try
        {
            await action(_lifetimeCancellation.Token);
        }
        catch (UnauthorizedAccessException ex)
        {
            Log.Warning(ex, "Aktion fehlgeschlagen (Berechtigung): {BusyText}", busyText);
            AddNotification(ex.Message, "Warning");
            StatusText = "Keine Berechtigung";
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Aktion fehlgeschlagen: {BusyText}", busyText);
            AddNotification($"Fehler: {ex.Message}", "Error");
            StatusText = "Fehler";
        }
        finally
        {
            IsBusy = false;
            BusyText = "Bitte warten...";
            BusyProgressPercent = -1;

            if (!showNotificationOnErrorOnly && !StatusText.Equals("Fehler", StringComparison.OrdinalIgnoreCase))
            {
                StatusText = "Bereit";
            }
        }
    }

    private void AddNotification(string message, string level)
    {
        var entry = new UiNotification
        {
            Message = message,
            Level = level,
            Timestamp = DateTime.Now
        };

        Notifications.Insert(0, entry);
        while (Notifications.Count > 200)
        {
            Notifications.RemoveAt(Notifications.Count - 1);
        }

        Log.Information("UI Notification ({Level}): {Message}", level, message);
    }

    private string? PickFolderPath(string description)
    {
        return _uiInteropService.PickFolderPath(description);
    }

    private static string ResolveLogDirectoryPath()
    {
        var logDirectoryCandidates = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "HyperTool", "logs"),
            Path.Combine(AppContext.BaseDirectory, "logs"),
            Path.Combine(Path.GetTempPath(), "HyperTool", "logs")
        };

        foreach (var directory in logDirectoryCandidates)
        {
            if (Directory.Exists(directory))
            {
                return directory;
            }
        }

        return logDirectoryCandidates[0];
    }
}