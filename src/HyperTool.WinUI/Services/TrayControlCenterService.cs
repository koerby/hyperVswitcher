using HyperTool.Models;
using HyperTool.WinUI.Views;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Serilog;
using System.Runtime.InteropServices;
using Windows.Graphics;

namespace HyperTool.WinUI.Services;

internal sealed class TrayControlCenterService : ITrayControlCenterService
{
    private readonly DispatcherQueue _dispatcherQueue;
    private readonly object _syncLock = new();
    private readonly SemaphoreSlim _refreshGate = new(1, 1);
    private static readonly TimeSpan ToggleRefreshInterval = TimeSpan.FromSeconds(6);

    private TrayControlCenterWindow? _window;
    private Action? _showMainWindowAction;
    private Action? _hideMainWindowAction;
    private Func<bool>? _isMainWindowVisible;
    private Func<string>? _getUiTheme;
    private Func<IReadOnlyList<VmDefinition>>? _getVms;
    private Func<string, Task<IReadOnlyList<HyperVVmNetworkAdapterInfo>>>? _getVmAdapters;
    private Func<IReadOnlyList<HyperVSwitchInfo>>? _getSwitches;
    private Func<IReadOnlyList<UsbIpDeviceInfo>>? _getUsbDevices;
    private Func<string, Task>? _selectUsbDeviceAction;
    private Func<bool>? _isTrayMenuEnabled;
    private Func<Task>? _refreshTrayDataAction;
    private Func<string, Task>? _startVmAction;
    private Func<string, Task>? _stopVmAction;
    private Func<string, Task>? _restartVmAction;
    private Func<string, Task>? _openConsoleAction;
    private Func<string, Task>? _createSnapshotAction;
    private Func<string, string, string?, Task>? _connectVmToSwitchAction;
    private Func<UsbIpDeviceInfo?>? _getSelectedUsbDevice;
    private Func<Task>? _refreshUsbDevicesAction;
    private Func<Task>? _shareSelectedUsbAction;
    private Func<Task>? _unshareSelectedUsbAction;
    private Action? _exitAction;

    private readonly List<VmDefinition> _vms = [];
    private readonly List<HyperVVmNetworkAdapterInfo> _vmAdapters = [];
    private readonly List<HyperVSwitchInfo> _switches = [];
    private readonly List<UsbIpDeviceInfo> _usbDevices = [];
    private UsbIpDeviceInfo? _selectedUsbDevice;
    private int _selectedVmIndex = -1;
    private int _selectedUsbIndex = -1;
    private string? _selectedVmAdapterName;
    private string? _selectedSwitchName;
    private bool _isInitialized;
    private bool _isBusy;
    private bool _isBackgroundRefreshRunning;
    private DateTime _lastBackendRefreshUtc = DateTime.MinValue;
    private TrayControlCenterMode _mode = TrayControlCenterMode.Full;

    public TrayControlCenterService(DispatcherQueue dispatcherQueue)
    {
        _dispatcherQueue = dispatcherQueue;
    }

    public void Initialize(
        Action showMainWindowAction,
        Action hideMainWindowAction,
        Func<bool> isMainWindowVisible,
        Func<string> getUiTheme,
        Func<IReadOnlyList<VmDefinition>> getVms,
        Func<string, Task<IReadOnlyList<HyperVVmNetworkAdapterInfo>>> getVmAdapters,
        Func<IReadOnlyList<HyperVSwitchInfo>> getSwitches,
        Func<IReadOnlyList<UsbIpDeviceInfo>> getUsbDevices,
        Func<string, Task> selectUsbDeviceAction,
        Func<bool> isTrayMenuEnabled,
        Func<Task> refreshTrayDataAction,
        Func<string, Task> startVmAction,
        Func<string, Task> stopVmAction,
        Func<string, Task> restartVmAction,
        Func<string, Task> openConsoleAction,
        Func<string, Task> createSnapshotAction,
        Func<string, string, string?, Task> connectVmToSwitchAction,
        Func<UsbIpDeviceInfo?> getSelectedUsbDevice,
        Func<Task> refreshUsbDevicesAction,
        Func<Task> shareSelectedUsbAction,
        Func<Task> unshareSelectedUsbAction,
        Action exitAction)
    {
        _showMainWindowAction = showMainWindowAction;
        _hideMainWindowAction = hideMainWindowAction;
        _isMainWindowVisible = isMainWindowVisible;
        _getUiTheme = getUiTheme;
        _getVms = getVms;
        _getVmAdapters = getVmAdapters;
        _getSwitches = getSwitches;
        _getUsbDevices = getUsbDevices;
        _selectUsbDeviceAction = selectUsbDeviceAction;
        _isTrayMenuEnabled = isTrayMenuEnabled;
        _refreshTrayDataAction = refreshTrayDataAction;
        _startVmAction = startVmAction;
        _stopVmAction = stopVmAction;
        _restartVmAction = restartVmAction;
        _openConsoleAction = openConsoleAction;
        _createSnapshotAction = createSnapshotAction;
        _connectVmToSwitchAction = connectVmToSwitchAction;
        _getSelectedUsbDevice = getSelectedUsbDevice;
        _refreshUsbDevicesAction = refreshUsbDevicesAction;
        _shareSelectedUsbAction = shareSelectedUsbAction;
        _unshareSelectedUsbAction = unshareSelectedUsbAction;
        _exitAction = exitAction;
        _isInitialized = true;
    }

    public void ToggleFull()
    {
        if (!_isInitialized)
        {
            return;
        }

        Enqueue(() =>
        {
            _ = ToggleInternalAsync(TrayControlCenterMode.Full);
        });
    }

    public void ToggleCompact()
    {
        if (!_isInitialized)
        {
            return;
        }

        Enqueue(() =>
        {
            _ = ToggleInternalAsync(TrayControlCenterMode.Compact);
        });
    }

    private Task ToggleInternalAsync(TrayControlCenterMode mode)
    {
        try
        {
            EnsureWindow();
            if (_window is null)
            {
                return Task.CompletedTask;
            }

            if (_window.AppWindow.IsVisible && _mode == mode)
            {
                _window.AppWindow.Hide();
                return Task.CompletedTask;
            }

            _mode = mode;

            ReloadDataFromSources();
            UpdateWindowTheme();
            UpdateWindowView();
            PositionWindowNearTray();

            _window.AppWindow.Show();
            _window.Activate();

            StartBackgroundRefreshIfNeeded(force: true);
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Tray control center toggle failed.");

            try
            {
                _window?.AppWindow.Hide();
            }
            catch
            {
            }
        }

        return Task.CompletedTask;
    }

    public void Hide()
    {
        Enqueue(() =>
        {
            if (_window?.AppWindow.IsVisible == true)
            {
                _window.AppWindow.Hide();
            }
        });
    }

    public void Dispose()
    {
        Enqueue(() =>
        {
            if (_window is not null)
            {
                _window.Close();
                _window = null;
            }
        });
    }

    private void EnsureWindow()
    {
        if (_window is not null)
        {
            return;
        }

        var window = new TrayControlCenterWindow();
        _window = window;

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

        window.VmSelected += OnVmSelected;
        window.NetworkAdapterSelected += OnNetworkAdapterSelected;
        window.UsbDeviceSelected += selectionKey => ExecuteAction(async () =>
        {
            if (string.IsNullOrWhiteSpace(selectionKey))
            {
                return;
            }

            var idx = _usbDevices.FindIndex(device =>
                string.Equals(BuildUsbSelectionKey(device), selectionKey, StringComparison.OrdinalIgnoreCase));
            if (idx >= 0)
            {
                _selectedUsbIndex = idx;
                _selectedUsbDevice = _usbDevices[idx];
            }

            if (_selectUsbDeviceAction is not null)
            {
                await _selectUsbDeviceAction(selectionKey);
            }

            await RefreshDataAsync(refreshBackend: false);
            UpdateWindowView();
        }, "tray-usb-select");
        window.StartRequested += () => ExecuteVmAction(_startVmAction, "tray-start");
        window.StopRequested += () => ExecuteVmAction(_stopVmAction, "tray-stop");
        window.RestartRequested += () => ExecuteVmAction(_restartVmAction, "tray-restart");
        window.OpenConsoleRequested += () => ExecuteVmAction(_openConsoleAction, "tray-console");
        window.SnapshotRequested += () => ExecuteVmAction(_createSnapshotAction, "tray-snapshot");
        window.SwitchSelected += OnSwitchSelected;
        window.UsbRefreshRequested += () => ExecuteAction(async () =>
        {
            if (_refreshUsbDevicesAction is null)
            {
                return;
            }

            await _refreshUsbDevicesAction();
            await RefreshDataAsync(refreshBackend: false);
            UpdateWindowView();
        }, "tray-usb-refresh");
        window.UsbShareRequested += () => ExecuteAction(async () =>
        {
            if (_shareSelectedUsbAction is null)
            {
                return;
            }

            await _shareSelectedUsbAction();
            await RefreshDataAsync(refreshBackend: false);
            UpdateWindowView();
        }, "tray-usb-share");
        window.UsbUnshareRequested += () => ExecuteAction(async () =>
        {
            if (_unshareSelectedUsbAction is null)
            {
                return;
            }

            var usb = GetSelectedUsb();
            if (usb is null || string.IsNullOrWhiteSpace(usb.BusId))
            {
                return;
            }

            if (_selectUsbDeviceAction is not null)
            {
                await _selectUsbDeviceAction(BuildUsbSelectionKey(usb));
            }

            await _unshareSelectedUsbAction();
            await RefreshDataAsync(refreshBackend: false);
            UpdateWindowView();
        }, "tray-usb-unshare");
        window.ToggleVisibilityRequested += OnToggleVisibilityRequested;
        window.ExitRequested += () => _exitAction?.Invoke();
        window.Closed += (_, _) =>
        {
            if (ReferenceEquals(_window, window))
            {
                _window = null;
            }
        };
    }

    private void OnToggleVisibilityRequested()
    {
        try
        {
            var isVisible = _isMainWindowVisible?.Invoke() ?? true;
            if (isVisible)
            {
                _hideMainWindowAction?.Invoke();
            }
            else
            {
                _showMainWindowAction?.Invoke();
            }
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Tray compact visibility toggle failed.");
        }

        UpdateWindowView();
    }

    private void OnVmSelected(string vmName)
    {
        if (string.IsNullOrWhiteSpace(vmName) || _vms.Count == 0)
        {
            return;
        }

        var idx = _vms.FindIndex(vm => string.Equals(vm.Name, vmName, StringComparison.OrdinalIgnoreCase));
        if (idx < 0 || idx == _selectedVmIndex)
        {
            return;
        }

        ExecuteAction(async () =>
        {
            _selectedVmIndex = idx;
            _selectedVmAdapterName = null;
            SyncSelectedSwitchWithVm();
            await RefreshSelectedVmAdaptersAsync();
            SyncSelectedSwitchWithVm();
            UpdateWindowView();
        }, "tray-vm-select");
    }

    private void OnNetworkAdapterSelected(string adapterName)
    {
        _selectedVmAdapterName = string.IsNullOrWhiteSpace(adapterName) ? null : adapterName.Trim();
        SyncSelectedSwitchWithVm();
        UpdateWindowView();
    }

    private void OnSwitchSelected(string switchName)
    {
        _selectedSwitchName = string.IsNullOrWhiteSpace(switchName) ? null : switchName;
        UpdateWindowView();

        if (_isBusy || string.IsNullOrWhiteSpace(_selectedSwitchName) || _connectVmToSwitchAction is null)
        {
            return;
        }

        var vm = GetSelectedVm();
        if (vm is null)
        {
            return;
        }

        var selectedAdapter = _vmAdapters.FirstOrDefault(adapter =>
            string.Equals(adapter.Name, _selectedVmAdapterName, StringComparison.OrdinalIgnoreCase));
        var currentSwitch = NormalizeSwitchDisplayName(selectedAdapter?.SwitchName ?? vm.RuntimeSwitchName);
        if (string.Equals(currentSwitch, _selectedSwitchName, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var selectedSwitch = _selectedSwitchName;
        ExecuteAction(async () =>
        {
            var adapterName = _vmAdapters.Count > 1 ? _selectedVmAdapterName : null;
            await _connectVmToSwitchAction(vm.Name, selectedSwitch, adapterName);
            await RefreshDataAsync();
            UpdateWindowView();
        }, "tray-switch-select-connect");
    }

    private void ExecuteVmAction(Func<string, Task>? action, string actionName)
    {
        var vm = GetSelectedVm();
        if (vm is null || action is null)
        {
            return;
        }

        ExecuteAction(async () =>
        {
            await action(vm.Name);
            await RefreshDataAsync();
            UpdateWindowView();
        }, actionName);
    }

    private void ExecuteAction(Func<Task> action, string actionName)
    {
        if (_isBusy)
        {
            return;
        }

        _isBusy = true;
        UpdateWindowView();

        _ = ExecuteActionAsync(action, actionName);
    }

    private async Task ExecuteActionAsync(Func<Task> action, string actionName)
    {
        try
        {
            await action();
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Tray control center action failed: {ActionName}", actionName);
        }
        finally
        {
            Enqueue(() =>
            {
                _isBusy = false;
                UpdateWindowView();
            });
        }
    }

    private async Task RefreshDataAsync()
    {
        await RefreshDataAsync(refreshBackend: true);
    }

    private async Task RefreshDataAsync(bool refreshBackend)
    {
        await _refreshGate.WaitAsync();
        try
        {
            if (refreshBackend && _refreshTrayDataAction is not null)
            {
                try
                {
                    await _refreshTrayDataAction();
                    _lastBackendRefreshUtc = DateTime.UtcNow;
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "Tray control center data refresh failed.");
                }
            }

            ReloadDataFromSources();
        }
        finally
        {
            _refreshGate.Release();
        }
    }

    private void ReloadDataFromSources()
    {
        var previousVmName = GetSelectedVm()?.Name;

        _vms.Clear();
        _vms.AddRange(_getVms?.Invoke() ?? []);

        _switches.Clear();
        _switches.AddRange((_getSwitches?.Invoke() ?? []).OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase));

        var preferredUsb = _getSelectedUsbDevice?.Invoke();
        _usbDevices.Clear();
        _usbDevices.AddRange(_getUsbDevices?.Invoke() ?? []);

        if (_usbDevices.Count == 0)
        {
            _selectedUsbIndex = -1;
            _selectedUsbDevice = null;
        }
        else
        {
                var preferredSelectionKey = preferredUsb is null ? null : BuildUsbSelectionKey(preferredUsb);
                if (!string.IsNullOrWhiteSpace(preferredSelectionKey))
            {
                _selectedUsbIndex = _usbDevices.FindIndex(device =>
                    string.Equals(BuildUsbSelectionKey(device), preferredSelectionKey, StringComparison.OrdinalIgnoreCase));
            }

            if (_selectedUsbIndex < 0 || _selectedUsbIndex >= _usbDevices.Count)
            {
                _selectedUsbIndex = 0;
            }

            _selectedUsbDevice = _usbDevices[_selectedUsbIndex];
        }

        if (_vms.Count == 0)
        {
            _selectedVmIndex = -1;
            _selectedVmAdapterName = null;
            _vmAdapters.Clear();
            _selectedSwitchName = null;
            return;
        }

        if (!string.IsNullOrWhiteSpace(previousVmName))
        {
            var idx = _vms.FindIndex(vm => string.Equals(vm.Name, previousVmName, StringComparison.OrdinalIgnoreCase));
            _selectedVmIndex = idx >= 0 ? idx : 0;
        }
        else if (_selectedVmIndex < 0 || _selectedVmIndex >= _vms.Count)
        {
            _selectedVmIndex = 0;
        }

        SyncSelectedSwitchWithVm();
    }

    private void StartBackgroundRefreshIfNeeded(bool force)
    {
        if (_isBackgroundRefreshRunning)
        {
            return;
        }

        var shouldRefresh = force || DateTime.UtcNow - _lastBackendRefreshUtc >= ToggleRefreshInterval;
        if (!shouldRefresh)
        {
            return;
        }

        _isBackgroundRefreshRunning = true;
        _ = RunBackgroundRefreshAsync();
    }

    private async Task RefreshSelectedVmAdaptersAsync()
    {
        _vmAdapters.Clear();

        var vm = GetSelectedVm();
        if (vm is null || _getVmAdapters is null)
        {
            _selectedVmAdapterName = null;
            return;
        }

        try
        {
            var adapters = await _getVmAdapters(vm.Name);
            _vmAdapters.AddRange(adapters
                .Where(adapter => !string.IsNullOrWhiteSpace(adapter.Name))
                .OrderBy(adapter => adapter.DisplayName, StringComparer.OrdinalIgnoreCase));
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Tray control center adapter refresh failed for VM {VmName}", vm.Name);
        }

        if (_vmAdapters.Count <= 1)
        {
            _selectedVmAdapterName = null;
            return;
        }

        if (!string.IsNullOrWhiteSpace(_selectedVmAdapterName)
            && _vmAdapters.Any(adapter => string.Equals(adapter.Name, _selectedVmAdapterName, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        var configuredAdapterName = vm.TrayAdapterName?.Trim();
        if (!string.IsNullOrWhiteSpace(configuredAdapterName)
            && _vmAdapters.Any(adapter => string.Equals(adapter.Name, configuredAdapterName, StringComparison.OrdinalIgnoreCase)))
        {
            _selectedVmAdapterName = configuredAdapterName;
            return;
        }

        _selectedVmAdapterName = _vmAdapters[0].Name;
    }

    private async Task RunBackgroundRefreshAsync()
    {
        try
        {
            await RefreshDataAsync(refreshBackend: true);
            await RefreshSelectedVmAdaptersAsync();
            SyncSelectedSwitchWithVm();
            Enqueue(UpdateWindowView);
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Tray control center background refresh failed.");
        }
        finally
        {
            Enqueue(() => _isBackgroundRefreshRunning = false);
        }
    }

    private void SyncSelectedSwitchWithVm()
    {
        var vm = GetSelectedVm();
        if (vm is null)
        {
            _selectedSwitchName = null;
            return;
        }

        if (!string.IsNullOrWhiteSpace(_selectedSwitchName)
            && _switches.Any(vmSwitch => string.Equals(vmSwitch.Name, _selectedSwitchName, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        var selectedAdapter = _vmAdapters.FirstOrDefault(adapter =>
            string.Equals(adapter.Name, _selectedVmAdapterName, StringComparison.OrdinalIgnoreCase));
        var runtimeSwitch = NormalizeSwitchDisplayName(selectedAdapter?.SwitchName ?? vm.RuntimeSwitchName);
        var runtimeMatch = _switches.FirstOrDefault(vmSwitch => string.Equals(vmSwitch.Name, runtimeSwitch, StringComparison.OrdinalIgnoreCase));
        _selectedSwitchName = runtimeMatch?.Name ?? _switches.FirstOrDefault()?.Name;
    }

    private void UpdateWindowTheme()
    {
        if (_window is null)
        {
            return;
        }

        var isDark = string.Equals(_getUiTheme?.Invoke(), "Dark", StringComparison.OrdinalIgnoreCase);
        _window.ApplyTheme(isDark);
    }

    private void UpdateWindowView()
    {
        if (_window is null)
        {
            return;
        }

        var trayEnabled = _isTrayMenuEnabled?.Invoke() ?? true;
        var vm = GetSelectedVm();
        var hasVm = vm is not null;
        var runtimeState = vm?.RuntimeState?.Trim() ?? "Unbekannt";
        var selectedAdapter = _vmAdapters.FirstOrDefault(adapter =>
            string.Equals(adapter.Name, _selectedVmAdapterName, StringComparison.OrdinalIgnoreCase));
        var runtimeSwitch = NormalizeSwitchDisplayName(selectedAdapter?.SwitchName ?? vm?.RuntimeSwitchName);

        var state = new TrayControlCenterViewState
        {
            IsCompactMode = _mode == TrayControlCenterMode.Compact,
            HasVm = hasVm,
            CanMoveVm = trayEnabled && _vms.Count > 1,
            CanStart = trayEnabled && hasVm && !IsVmRunning(runtimeState),
            CanStop = trayEnabled && hasVm && IsVmRunning(runtimeState),
            CanRestart = trayEnabled && hasVm,
            SelectedVmDisplay = hasVm ? vm!.DisplayLabel : "Keine VM ausgewählt",
            SelectedVmMeta = hasVm ? $"{runtimeState} · {runtimeSwitch}" : "-",
            SelectedVmName = hasVm ? vm!.Name : null,
            ShowNetworkAdapterSelection = true,
            CanSelectNetworkAdapter = trayEnabled && hasVm && _vmAdapters.Count > 1,
            SelectedNetworkAdapterName = _vmAdapters.Count > 1 ? _selectedVmAdapterName : null,
            ActiveSwitchDisplay = $"Aktiv: {runtimeSwitch}",
            VisibilityButtonText = (_isMainWindowVisible?.Invoke() ?? true) ? "⌂  Ausblenden" : "⌂  Einblenden",
            UsbSelectedDisplay = BuildUsbSelectedDisplay(_selectedUsbDevice),
            SelectedUsbKey = _selectedUsbDevice is null ? null : BuildUsbSelectionKey(_selectedUsbDevice),
            CanUsbRefresh = trayEnabled,
            CanUsbShare = trayEnabled && CanUsbShare(_selectedUsbDevice),
            CanUsbUnshare = trayEnabled && CanUsbUnshare(_selectedUsbDevice)
        };

        foreach (var vmItem in _vms)
        {
            var label = string.IsNullOrWhiteSpace(vmItem.DisplayLabel)
                ? vmItem.Name
                : vmItem.DisplayLabel;

            state.Vms.Add(new TrayVmItem(vmItem.Name, label));
        }

        foreach (var adapter in _vmAdapters)
        {
            state.NetworkAdapters.Add(new TrayVmAdapterItem(
                adapter.Name,
                adapter.DisplayName,
                NormalizeSwitchDisplayName(adapter.SwitchName)));
        }

        foreach (var usbDevice in _usbDevices)
        {
            var label = string.IsNullOrWhiteSpace(usbDevice.Description)
                ? usbDevice.BusId
                : usbDevice.Description;

            if (string.IsNullOrWhiteSpace(label))
            {
                label = "USB-Gerät";
            }

            state.UsbDevices.Add(new TrayUsbItem(BuildUsbSelectionKey(usbDevice), label));
        }

        foreach (var vmSwitch in _switches)
        {
            state.Switches.Add(new TraySwitchItem(
                vmSwitch.Name,
                vmSwitch.Name,
                string.Equals(vmSwitch.Name, _selectedSwitchName, StringComparison.OrdinalIgnoreCase),
                trayEnabled && hasVm));
        }

        if (_isBusy)
        {
            state.CanMoveVm = false;
            state.CanStart = false;
            state.CanStop = false;
            state.CanRestart = false;
            state.CanSelectNetworkAdapter = false;
            state.CanUsbRefresh = false;
            state.CanUsbShare = false;
            state.CanUsbUnshare = false;
        }

        _window.UpdateView(state);
    }

    private void PositionWindowNearTray()
    {
        if (_window is null)
        {
            return;
        }

        var popupWidth = GetPopupWidth(_mode);
        var popupHeight = GetPopupHeight(_mode);
        _window.SetPanelSize(popupWidth, popupHeight);

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

        _window.SetPosition(x, y);
    }

    private VmDefinition? GetSelectedVm()
    {
        if (_selectedVmIndex < 0 || _selectedVmIndex >= _vms.Count)
        {
            return null;
        }

        return _vms[_selectedVmIndex];
    }

    private static bool IsVmRunning(string? state)
    {
        if (string.IsNullOrWhiteSpace(state))
        {
            return false;
        }

        return state.Contains("Running", StringComparison.OrdinalIgnoreCase)
               || state.Contains("Wird ausgeführt", StringComparison.OrdinalIgnoreCase)
               || state.Contains("Ausgeführt", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeSwitchDisplayName(string? switchName)
    {
        return string.IsNullOrWhiteSpace(switchName)
            ? "Nicht verbunden"
            : switchName.Trim();
    }

    private static bool CanUsbShare(UsbIpDeviceInfo? usbDevice)
    {
        return usbDevice is not null
               && !string.IsNullOrWhiteSpace(usbDevice.BusId)
               && !usbDevice.IsShared;
    }

    private static bool CanUsbUnshare(UsbIpDeviceInfo? usbDevice)
    {
        return usbDevice is not null
               && !string.IsNullOrWhiteSpace(usbDevice.BusId)
               && usbDevice.IsShared;
    }

    private static string BuildUsbSelectedDisplay(UsbIpDeviceInfo? usbDevice)
    {
        if (usbDevice is null)
        {
            return "Selected: -";
        }

        var name = string.IsNullOrWhiteSpace(usbDevice.Description)
            ? "-"
            : usbDevice.Description.Trim();
        var status = usbDevice.IsShared ? "Shared" : "Not shared";
        return $"Selected: {name} ({status})";
    }

    private static string BuildUsbSelectionKey(UsbIpDeviceInfo usbDevice)
    {
        if (!string.IsNullOrWhiteSpace(usbDevice.BusId))
        {
            return "busid:" + usbDevice.BusId.Trim();
        }

        if (!string.IsNullOrWhiteSpace(usbDevice.PersistedGuid))
        {
            return "guid:" + usbDevice.PersistedGuid.Trim();
        }

        return "instance:" + (usbDevice.InstanceId?.Trim() ?? string.Empty);
    }

    private UsbIpDeviceInfo? GetSelectedUsb()
    {
        if (_selectedUsbIndex < 0 || _selectedUsbIndex >= _usbDevices.Count)
        {
            return null;
        }

        return _usbDevices[_selectedUsbIndex];
    }

    private void Enqueue(Action action)
    {
        if (_dispatcherQueue.HasThreadAccess)
        {
            action();
            return;
        }

        _dispatcherQueue.TryEnqueue(() =>
        {
            lock (_syncLock)
            {
                action();
            }
        });
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetCursorPos(out NativePoint lpPoint);

    [StructLayout(LayoutKind.Sequential)]
    private struct NativePoint
    {
        public int X;
        public int Y;
    }

    private enum TrayControlCenterMode
    {
        Full,
        Compact
    }

    private static int GetPopupWidth(TrayControlCenterMode mode)
    {
        return mode == TrayControlCenterMode.Compact ? 228 : 404;
    }

    private static int GetPopupHeight(TrayControlCenterMode mode)
    {
        if (mode == TrayControlCenterMode.Compact)
        {
            return 196;
        }

        return 740;
    }
}
