using HyperTool.Models;
using HyperTool.Services;

namespace HyperTool.WinUI.Services;

public sealed class TrayService : ITrayService
{
    private readonly HyperTool.Services.TrayService _inner = new();

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
        Func<UsbIpDeviceInfo?> getSelectedUsbDevice,
        Func<Task> refreshUsbDevicesAction,
        Func<Task> shareSelectedUsbAction,
        Func<Task> unshareSelectedUsbAction,
        Action exitAction)
    {
        _inner.Initialize(
            showAction,
            hideAction,
            toggleControlCenterAction,
            toggleControlCenterCompactAction,
            hideControlCenterAction,
            getUiTheme,
            isWindowVisible,
            isTrayMenuEnabled,
            getVms,
            getSwitches,
            refreshTrayDataAction,
            subscribeTrayStateChanged,
            unsubscribeTrayStateChanged,
            startVmAction,
            stopVmAction,
            restartVmAction,
            openConsoleAction,
            createSnapshotAction,
            connectVmToSwitchAction,
            disconnectVmSwitchAction,
                getSelectedUsbDevice,
                refreshUsbDevicesAction,
                shareSelectedUsbAction,
                unshareSelectedUsbAction,
            exitAction);
    }

    public void UpdateTrayMenu()
    {
        _inner.UpdateTrayMenu();
    }

    public void Dispose()
    {
        _inner.Dispose();
    }
}
