using HyperTool.Models;

namespace HyperTool.Services;

public interface ITrayService : IDisposable
{
    void Initialize(
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
        Action exitAction);

    void UpdateTrayMenu();
}