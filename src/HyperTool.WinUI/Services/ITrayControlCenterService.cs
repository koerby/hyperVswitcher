using HyperTool.Models;

namespace HyperTool.WinUI.Services;

internal interface ITrayControlCenterService : IDisposable
{
    void Initialize(
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
        Action exitAction);

    void ToggleFull();

    void ToggleCompact();

    void Hide();
}
