using HyperTool.Models;

namespace HyperTool.Services;

public interface IUsbIpService
{
    Task<bool> IsUsbClientAvailableAsync(CancellationToken cancellationToken);

    Task<IReadOnlyList<UsbIpDeviceInfo>> GetDevicesAsync(CancellationToken cancellationToken);

    Task<IReadOnlyList<UsbIpDeviceInfo>> GetRemoteDevicesAsync(string hostAddress, CancellationToken cancellationToken);

    Task AttachToHostAsync(string busId, string hostAddress, CancellationToken cancellationToken);

    Task DetachFromHostAsync(string busId, string? hostAddress, CancellationToken cancellationToken);

    Task BindAsync(string busId, bool force, CancellationToken cancellationToken);

    Task UnbindAsync(string busId, CancellationToken cancellationToken);

    Task UnbindByPersistedGuidAsync(string persistedGuid, CancellationToken cancellationToken);

    Task AttachToWslAsync(string busId, string? distribution, CancellationToken cancellationToken);

    Task DetachAsync(string busId, CancellationToken cancellationToken);
}
