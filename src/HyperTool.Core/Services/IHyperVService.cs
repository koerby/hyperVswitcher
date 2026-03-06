using HyperTool.Models;

namespace HyperTool.Services;

public interface IHyperVService
{
    Task<IReadOnlyList<HyperVVmInfo>> GetVmsAsync(CancellationToken cancellationToken);

    Task StartVmAsync(string vmName, CancellationToken cancellationToken);

    Task StopVmGracefulAsync(string vmName, CancellationToken cancellationToken);

    Task TurnOffVmAsync(string vmName, CancellationToken cancellationToken);

    Task RestartVmAsync(string vmName, CancellationToken cancellationToken);

    Task<IReadOnlyList<HyperVSwitchInfo>> GetVmSwitchesAsync(CancellationToken cancellationToken);

    Task<IReadOnlyList<HyperVVmNetworkAdapterInfo>> GetVmNetworkAdaptersAsync(string vmName, CancellationToken cancellationToken);

    Task<string?> GetVmCurrentSwitchNameAsync(string vmName, CancellationToken cancellationToken);

    Task ConnectVmNetworkAdapterAsync(string vmName, string switchName, string? adapterName, CancellationToken cancellationToken);

    Task DisconnectVmNetworkAdapterAsync(string vmName, string? adapterName, CancellationToken cancellationToken);

    Task RenameVmNetworkAdapterAsync(string vmName, string adapterName, string newAdapterName, CancellationToken cancellationToken);

    Task<string> GetHostNetworkProfileCategoryAsync(CancellationToken cancellationToken);

    Task SetHostNetworkProfileCategoryAsync(string adapterName, string networkCategory, CancellationToken cancellationToken);

    Task<IReadOnlyList<HostNetworkAdapterInfo>> GetHostNetworkAdaptersWithUplinkAsync(CancellationToken cancellationToken);

    Task<IReadOnlyList<HyperVCheckpointInfo>> GetCheckpointsAsync(string vmName, CancellationToken cancellationToken);

    Task CreateCheckpointAsync(string vmName, string checkpointName, string? description, CancellationToken cancellationToken);

    Task ApplyCheckpointAsync(string vmName, string checkpointName, string? checkpointId, CancellationToken cancellationToken);

    Task RemoveCheckpointAsync(string vmName, string checkpointName, string? checkpointId, CancellationToken cancellationToken);

    Task OpenVmConnectAsync(string vmName, string computerName, bool openWithSessionEdit, CancellationToken cancellationToken);

    Task ReopenVmConnectWithSessionEditAsync(string vmName, string computerName, CancellationToken cancellationToken);

    Task<(bool HasEnoughSpace, long RequiredBytes, long AvailableBytes, string TargetDrive)> CheckExportDiskSpaceAsync(string vmName, string destinationPath, CancellationToken cancellationToken);

    Task ExportVmAsync(string vmName, string destinationPath, IProgress<int>? progress, CancellationToken cancellationToken);

    Task<ImportVmResult> ImportVmAsync(
        string importPath,
        string destinationPath,
        string? requestedVmName,
        string? requestedFolderName,
        string importMode,
        IProgress<int>? progress,
        CancellationToken cancellationToken);
}