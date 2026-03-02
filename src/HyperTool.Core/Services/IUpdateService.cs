namespace HyperTool.Services;

public interface IUpdateService
{
    Task<(bool Success, bool HasUpdate, string Message, string? LatestVersion, string? ReleaseUrl, string? InstallerDownloadUrl, string? InstallerFileName)> CheckForUpdateAsync(
        string owner,
        string repo,
        string currentVersion,
    CancellationToken cancellationToken,
    string? installerAssetHint = null);
}
