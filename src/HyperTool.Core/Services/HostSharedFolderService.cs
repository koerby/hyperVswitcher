using HyperTool.Models;

namespace HyperTool.Services;

public sealed class HostSharedFolderService : IHostSharedFolderService
{
    public Task EnsureShareAsync(HostSharedFolderDefinition definition, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(definition);

        var shareName = (definition.ShareName ?? string.Empty).Trim();
        var localPath = (definition.LocalPath ?? string.Empty).Trim();

        if (string.IsNullOrWhiteSpace(shareName))
        {
            throw new ArgumentException("Freigabename darf nicht leer sein.", nameof(definition));
        }

        if (string.IsNullOrWhiteSpace(localPath))
        {
            throw new ArgumentException("Lokaler Ordnerpfad darf nicht leer sein.", nameof(definition));
        }

        if (!Directory.Exists(localPath))
        {
            throw new DirectoryNotFoundException($"Lokaler Ordner '{localPath}' wurde nicht gefunden.");
        }

        return Task.CompletedTask;
    }

    public Task RemoveShareAsync(string shareName, CancellationToken cancellationToken)
    {
        var trimmed = (shareName ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            throw new ArgumentException("Freigabename darf nicht leer sein.", nameof(shareName));
        }

        return Task.CompletedTask;
    }

    public Task<bool> ShareExistsAsync(string shareName, CancellationToken cancellationToken)
    {
        var trimmed = (shareName ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return Task.FromResult(false);
        }
        return Task.FromResult(true);
    }
}
