using HyperTool.Models;
using HyperTool.Services;

namespace HyperTool.Guest;

internal static class GuestFileServiceModeHelper
{
    public static bool IsHyperToolFileMode(GuestConfig? config)
    {
        if (config?.FileService is null || !config.FileService.Enabled)
        {
            return false;
        }

        return string.Equals(
            GuestConfigService.NormalizeMappingMode(config.FileService.MappingMode),
            "hypertool-file",
            StringComparison.OrdinalIgnoreCase);
    }

    public static async Task<IReadOnlyList<HostFileServiceShareInfo>> FetchSharesAsync(CancellationToken cancellationToken)
    {
        var client = new HyperVSocketFileGuestClient();
        var response = await client.ListSharesAsync(cancellationToken);
        if (!response.Success)
        {
            throw new InvalidOperationException($"HyperTool File-Dienstfehler: {response.ErrorCode} {response.ErrorMessage}".Trim());
        }

        return response.Shares ?? [];
    }

    public static async Task EnsureShareAvailableAsync(string sharePath, CancellationToken cancellationToken)
    {
        var shareName = ResolveShareName(sharePath);
        if (string.IsNullOrWhiteSpace(shareName))
        {
            throw new InvalidOperationException("Ungültiger SharePath für HyperTool-File-Modus.");
        }

        var shares = await FetchSharesAsync(cancellationToken);
        var found = shares.Any(item =>
            string.Equals(item.Id, shareName, StringComparison.OrdinalIgnoreCase)
            || string.Equals(item.ShareName, shareName, StringComparison.OrdinalIgnoreCase));

        if (!found)
        {
            throw new InvalidOperationException($"Freigabe '{shareName}' wurde im HyperTool File-Dienst nicht gefunden.");
        }
    }

    public static string ResolveShareName(string? sharePath)
    {
        var value = (sharePath ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        if (!value.StartsWith("\\\\", StringComparison.Ordinal))
        {
            return value;
        }

        var withoutPrefix = value[2..];
        var firstSeparator = withoutPrefix.IndexOf('\\');
        if (firstSeparator <= 0 || firstSeparator == withoutPrefix.Length - 1)
        {
            return string.Empty;
        }

        var remainder = withoutPrefix[(firstSeparator + 1)..];
        var secondSeparator = remainder.IndexOf('\\');
        if (secondSeparator <= 0)
        {
            return remainder.Trim();
        }

        return remainder[..secondSeparator].Trim();
    }
}
