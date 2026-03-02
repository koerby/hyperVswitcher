using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;

namespace HyperTool.Services;

public sealed class GitHubUpdateService : IUpdateService
{
    private static readonly HttpClient HttpClient = new();

    static GitHubUpdateService()
    {
        HttpClient.DefaultRequestHeaders.UserAgent.Clear();
        HttpClient.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("HyperTool", "1.0"));
        HttpClient.DefaultRequestHeaders.Accept.Clear();
        HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
    }

    public async Task<(bool Success, bool HasUpdate, string Message, string? LatestVersion, string? ReleaseUrl, string? InstallerDownloadUrl, string? InstallerFileName)> CheckForUpdateAsync(
        string owner,
        string repo,
        string currentVersion,
        CancellationToken cancellationToken,
        string? installerAssetHint = null)
    {
        if (string.IsNullOrWhiteSpace(owner) || string.IsNullOrWhiteSpace(repo))
        {
            return (false, false, "GitHub Owner/Repo fehlt.", null, null, null, null);
        }

        var repoUrl = $"https://github.com/{owner}/{repo}";
        var releasePageUrl = $"{repoUrl}/releases";
        var latestReleaseApi = $"https://api.github.com/repos/{owner}/{repo}/releases/latest";
        var tagsApi = $"https://api.github.com/repos/{owner}/{repo}/tags?per_page=1";

        try
        {
            using var releaseResponse = await HttpClient.GetAsync(latestReleaseApi, cancellationToken);
            if (releaseResponse.IsSuccessStatusCode)
            {
                var json = await releaseResponse.Content.ReadAsStringAsync(cancellationToken);
                using var doc = JsonDocument.Parse(json);

                var latestTag = doc.RootElement.TryGetProperty("tag_name", out var tagElement)
                    ? (tagElement.GetString() ?? string.Empty).Trim()
                    : string.Empty;

                var htmlUrl = doc.RootElement.TryGetProperty("html_url", out var urlElement)
                    ? urlElement.GetString()
                    : releasePageUrl;

                var installerAsset = TryFindInstallerAsset(doc.RootElement, installerAssetHint);

                if (string.IsNullOrWhiteSpace(latestTag))
                {
                    return (false, false, "Keine Versionsinformation im neuesten Release gefunden.", null, htmlUrl, installerAsset?.DownloadUrl, installerAsset?.Name);
                }

                var releaseComparison = CompareVersions(currentVersion, latestTag, htmlUrl);
                return (releaseComparison.Success, releaseComparison.HasUpdate, releaseComparison.Message, releaseComparison.LatestVersion, releaseComparison.ReleaseUrl, installerAsset?.DownloadUrl, installerAsset?.Name);
            }

            if (releaseResponse.StatusCode != System.Net.HttpStatusCode.NotFound)
            {
                var errorMessage = await TryReadGitHubErrorMessageAsync(releaseResponse, cancellationToken);
                return (false, false, $"GitHub API Fehler: {(int)releaseResponse.StatusCode} {errorMessage}", null, releasePageUrl, null, null);
            }

            using var tagsResponse = await HttpClient.GetAsync(tagsApi, cancellationToken);
            if (!tagsResponse.IsSuccessStatusCode)
            {
                var tagsErrorMessage = await TryReadGitHubErrorMessageAsync(tagsResponse, cancellationToken);
                return (false, false, $"GitHub API Fehler: {(int)tagsResponse.StatusCode} {tagsErrorMessage}", null, releasePageUrl, null, null);
            }

            var tagsJson = await tagsResponse.Content.ReadAsStringAsync(cancellationToken);
            using var tagsDoc = JsonDocument.Parse(tagsJson);

            if (tagsDoc.RootElement.ValueKind != JsonValueKind.Array || tagsDoc.RootElement.GetArrayLength() == 0)
            {
                return (true, false, "Keine Releases oder Tags gefunden.", null, releasePageUrl, null, null);
            }

            var firstTag = tagsDoc.RootElement[0];
            var tagName = firstTag.TryGetProperty("name", out var nameElement)
                ? (nameElement.GetString() ?? string.Empty).Trim()
                : string.Empty;

            if (string.IsNullOrWhiteSpace(tagName))
            {
                return (false, false, "Tag-Information ohne Versionsnamen gefunden.", null, releasePageUrl, null, null);
            }

            var tagComparison = CompareVersions(currentVersion, tagName, $"{repoUrl}/tags");
            var prefix = tagComparison.HasUpdate ? "Update verfügbar (Tag): " : "Kein neuer Release gefunden, letzter Tag: ";
            return (tagComparison.Success, tagComparison.HasUpdate, prefix + tagName, tagName, tagComparison.ReleaseUrl, null, null);
        }
        catch (Exception ex)
        {
            return (false, false, $"Updatecheck fehlgeschlagen: {ex.Message}", null, releasePageUrl, null, null);
        }
    }

    private static (string Name, string DownloadUrl)? TryFindInstallerAsset(JsonElement releaseRoot, string? installerAssetHint)
    {
        if (!releaseRoot.TryGetProperty("assets", out var assetsElement) || assetsElement.ValueKind != JsonValueKind.Array)
        {
            return null;
        }

        var candidates = new List<(string Name, string DownloadUrl, int Score)>();

        foreach (var asset in assetsElement.EnumerateArray())
        {
            var name = asset.TryGetProperty("name", out var nameElement)
                ? (nameElement.GetString() ?? string.Empty).Trim()
                : string.Empty;

            var downloadUrl = asset.TryGetProperty("browser_download_url", out var urlElement)
                ? (urlElement.GetString() ?? string.Empty).Trim()
                : string.Empty;

            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(downloadUrl))
            {
                continue;
            }

            var lowerName = name.ToLowerInvariant();
            if (!lowerName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
                && !lowerName.EndsWith(".msi", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var score = 0;
            if (lowerName.Contains("setup", StringComparison.OrdinalIgnoreCase) || lowerName.Contains("installer", StringComparison.OrdinalIgnoreCase))
            {
                score += 4;
            }

            if (lowerName.Contains("install", StringComparison.OrdinalIgnoreCase))
            {
                score += 2;
            }

            if (lowerName.Contains("hypertool", StringComparison.OrdinalIgnoreCase))
            {
                score += 1;
            }

            if (lowerName.EndsWith(".msi", StringComparison.OrdinalIgnoreCase))
            {
                score += 1;
            }

            if (!string.IsNullOrWhiteSpace(installerAssetHint)
                && lowerName.Contains(installerAssetHint.Trim().ToLowerInvariant(), StringComparison.OrdinalIgnoreCase))
            {
                score += 8;
            }

            candidates.Add((name, downloadUrl, score));
        }

        var selected = candidates
            .OrderByDescending(candidate => candidate.Score)
            .ThenBy(candidate => candidate.Name, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault();

        return string.IsNullOrWhiteSpace(selected.Name) || string.IsNullOrWhiteSpace(selected.DownloadUrl)
            ? null
            : (selected.Name, selected.DownloadUrl);
    }

    private static (bool Success, bool HasUpdate, string Message, string? LatestVersion, string? ReleaseUrl) CompareVersions(
        string currentVersion,
        string latestTag,
        string? releaseUrl)
    {
        var latestParsed = SemanticVersion.TryParse(latestTag, out var latestSemVersion);
        var currentParsed = SemanticVersion.TryParse(currentVersion, out var currentSemVersion);

        if (latestParsed && currentParsed)
        {
            var semverHasUpdate = latestSemVersion.CompareTo(currentSemVersion) > 0;
            return (true, semverHasUpdate, semverHasUpdate ? $"Update verfügbar: {latestTag}" : "Bereits aktuell.", latestTag, releaseUrl);
        }

        var latestVersion = NormalizeVersion(latestTag);
        var current = NormalizeVersion(currentVersion);
        var hasUpdate = !string.Equals(latestVersion, current, StringComparison.OrdinalIgnoreCase);
        return (true, hasUpdate, hasUpdate ? $"Update verfügbar: {latestTag}" : "Bereits aktuell.", latestTag, releaseUrl);
    }

    private static async Task<string> TryReadGitHubErrorMessageAsync(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        try
        {
            var content = await response.Content.ReadAsStringAsync(cancellationToken);
            if (string.IsNullOrWhiteSpace(content))
            {
                return string.Empty;
            }

            using var doc = JsonDocument.Parse(content);
            if (doc.RootElement.TryGetProperty("message", out var messageElement))
            {
                var message = messageElement.GetString();
                return string.IsNullOrWhiteSpace(message) ? string.Empty : $"- {message}";
            }
        }
        catch
        {
        }

        return string.Empty;
    }

    private static string NormalizeVersion(string value)
    {
        var normalized = value.Trim();
        if (normalized.StartsWith("v", StringComparison.OrdinalIgnoreCase))
        {
            normalized = normalized[1..];
        }

        return normalized;
    }

    private readonly struct SemanticVersion : IComparable<SemanticVersion>
    {
        private readonly int[] _core;
        private readonly string[] _prerelease;

        private SemanticVersion(int[] core, string[] prerelease)
        {
            _core = core;
            _prerelease = prerelease;
        }

        public static bool TryParse(string? value, out SemanticVersion version)
        {
            version = default;

            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var normalized = NormalizeVersion(value);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            var withoutBuildMetadata = normalized.Split('+', 2)[0];
            var prereleaseIndex = withoutBuildMetadata.IndexOf('-');

            var corePart = prereleaseIndex >= 0
                ? withoutBuildMetadata[..prereleaseIndex]
                : withoutBuildMetadata;

            var prereleasePart = prereleaseIndex >= 0
                ? withoutBuildMetadata[(prereleaseIndex + 1)..]
                : string.Empty;

            var coreSegments = corePart.Split('.', StringSplitOptions.RemoveEmptyEntries);
            if (coreSegments.Length == 0)
            {
                return false;
            }

            var coreNumbers = new int[coreSegments.Length];
            for (var index = 0; index < coreSegments.Length; index++)
            {
                if (!int.TryParse(coreSegments[index], out coreNumbers[index]) || coreNumbers[index] < 0)
                {
                    return false;
                }
            }

            var prereleaseSegments = string.IsNullOrWhiteSpace(prereleasePart)
                ? []
                : prereleasePart.Split('.', StringSplitOptions.RemoveEmptyEntries);

            version = new SemanticVersion(coreNumbers, prereleaseSegments);
            return true;
        }

        public int CompareTo(SemanticVersion other)
        {
            var segmentCount = Math.Max(_core.Length, other._core.Length);
            for (var index = 0; index < segmentCount; index++)
            {
                var left = index < _core.Length ? _core[index] : 0;
                var right = index < other._core.Length ? other._core[index] : 0;
                if (left != right)
                {
                    return left.CompareTo(right);
                }
            }

            var hasPrerelease = _prerelease.Length > 0;
            var otherHasPrerelease = other._prerelease.Length > 0;

            if (!hasPrerelease && !otherHasPrerelease)
            {
                return 0;
            }

            if (!hasPrerelease)
            {
                return 1;
            }

            if (!otherHasPrerelease)
            {
                return -1;
            }

            var prereleaseCount = Math.Max(_prerelease.Length, other._prerelease.Length);
            for (var index = 0; index < prereleaseCount; index++)
            {
                if (index >= _prerelease.Length)
                {
                    return -1;
                }

                if (index >= other._prerelease.Length)
                {
                    return 1;
                }

                var left = _prerelease[index];
                var right = other._prerelease[index];

                var leftIsNumeric = int.TryParse(left, out var leftNumber);
                var rightIsNumeric = int.TryParse(right, out var rightNumber);

                if (leftIsNumeric && rightIsNumeric)
                {
                    if (leftNumber != rightNumber)
                    {
                        return leftNumber.CompareTo(rightNumber);
                    }

                    continue;
                }

                if (leftIsNumeric != rightIsNumeric)
                {
                    return leftIsNumeric ? -1 : 1;
                }

                var compare = string.Compare(left, right, StringComparison.OrdinalIgnoreCase);
                if (compare != 0)
                {
                    return compare;
                }
            }

            return 0;
        }
    }
}
