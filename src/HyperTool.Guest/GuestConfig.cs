using System.Text.Json;

namespace HyperTool.Guest;

internal sealed class GuestConfig
{
    public bool Enabled { get; set; } = true;

    public string SharePath { get; set; } = "\\\\HOST\\HyperToolShare";

    public string DriveLetter { get; set; } = "Z";

    public bool Persistent { get; set; } = true;

    public int PollIntervalSeconds { get; set; } = 15;

    public GuestCredential Credential { get; set; } = new();

    public GuestAutostartSettings Autostart { get; set; } = new();

    public GuestLoggingSettings Logging { get; set; } = new();

    public GuestHandshakeSettings Handshake { get; set; } = new();

    public GuestUsbSettings Usb { get; set; } = new();

    public GuestUiSettings Ui { get; set; } = new();
}

internal sealed class GuestUsbSettings
{
    public string HostAddress { get; set; } = string.Empty;
}

internal sealed class GuestUiSettings
{
    public string Theme { get; set; } = "dark";

    public bool StartWithWindows { get; set; } = false;

    public bool StartMinimized { get; set; } = false;

    public bool MinimizeToTray { get; set; } = true;
}

internal sealed class GuestCredential
{
    public string Username { get; set; } = string.Empty;

    public string Password { get; set; } = string.Empty;
}

internal sealed class GuestAutostartSettings
{
    public string PreferredMode { get; set; } = "run-registry";

    public string RunValueName { get; set; } = "HyperTool.Guest";

    public string TaskName { get; set; } = "HyperTool.Guest";
}

internal sealed class GuestLoggingSettings
{
    public string DirectoryPath { get; set; } = string.Empty;

    public string FileName { get; set; } = "hypertool-guest.ndjson";

    public bool EchoToConsole { get; set; } = true;
}

internal sealed class GuestHandshakeSettings
{
    public bool Enabled { get; set; } = true;

    public string FilePath { get; set; } = string.Empty;
}

internal static class GuestConfigService
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public static readonly string DefaultConfigPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
        "HyperTool",
        "HyperTool.Guest.json");

    public static readonly string DefaultLogDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
        "HyperTool",
        "logs");

    public static readonly string DefaultHandshakePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
        "HyperTool",
        "HyperTool.Guest.handshake.json");

    public static GuestConfig LoadOrCreate(string configPath, out bool created)
    {
        if (!File.Exists(configPath))
        {
            var config = new GuestConfig();
            Write(configPath, config);
            created = true;
            return config;
        }

        var raw = File.ReadAllText(configPath);
        var loaded = JsonSerializer.Deserialize<GuestConfig>(raw, SerializerOptions) ?? new GuestConfig();
        Normalize(loaded);
        created = false;
        return loaded;
    }

    public static void Save(string configPath, GuestConfig config)
    {
        Normalize(config);
        Write(configPath, config);
    }

    private static void Write(string configPath, GuestConfig config)
    {
        var directory = Path.GetDirectoryName(configPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var json = JsonSerializer.Serialize(config, SerializerOptions);
        File.WriteAllText(configPath, json);
    }

    private static void Normalize(GuestConfig config)
    {
        config.SharePath = (config.SharePath ?? string.Empty).Trim();
        config.DriveLetter = NormalizeDriveLetter(config.DriveLetter);
        config.PollIntervalSeconds = Math.Clamp(config.PollIntervalSeconds, 5, 3600);
        config.Credential ??= new GuestCredential();
        config.Credential.Username = (config.Credential.Username ?? string.Empty).Trim();
        config.Credential.Password ??= string.Empty;

        config.Autostart ??= new GuestAutostartSettings();
        config.Autostart.PreferredMode = NormalizeAutostartMode(config.Autostart.PreferredMode);
        config.Autostart.RunValueName = string.IsNullOrWhiteSpace(config.Autostart.RunValueName)
            ? "HyperTool.Guest"
            : config.Autostart.RunValueName.Trim();
        config.Autostart.TaskName = string.IsNullOrWhiteSpace(config.Autostart.TaskName)
            ? "HyperTool.Guest"
            : config.Autostart.TaskName.Trim();

        config.Logging ??= new GuestLoggingSettings();
        config.Logging.DirectoryPath = string.IsNullOrWhiteSpace(config.Logging.DirectoryPath)
            ? DefaultLogDirectory
            : config.Logging.DirectoryPath.Trim();
        config.Logging.FileName = string.IsNullOrWhiteSpace(config.Logging.FileName)
            ? "hypertool-guest.ndjson"
            : config.Logging.FileName.Trim();

        config.Handshake ??= new GuestHandshakeSettings();
        config.Handshake.FilePath = string.IsNullOrWhiteSpace(config.Handshake.FilePath)
            ? DefaultHandshakePath
            : config.Handshake.FilePath.Trim();

        config.Usb ??= new GuestUsbSettings();
        config.Usb.HostAddress = (config.Usb.HostAddress ?? string.Empty).Trim();

        config.Ui ??= new GuestUiSettings();
        config.Ui.Theme = NormalizeTheme(config.Ui.Theme);
    }

    public static string NormalizeDriveLetter(string? driveLetter)
    {
        var value = (driveLetter ?? "Z").Trim().TrimEnd(':');
        if (value.Length == 0)
        {
            return "Z";
        }

        var letter = char.ToUpperInvariant(value[0]);
        return letter is >= 'A' and <= 'Z' ? letter.ToString() : "Z";
    }

    public static string NormalizeAutostartMode(string? mode)
    {
        var normalized = (mode ?? "run-registry").Trim().ToLowerInvariant();
        return normalized is "run-registry" or "task"
            ? normalized
            : "run-registry";
    }

    public static string NormalizeTheme(string? theme)
    {
        var normalized = (theme ?? "dark").Trim().ToLowerInvariant();
        return normalized is "light" or "dark"
            ? normalized
            : "dark";
    }
}