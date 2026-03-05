using System.Text.Json;
using System.Text.Json.Serialization;
using System.Security.Cryptography;
using System.Text;
using System.ComponentModel;
using System.Runtime.CompilerServices;

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

    public GuestSharedFolderSettings SharedFolders { get; set; } = new();

    public GuestFileServiceSettings FileService { get; set; } = new();

    public GuestUiSettings Ui { get; set; } = new();
}

internal sealed class GuestFileServiceSettings
{
    public bool Enabled { get; set; } = true;

    public string MappingMode { get; set; } = "hypertool-file";

    public bool PreferHyperVSocket { get; set; } = true;
}

internal sealed class GuestUsbSettings
{
    public bool Enabled { get; set; } = true;

    public bool DisconnectOnExit { get; set; } = true;

    public string HostAddress { get; set; } = string.Empty;

    public string HostName { get; set; } = string.Empty;

    public bool HostFeatureEnabled { get; set; } = true;

    public bool UseHyperVSocket { get; set; } = true;

    public string HyperVSocketServiceId { get; set; } = HyperTool.Services.HyperVSocketUsbTunnelDefaults.ServiceIdString;

    public List<string> AutoConnectDeviceKeys { get; set; } = [];
}

internal sealed class GuestSharedFolderSettings
{
    public bool Enabled { get; set; } = true;

    public bool HostFeatureEnabled { get; set; } = true;

    public string BaseDriveLetter { get; set; } = "Z";

    public List<GuestSharedFolderMapping> Mappings { get; set; } = [];
}

internal sealed class GuestSharedFolderMapping : INotifyPropertyChanged
{
    private string _id = Guid.NewGuid().ToString("N");
    private string _label = string.Empty;
    private string _sharePath = string.Empty;
    private string _driveLetter = "Z";
    private bool _persistent = true;
    private bool _enabled = true;
    private string _mountStateDot = "⚪";
    private string _mountStateText = "deaktiviert";

    public event PropertyChangedEventHandler? PropertyChanged;

    public string Id
    {
        get => _id;
        set => SetField(ref _id, value);
    }

    public string Label
    {
        get => _label;
        set => SetField(ref _label, value);
    }

    public string SharePath
    {
        get => _sharePath;
        set => SetField(ref _sharePath, value);
    }

    public string DriveLetter
    {
        get => _driveLetter;
        set => SetField(ref _driveLetter, value);
    }

    public bool Persistent
    {
        get => _persistent;
        set => SetField(ref _persistent, value);
    }

    public bool Enabled
    {
        get => _enabled;
        set => SetField(ref _enabled, value);
    }

    [JsonIgnore]
    public string MountStateDot
    {
        get => _mountStateDot;
        set => SetField(ref _mountStateDot, value);
    }

    [JsonIgnore]
    public string MountStateText
    {
        get => _mountStateText;
        set => SetField(ref _mountStateText, value);
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        return true;
    }
}

internal sealed class GuestUiSettings
{
    public string Theme { get; set; } = "dark";

    public bool StartWithWindows { get; set; } = true;

    public bool StartMinimized { get; set; } = false;

    public bool MinimizeToTray { get; set; } = true;

    public bool CheckForUpdatesOnStartup { get; set; } = true;
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

    public string FileName { get; set; } = "hypertool-guest.log";

    public bool EchoToConsole { get; set; } = true;
}

internal sealed class GuestHandshakeSettings
{
    public bool Enabled { get; set; } = true;

    public string FilePath { get; set; } = string.Empty;
}

internal static class GuestConfigService
{
    private const string ProtectedPasswordPrefix = "dpapi:";
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public static readonly string DefaultConfigPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "HyperTool",
        "HyperTool.Guest.json");

    public static readonly string DefaultLogDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "HyperTool",
        "logs");

    public static readonly string DefaultHandshakePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
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

        var serializableConfig = CreateSerializableConfig(config);
        var json = JsonSerializer.Serialize(serializableConfig, SerializerOptions);
        File.WriteAllText(configPath, json);
    }

    private static void Normalize(GuestConfig config)
    {
        config.SharePath = (config.SharePath ?? string.Empty).Trim();
        config.DriveLetter = NormalizeDriveLetter(config.DriveLetter);
        config.PollIntervalSeconds = Math.Clamp(config.PollIntervalSeconds, 5, 3600);
        config.Credential ??= new GuestCredential();
        config.Credential.Username = (config.Credential.Username ?? string.Empty).Trim();
        config.Credential.Password = UnprotectCredentialPassword(config.Credential.Password);

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
        config.Logging.FileName = NormalizeLogFileName(config.Logging.FileName);

        config.Handshake ??= new GuestHandshakeSettings();
        config.Handshake.FilePath = string.IsNullOrWhiteSpace(config.Handshake.FilePath)
            ? DefaultHandshakePath
            : config.Handshake.FilePath.Trim();

        config.Usb ??= new GuestUsbSettings();
        config.Usb.HostAddress = (config.Usb.HostAddress ?? string.Empty).Trim();
        config.Usb.HostName = (config.Usb.HostName ?? string.Empty).Trim();
        config.Usb.HyperVSocketServiceId = string.IsNullOrWhiteSpace(config.Usb.HyperVSocketServiceId)
            ? HyperTool.Services.HyperVSocketUsbTunnelDefaults.ServiceIdString
            : config.Usb.HyperVSocketServiceId.Trim();
        config.Usb.AutoConnectDeviceKeys = (config.Usb.AutoConnectDeviceKeys ?? [])
            .Where(static key => !string.IsNullOrWhiteSpace(key))
            .Select(static key => key.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        config.SharedFolders ??= new GuestSharedFolderSettings();
        config.SharedFolders.BaseDriveLetter = NormalizeDriveLetter(config.SharedFolders.BaseDriveLetter);
        config.SharedFolders.Mappings = (config.SharedFolders.Mappings ?? [])
            .Where(static mapping => mapping is not null)
            .Select(NormalizeSharedFolderMapping)
            .Where(static mapping => !string.IsNullOrWhiteSpace(mapping.SharePath)
                                     && !string.IsNullOrWhiteSpace(mapping.DriveLetter))
            .GroupBy(static mapping => mapping.Id, StringComparer.OrdinalIgnoreCase)
            .Select(static group => group.First())
            .ToList();

        if (config.SharedFolders.Mappings.Count == 0
            && !string.IsNullOrWhiteSpace(config.SharePath)
            && !string.IsNullOrWhiteSpace(config.DriveLetter))
        {
            config.SharedFolders.Mappings.Add(new GuestSharedFolderMapping
            {
                Id = Guid.NewGuid().ToString("N"),
                Label = "Standard",
                SharePath = config.SharePath.Trim(),
                DriveLetter = NormalizeDriveLetter(config.DriveLetter),
                Persistent = true,
                Enabled = config.Enabled
            });
        }

        config.FileService ??= new GuestFileServiceSettings();
        config.FileService.MappingMode = NormalizeMappingMode(config.FileService.MappingMode);

        config.Ui ??= new GuestUiSettings();
        config.Ui.Theme = NormalizeTheme(config.Ui.Theme);
    }

    private static GuestConfig CreateSerializableConfig(GuestConfig config)
    {
        var cloneJson = JsonSerializer.Serialize(config, SerializerOptions);
        var clone = JsonSerializer.Deserialize<GuestConfig>(cloneJson, SerializerOptions) ?? new GuestConfig();

        clone.Credential ??= new GuestCredential();
        clone.Credential.Password = ProtectCredentialPassword(config.Credential?.Password);

        return clone;
    }

    private static string ProtectCredentialPassword(string? plainPassword)
    {
        var normalized = (plainPassword ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return string.Empty;
        }

        if (normalized.StartsWith(ProtectedPasswordPrefix, StringComparison.OrdinalIgnoreCase))
        {
            return normalized;
        }

        var plainBytes = Encoding.UTF8.GetBytes(normalized);
        var protectedBytes = ProtectedData.Protect(plainBytes, GetCredentialEntropy(), DataProtectionScope.CurrentUser);
        return ProtectedPasswordPrefix + Convert.ToBase64String(protectedBytes);
    }

    private static string UnprotectCredentialPassword(string? storedPassword)
    {
        var normalized = (storedPassword ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return string.Empty;
        }

        if (!normalized.StartsWith(ProtectedPasswordPrefix, StringComparison.OrdinalIgnoreCase))
        {
            return normalized;
        }

        var payload = normalized[ProtectedPasswordPrefix.Length..].Trim();
        if (string.IsNullOrWhiteSpace(payload))
        {
            return string.Empty;
        }

        try
        {
            var encryptedBytes = Convert.FromBase64String(payload);
            var plainBytes = ProtectedData.Unprotect(encryptedBytes, GetCredentialEntropy(), DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(plainBytes);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static byte[] GetCredentialEntropy()
    {
        return Encoding.UTF8.GetBytes("HyperTool.Guest.Credential.v1");
    }

    private static GuestSharedFolderMapping NormalizeSharedFolderMapping(GuestSharedFolderMapping mapping)
    {
        return new GuestSharedFolderMapping
        {
            Id = string.IsNullOrWhiteSpace(mapping.Id) ? Guid.NewGuid().ToString("N") : mapping.Id.Trim(),
            Label = (mapping.Label ?? string.Empty).Trim(),
            SharePath = (mapping.SharePath ?? string.Empty).Trim(),
            DriveLetter = NormalizeDriveLetter(mapping.DriveLetter),
            Persistent = true,
            Enabled = mapping.Enabled
        };
    }

    private static string NormalizeLogFileName(string? fileName)
    {
        var value = (fileName ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(value))
        {
            return "hypertool-guest.log";
        }

        var baseName = Path.GetFileNameWithoutExtension(value);
        if (string.IsNullOrWhiteSpace(baseName))
        {
            baseName = "hypertool-guest";
        }

        return $"{baseName}.log";
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

    public static string NormalizeMappingMode(string? mode)
    {
        return "hypertool-file";
    }
}