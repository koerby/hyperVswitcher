using System.Text.Json;
using System.Text;

namespace HyperTool.Guest;

internal static class GuestLogger
{
    private static readonly object Sync = new();
    private static readonly TimeSpan LogRetentionPeriod = TimeSpan.FromDays(3);
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private static string _logFilePath = string.Empty;
    private static bool _echoToConsole = true;

    public static event Action<string>? EntryWritten;

    public static void Initialize(GuestLoggingSettings settings)
    {
        var directory = string.IsNullOrWhiteSpace(settings.DirectoryPath)
            ? GuestConfigService.DefaultLogDirectory
            : settings.DirectoryPath;

        var fileName = string.IsNullOrWhiteSpace(settings.FileName)
            ? "hypertool-guest.log"
            : settings.FileName;

        Directory.CreateDirectory(directory);
        CleanupOldLogFiles(directory, LogRetentionPeriod);
        _logFilePath = Path.Combine(directory, fileName);
        _echoToConsole = settings.EchoToConsole;

        try
        {
            if (!File.Exists(_logFilePath))
            {
                File.WriteAllText(_logFilePath, string.Empty, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            }
        }
        catch
        {
        }

        try
        {
            Console.OutputEncoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        }
        catch
        {
        }
    }

    public static void Info(string eventName, string message, object? data = null)
        => Write("INFO", eventName, message, data);

    public static void Warn(string eventName, string message, object? data = null)
        => Write("WARN", eventName, message, data);

    public static void Error(string eventName, string message, object? data = null)
        => Write("ERROR", eventName, message, data);

    private static void Write(string level, string eventName, string message, object? data)
    {
        var entry = new Dictionary<string, object?>
        {
            ["timestampUtc"] = DateTime.UtcNow,
            ["level"] = level,
            ["event"] = eventName,
            ["message"] = message,
            ["machine"] = Environment.MachineName,
            ["processId"] = Environment.ProcessId,
            ["data"] = data
        };

        var line = JsonSerializer.Serialize(entry, SerializerOptions);

        lock (Sync)
        {
            if (!string.IsNullOrWhiteSpace(_logFilePath))
            {
                File.AppendAllText(_logFilePath, line + Environment.NewLine, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            }

            if (_echoToConsole)
            {
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {level} {eventName}: {message}");
            }

            EntryWritten?.Invoke($"[{DateTime.Now:HH:mm:ss}] [{level}] {message}");
        }
    }

    private static void CleanupOldLogFiles(string directory, TimeSpan maxAge)
    {
        try
        {
            if (!Directory.Exists(directory))
            {
                return;
            }

            var cutoffUtc = DateTime.UtcNow - maxAge;
            foreach (var file in Directory.EnumerateFiles(directory, "*", SearchOption.TopDirectoryOnly))
            {
                try
                {
                    if (File.GetLastWriteTimeUtc(file) < cutoffUtc)
                    {
                        File.Delete(file);
                    }
                }
                catch
                {
                }
            }
        }
        catch
        {
        }
    }
}

internal sealed class GuestHandshakeState
{
    public DateTime TimestampUtc { get; set; } = DateTime.UtcNow;

    public string GuestMachine { get; set; } = Environment.MachineName;

    public string GuestUser { get; set; } = Environment.UserName;

    public string AgentVersion { get; set; } = typeof(GuestHandshakeState).Assembly.GetName().Version?.ToString() ?? "0.0.0";

    public string Command { get; set; } = string.Empty;

    public string ConfigPath { get; set; } = string.Empty;

    public string SharePath { get; set; } = string.Empty;

    public string DriveLetter { get; set; } = string.Empty;

    public bool IsMapped { get; set; }

    public string MappedRemotePath { get; set; } = string.Empty;

    public int LastExitCode { get; set; }

    public string LastError { get; set; } = string.Empty;

    public string PreferredAutostartMode { get; set; } = "run-registry";

    public bool RunRegistryInstalled { get; set; }

    public bool TaskInstalled { get; set; }
}

internal static class GuestHandshakeWriter
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    public static void TryWrite(GuestConfig config, GuestHandshakeState state)
    {
        if (!config.Handshake.Enabled)
        {
            return;
        }

        var path = string.IsNullOrWhiteSpace(config.Handshake.FilePath)
            ? GuestConfigService.DefaultHandshakePath
            : config.Handshake.FilePath;

        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var tempPath = path + ".tmp";
        var json = JsonSerializer.Serialize(state, SerializerOptions);
        File.WriteAllText(tempPath, json);
        File.Copy(tempPath, path, overwrite: true);
        File.Delete(tempPath);
    }
}