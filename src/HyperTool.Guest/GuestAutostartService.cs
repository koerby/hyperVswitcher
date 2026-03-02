using System.Diagnostics;
using System.Reflection;

namespace HyperTool.Guest;

internal sealed class GuestAutostartStatus
{
    public bool RunRegistryInstalled { get; set; }

    public bool TaskInstalled { get; set; }
}

internal static class GuestAutostartService
{
    private const string RunRegistryKey = "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run";

    public static async Task<int> InstallAsync(GuestConfig config, string configPath, string mode)
    {
        var normalizedMode = GuestConfigService.NormalizeAutostartMode(mode);
        var startupCommand = BuildStartupCommand(configPath);

        if (normalizedMode == "task")
        {
            var escaped = startupCommand.Replace("\"", "\"\"");
            var code = await RunProcessAsync(
                "schtasks",
                $"/Create /TN \"{config.Autostart.TaskName}\" /SC ONLOGON /TR \"{escaped}\" /RL LIMITED /F");

            return code;
        }

        return await RunProcessAsync(
            "reg",
            $"add \"{RunRegistryKey}\" /v \"{config.Autostart.RunValueName}\" /t REG_SZ /d \"{startupCommand}\" /f");
    }

    public static async Task<int> RemoveAsync(GuestConfig config)
    {
        var runInstalled = await RunProcessAsync(
            "reg",
            $"query \"{RunRegistryKey}\" /v \"{config.Autostart.RunValueName}\"") == 0;

        var taskInstalled = await RunProcessAsync(
            "schtasks",
            $"/Query /TN \"{config.Autostart.TaskName}\"") == 0;

        var runCode = runInstalled
            ? await RunProcessAsync(
                "reg",
                $"delete \"{RunRegistryKey}\" /v \"{config.Autostart.RunValueName}\" /f")
            : 0;

        var taskCode = taskInstalled
            ? await RunProcessAsync(
                "schtasks",
                $"/Delete /TN \"{config.Autostart.TaskName}\" /F")
            : 0;

        return runCode == 0 || taskCode == 0 ? 0 : 1;
    }

    public static async Task<GuestAutostartStatus> QueryStatusAsync(GuestConfig config)
    {
        var runCode = await RunProcessAsync(
            "reg",
            $"query \"{RunRegistryKey}\" /v \"{config.Autostart.RunValueName}\"");

        var taskCode = await RunProcessAsync(
            "schtasks",
            $"/Query /TN \"{config.Autostart.TaskName}\"");

        return new GuestAutostartStatus
        {
            RunRegistryInstalled = runCode == 0,
            TaskInstalled = taskCode == 0
        };
    }

    public static string BuildStartupCommand(string configPath)
    {
        var processPath = Environment.ProcessPath ?? string.Empty;
        var quotedConfig = Quote(configPath);

        if (!string.IsNullOrWhiteSpace(processPath)
            && string.Equals(Path.GetFileName(processPath), "dotnet.exe", StringComparison.OrdinalIgnoreCase))
        {
            var entryAssemblyPath = Assembly.GetEntryAssembly()?.Location ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(entryAssemblyPath))
            {
                return $"{Quote(processPath)} {Quote(entryAssemblyPath)} run --config {quotedConfig}";
            }
        }

        return $"{Quote(processPath)} run --config {quotedConfig}";
    }

    private static string Quote(string value)
        => $"\"{value}\"";

    private static async Task<int> RunProcessAsync(string fileName, string arguments)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using var process = new Process { StartInfo = startInfo };
        process.Start();
        _ = process.StandardOutput.ReadToEndAsync();
        _ = process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();
        return process.ExitCode;
    }
}