using HyperTool.Models;
using System.ComponentModel;
using System.Diagnostics;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace HyperTool.Services;

public sealed class UsbIpdCliService : IUsbIpService
{
    private static readonly Regex HardwareIdRegex = new("VID_([0-9A-Fa-f]{4})&PID_([0-9A-Fa-f]{4})", RegexOptions.Compiled);
    private static readonly Regex ServiceStateRegex = new(@"STATE\s*:\s*(\d+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex RemoteUsbListRegex = new(@"^\s*([0-9]+-[0-9]+(?:\.[0-9]+)*)\s*:\s*(.+?)\s*\(([0-9A-Fa-f]{4}:[0-9A-Fa-f]{4})\)\s*$", RegexOptions.Compiled);
    private static readonly Regex UsbipPortHeaderRegex = new(@"^\s*Port\s+(\d+):", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex UsbipPortBusIdRegex = new(@"\bbusid\s+([0-9]+-[0-9]+(?:\.[0-9]+)*)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex UsbipPortUriBusIdRegex = new(@"usbip://[^\s/]+:\d+/([0-9]+-[0-9]+(?:\.[0-9]+)*)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private const string WslUsbipFallbackPath = "/mnt/c/Program Files/usbipd-win/WSL/usbip";
    private const string NativeUsbipFallbackPath = @"C:\Program Files\USBip\usbip.exe";
    private static bool _autoInstallAttempted;
    private static readonly SemaphoreSlim EnsureReadyGate = new(1, 1);

    public async Task<bool> IsUsbClientAvailableAsync(CancellationToken cancellationToken)
    {
        foreach (var candidate in GetNativeUsbipCandidates().Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (string.Equals(candidate, "usbip", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (File.Exists(candidate))
            {
                return true;
            }
        }

        var whereResult = await RunUtilityAsync("where", "usbip", cancellationToken);
        return whereResult.ExitCode == 0 && !string.IsNullOrWhiteSpace(whereResult.StandardOutput);
    }

    public async Task<IReadOnlyList<UsbIpDeviceInfo>> GetDevicesAsync(CancellationToken cancellationToken)
    {
        await EnsureReadyAsync(cancellationToken);
        var result = await RunCommandAsync(["state"], cancellationToken);
        EnsureSuccess(result, "USB-Status konnte nicht geladen werden.");
        return ParseState(result.StandardOutput);
    }

    public async Task<IReadOnlyList<UsbIpDeviceInfo>> GetRemoteDevicesAsync(string hostAddress, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(hostAddress))
        {
            throw new ArgumentException("Host-Adresse darf nicht leer sein.", nameof(hostAddress));
        }

        var trimmedHost = hostAddress.Trim();
        var result = await RunNativeUsbipCommandWithFallbackAsync([
            "list",
            "-r",
            trimmedHost
        ], cancellationToken);

        EnsureSuccess(result, $"Remote-USB-Liste vom Host '{hostAddress}' konnte nicht geladen werden. Stelle sicher, dass usbip-win2 installiert ist und 'usbip.exe' verfügbar ist.");

        var devices = ParseRemoteUsbList(result.StandardOutput).ToList();

        try
        {
            var portResult = await RunNativeUsbipCommandWithFallbackAsync(["port"], cancellationToken);
            if (portResult.ExitCode == 0)
            {
                var attachedBusIds = ParseAttachedBusIdsFromPortOutput(portResult.StandardOutput);
                if (attachedBusIds.Count > 0)
                {
                    foreach (var device in devices)
                    {
                        if (attachedBusIds.Contains(device.BusId))
                        {
                            device.ClientIpAddress = hostAddress.Trim();
                        }
                    }
                }
            }
        }
        catch
        {
        }

        return devices;
    }

    public async Task AttachToHostAsync(string busId, string hostAddress, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(busId))
        {
            throw new ArgumentException("BUSID darf nicht leer sein.", nameof(busId));
        }

        if (string.IsNullOrWhiteSpace(hostAddress))
        {
            throw new ArgumentException("Host-Adresse darf nicht leer sein.", nameof(hostAddress));
        }

        var result = await RunNativeUsbipCommandWithFallbackAsync([
            "attach",
            "-r",
            hostAddress.Trim(),
            "-b",
            busId.Trim()
        ], cancellationToken);

        EnsureSuccess(result, $"USB-Gerät mit BUSID '{busId}' konnte nicht vom Host '{hostAddress}' angehängt werden.");
    }

    public async Task DetachFromHostAsync(string busId, string? hostAddress, CancellationToken cancellationToken)
    {
        try
        {
            var normalizedBusId = (busId ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalizedBusId))
            {
                throw new ArgumentException("BUSID darf nicht leer sein.", nameof(busId));
            }

            var portResult = await RunNativeUsbipCommandWithFallbackAsync(["port"], cancellationToken);

            EnsureSuccess(portResult, "USB-Ports im Guest konnten nicht geladen werden.");

            var mappedPort = FindPortForBusId(portResult.StandardOutput, normalizedBusId);
            if (mappedPort is null)
            {
                return;
            }

            var detachResult = await RunNativeUsbipCommandWithFallbackAsync(["detach", "-p", mappedPort.Value.ToString()], cancellationToken);

            EnsureSuccess(detachResult, $"USB-Gerät mit BUSID '{normalizedBusId}' konnte nicht getrennt werden.");
        }
        catch (Exception ex)
        {
            var hostSuffix = string.IsNullOrWhiteSpace(hostAddress)
                ? string.Empty
                : $" (Host '{hostAddress}')";

            throw new InvalidOperationException(
                $"USB-Gerät mit BUSID '{busId}' konnte nicht vom Host getrennt werden{hostSuffix}. {ex.Message}", ex);
        }
    }

    public async Task BindAsync(string busId, bool force, CancellationToken cancellationToken)
    {
        await EnsureReadyAsync(cancellationToken);
        var args = new List<string> { "bind", "--busid", busId };
        if (force)
        {
            args.Add("--force");
        }

        if (!IsProcessElevated())
        {
            var exitCode = await RunCommandElevatedAsync(args, cancellationToken);
            EnsureSuccess(exitCode, $"USB-Gerät mit BUSID '{busId}' konnte nicht freigegeben werden.");
            return;
        }

        var result = await RunCommandAsync(args, cancellationToken);
        EnsureSuccess(result, $"USB-Gerät mit BUSID '{busId}' konnte nicht freigegeben werden.");
    }

    public async Task UnbindAsync(string busId, CancellationToken cancellationToken)
    {
        await EnsureReadyAsync(cancellationToken);
        if (!IsProcessElevated())
        {
            var exitCode = await RunCommandElevatedAsync(["unbind", "--busid", busId], cancellationToken);
            EnsureSuccess(exitCode, $"USB-Freigabe für BUSID '{busId}' konnte nicht entfernt werden.");
            return;
        }

        var result = await RunCommandAsync(["unbind", "--busid", busId], cancellationToken);
        EnsureSuccess(result, $"USB-Freigabe für BUSID '{busId}' konnte nicht entfernt werden.");
    }

    public async Task UnbindByPersistedGuidAsync(string persistedGuid, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(persistedGuid))
        {
            throw new ArgumentException("Persisted GUID darf nicht leer sein.", nameof(persistedGuid));
        }

        await EnsureReadyAsync(cancellationToken);
        if (!IsProcessElevated())
        {
            var exitCode = await RunCommandElevatedAsync(["unbind", "--guid", persistedGuid], cancellationToken);
            EnsureSuccess(exitCode, $"USB-Freigabe für GUID '{persistedGuid}' konnte nicht entfernt werden.");
            return;
        }

        var result = await RunCommandAsync(["unbind", "--guid", persistedGuid], cancellationToken);
        EnsureSuccess(result, $"USB-Freigabe für GUID '{persistedGuid}' konnte nicht entfernt werden.");
    }

    public async Task AttachToWslAsync(string busId, string? distribution, CancellationToken cancellationToken)
    {
        await EnsureReadyAsync(cancellationToken);
        var args = new List<string>
        {
            "attach",
            "--wsl",
            "--busid",
            busId
        };

        if (!string.IsNullOrWhiteSpace(distribution))
        {
            args.Add(distribution.Trim());
        }

        var result = await RunCommandAsync(args, cancellationToken);
        EnsureSuccess(result, $"USB-Gerät mit BUSID '{busId}' konnte nicht an WSL angehängt werden.");
    }

    public async Task DetachAsync(string busId, CancellationToken cancellationToken)
    {
        await EnsureReadyAsync(cancellationToken);
        if (!IsProcessElevated())
        {
            var exitCode = await RunCommandElevatedAsync(["detach", "--busid", busId], cancellationToken);
            EnsureSuccess(exitCode, $"USB-Gerät mit BUSID '{busId}' konnte nicht getrennt werden.");
            return;
        }

        var result = await RunCommandAsync(["detach", "--busid", busId], cancellationToken);
        EnsureSuccess(result, $"USB-Gerät mit BUSID '{busId}' konnte nicht getrennt werden.");
    }

    private static IReadOnlyList<UsbIpDeviceInfo> ParseState(string json)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            return [];
        }

        using var document = JsonDocument.Parse(json);
        if (!document.RootElement.TryGetProperty("Devices", out var devicesElement)
            || devicesElement.ValueKind != JsonValueKind.Array)
        {
            return [];
        }

        var devices = new List<UsbIpDeviceInfo>();
        foreach (var deviceElement in devicesElement.EnumerateArray())
        {
            var instanceId = GetString(deviceElement, "InstanceId");
            var hardwareId = TryExtractHardwareId(instanceId);

            devices.Add(new UsbIpDeviceInfo
            {
                BusId = GetString(deviceElement, "BusId"),
                Description = GetString(deviceElement, "Description"),
                HardwareId = hardwareId,
                InstanceId = instanceId,
                PersistedGuid = GetString(deviceElement, "PersistedGuid"),
                ClientIpAddress = GetString(deviceElement, "ClientIPAddress")
            });
        }

        return devices
            .OrderBy(device => string.IsNullOrWhiteSpace(device.BusId) ? 1 : 0)
            .ThenBy(device => device.BusId, StringComparer.OrdinalIgnoreCase)
            .ThenBy(device => device.Description, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string GetString(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var propertyElement))
        {
            return string.Empty;
        }

        return propertyElement.ValueKind switch
        {
            JsonValueKind.String => propertyElement.GetString() ?? string.Empty,
            JsonValueKind.Number => propertyElement.ToString(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            _ => string.Empty
        };
    }

    private static string TryExtractHardwareId(string instanceId)
    {
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return string.Empty;
        }

        var match = HardwareIdRegex.Match(instanceId);
        if (!match.Success)
        {
            return string.Empty;
        }

        return $"{match.Groups[1].Value}:{match.Groups[2].Value}".ToUpperInvariant();
    }

    private static IReadOnlyList<UsbIpDeviceInfo> ParseRemoteUsbList(string output)
    {
        if (string.IsNullOrWhiteSpace(output))
        {
            return [];
        }

        var devices = new List<UsbIpDeviceInfo>();
        var lines = output.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
        foreach (var line in lines)
        {
            var match = RemoteUsbListRegex.Match(line);
            if (!match.Success)
            {
                continue;
            }

            var busId = match.Groups[1].Value.Trim();
            var description = match.Groups[2].Value.Trim();
            var hardwareId = match.Groups[3].Value.Trim().ToUpperInvariant();

            devices.Add(new UsbIpDeviceInfo
            {
                BusId = busId,
                Description = description,
                HardwareId = hardwareId,
                InstanceId = string.Empty,
                PersistedGuid = string.Empty,
                ClientIpAddress = string.Empty
            });
        }

        return devices
            .OrderBy(device => device.BusId, StringComparer.OrdinalIgnoreCase)
            .ThenBy(device => device.Description, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static int? FindPortForBusId(string output, string busId)
    {
        if (string.IsNullOrWhiteSpace(output) || string.IsNullOrWhiteSpace(busId))
        {
            return null;
        }

        int? currentPort = null;
        var lines = output.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
        foreach (var line in lines)
        {
            var headerMatch = UsbipPortHeaderRegex.Match(line);
            if (headerMatch.Success && int.TryParse(headerMatch.Groups[1].Value, out var parsedPort))
            {
                currentPort = parsedPort;
                continue;
            }

            var busIdMatch = UsbipPortBusIdRegex.Match(line);
            if (busIdMatch.Success
                && currentPort.HasValue
                && string.Equals(busIdMatch.Groups[1].Value.Trim(), busId, StringComparison.OrdinalIgnoreCase))
            {
                return currentPort.Value;
            }

            var uriBusIdMatch = UsbipPortUriBusIdRegex.Match(line);
            if (uriBusIdMatch.Success
                && currentPort.HasValue
                && string.Equals(uriBusIdMatch.Groups[1].Value.Trim(), busId, StringComparison.OrdinalIgnoreCase))
            {
                return currentPort.Value;
            }
        }

        return null;
    }

    private static HashSet<string> ParseAttachedBusIdsFromPortOutput(string output)
    {
        var busIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(output))
        {
            return busIds;
        }

        var lines = output.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
        foreach (var line in lines)
        {
            var busIdMatch = UsbipPortBusIdRegex.Match(line);
            if (busIdMatch.Success)
            {
                var busId = busIdMatch.Groups[1].Value.Trim();
                if (!string.IsNullOrWhiteSpace(busId))
                {
                    busIds.Add(busId);
                }
            }

            var uriBusIdMatch = UsbipPortUriBusIdRegex.Match(line);
            if (uriBusIdMatch.Success)
            {
                var busId = uriBusIdMatch.Groups[1].Value.Trim();
                if (!string.IsNullOrWhiteSpace(busId))
                {
                    busIds.Add(busId);
                }
            }
        }

        return busIds;
    }

    private static string ResolveExecutablePath()
    {
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var candidate = Path.Combine(programFiles, "usbipd-win", "usbipd.exe");
        if (File.Exists(candidate))
        {
            return candidate;
        }

        return "usbipd";
    }

    private static async Task EnsureReadyAsync(CancellationToken cancellationToken)
    {
        await EnsureReadyGate.WaitAsync(cancellationToken);
        try
        {
            if (!await IsUsbipdAvailableAsync(cancellationToken))
            {
                await TryAutoInstallUsbipdAsync(cancellationToken);
            }

            if (!await IsUsbipdAvailableAsync(cancellationToken))
            {
                throw new InvalidOperationException(
                    "usbipd-win ist nicht installiert. Bitte usbipd-win installieren oder Winget aktivieren.");
            }

            await EnsureUsbipdServiceRunningAsync(cancellationToken);
        }
        finally
        {
            EnsureReadyGate.Release();
        }
    }

    private static async Task<bool> IsUsbipdAvailableAsync(CancellationToken cancellationToken)
    {
        var resolvedPath = ResolveExecutablePath();
        if (!string.Equals(resolvedPath, "usbipd", StringComparison.OrdinalIgnoreCase))
        {
            return File.Exists(resolvedPath);
        }

        var whereResult = await RunUtilityAsync("where", "usbipd", cancellationToken);
        return whereResult.ExitCode == 0 && !string.IsNullOrWhiteSpace(whereResult.StandardOutput);
    }

    private static async Task TryAutoInstallUsbipdAsync(CancellationToken cancellationToken)
    {
        if (_autoInstallAttempted)
        {
            return;
        }

        _autoInstallAttempted = true;

        var wingetCheck = await RunUtilityAsync("where", "winget", cancellationToken);
        if (wingetCheck.ExitCode != 0)
        {
            return;
        }

        var installArgs = "install --id dorssel.usbipd-win --exact --silent --accept-source-agreements --accept-package-agreements";

        if (IsProcessElevated())
        {
            var installResult = await RunUtilityAsync("winget", installArgs, cancellationToken);
            if (installResult.ExitCode != 0)
            {
                throw new InvalidOperationException(
                    "usbipd-win konnte nicht automatisch installiert werden. Bitte Installation manuell ausführen.");
            }

            return;
        }

        try
        {
            var exitCode = await RunUtilityElevatedAsync("winget", installArgs, cancellationToken);
            if (exitCode != 0)
            {
                throw new InvalidOperationException(
                    "usbipd-win konnte nicht automatisch installiert werden. Bitte Installation manuell ausführen.");
            }
        }
        catch (InvalidOperationException)
        {
            throw;
        }
    }

    private static async Task EnsureUsbipdServiceRunningAsync(CancellationToken cancellationToken)
    {
        var queryResult = await RunUtilityAsync("sc", "query usbipd", cancellationToken);
        if (queryResult.ExitCode != 0)
        {
            throw new InvalidOperationException("usbipd-Dienst ist nicht verfügbar.");
        }

        var match = ServiceStateRegex.Match(queryResult.StandardOutput + "\n" + queryResult.StandardError);
        if (!match.Success)
        {
            return;
        }

        if (!int.TryParse(match.Groups[1].Value, out var stateCode))
        {
            return;
        }

        if (stateCode == 4)
        {
            return;
        }

        if (IsProcessElevated())
        {
            var startResult = await RunUtilityAsync("sc", "start usbipd", cancellationToken);
            if (startResult.ExitCode != 0)
            {
                throw new InvalidOperationException("usbipd-Dienst konnte nicht gestartet werden.");
            }
        }
        else
        {
            var exitCode = await RunUtilityElevatedAsync("sc", "start usbipd", cancellationToken);
            if (exitCode != 0)
            {
                throw new InvalidOperationException("usbipd-Dienst konnte nicht gestartet werden.");
            }
        }
    }

    private static async Task<CommandResult> RunUtilityAsync(string fileName, string arguments, CancellationToken cancellationToken)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            CreateNoWindow = true
        };

        Process process;
        try
        {
            process = Process.Start(startInfo)
                      ?? throw new InvalidOperationException($"{fileName} konnte nicht gestartet werden.");
        }
        catch (Win32Exception)
        {
            return new CommandResult(1, string.Empty, string.Empty);
        }

        using (process)
        {
            using var registration = cancellationToken.Register(() =>
            {
                try
                {
                    if (!process.HasExited)
                    {
                        process.Kill(true);
                    }
                }
                catch
                {
                }
            });

            var stdoutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
            var stderrTask = process.StandardError.ReadToEndAsync(cancellationToken);
            await process.WaitForExitAsync(cancellationToken);

            return new CommandResult(process.ExitCode, await stdoutTask, await stderrTask);
        }
    }

    private static async Task<int> RunUtilityElevatedAsync(string fileName, string arguments, CancellationToken cancellationToken)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = fileName,
            UseShellExecute = true,
            Verb = "runas",
            Arguments = arguments,
            CreateNoWindow = false
        };

        Process process;
        try
        {
            process = Process.Start(startInfo)
                      ?? throw new InvalidOperationException($"{fileName} konnte nicht mit Administratorrechten gestartet werden.");
        }
        catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
        {
            throw new InvalidOperationException("UAC-Bestätigung abgebrochen. Aktion wurde nicht ausgeführt.", ex);
        }

        using (process)
        {
            using var registration = cancellationToken.Register(() =>
            {
                try
                {
                    if (!process.HasExited)
                    {
                        process.Kill(true);
                    }
                }
                catch
                {
                }
            });

            await process.WaitForExitAsync(cancellationToken);
            return process.ExitCode;
        }
    }

    private static async Task<CommandResult> RunCommandAsync(IReadOnlyList<string> args, CancellationToken cancellationToken)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = ResolveExecutablePath(),
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            CreateNoWindow = true
        };

        foreach (var arg in args)
        {
            startInfo.ArgumentList.Add(arg);
        }

        Process process;
        try
        {
            process = Process.Start(startInfo)
                      ?? throw new InvalidOperationException("usbipd konnte nicht gestartet werden.");
        }
        catch (Win32Exception ex)
        {
            throw new InvalidOperationException(
                "usbipd.exe wurde nicht gefunden. Bitte usbipd-win installieren und erneut versuchen.", ex);
        }

        using (process)
        {
            using var registration = cancellationToken.Register(() =>
            {
                try
                {
                    if (!process.HasExited)
                    {
                        process.Kill(true);
                    }
                }
                catch
                {
                }
            });

            var stdoutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
            var stderrTask = process.StandardError.ReadToEndAsync(cancellationToken);
            await process.WaitForExitAsync(cancellationToken);

            var stdout = await stdoutTask;
            var stderr = await stderrTask;

            return new CommandResult(process.ExitCode, stdout, stderr);
        }
    }

    private static async Task<CommandResult> RunWslCommandAsync(IReadOnlyList<string> args, CancellationToken cancellationToken)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "wsl",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        foreach (var arg in args)
        {
            startInfo.ArgumentList.Add(arg);
        }

        Process process;
        try
        {
            process = Process.Start(startInfo)
                      ?? throw new InvalidOperationException("WSL konnte nicht gestartet werden.");
        }
        catch (Win32Exception ex)
        {
            throw new InvalidOperationException("WSL ist nicht verfügbar. Bitte WSL installieren und erneut versuchen.", ex);
        }

        using (process)
        {
            using var registration = cancellationToken.Register(() =>
            {
                try
                {
                    if (!process.HasExited)
                    {
                        process.Kill(true);
                    }
                }
                catch
                {
                }
            });

            var stdoutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
            var stderrTask = process.StandardError.ReadToEndAsync(cancellationToken);
            await process.WaitForExitAsync(cancellationToken);

            return new CommandResult(process.ExitCode, await stdoutTask, await stderrTask);
        }
    }

    private static async Task<CommandResult> RunUsbipClientCommandWithFallbackAsync(IReadOnlyList<string> usbipArgs, CancellationToken cancellationToken)
    {
        var primaryArgs = new List<string> { "--exec", "usbip" };
        primaryArgs.AddRange(usbipArgs);

        var result = await RunWslCommandAsync(primaryArgs, cancellationToken);

        if (ShouldRetryWithExecCompatibilityFallback(result))
        {
            var execFallbackArgs = new List<string> { "-e", "usbip" };
            execFallbackArgs.AddRange(usbipArgs);
            result = await RunWslCommandAsync(execFallbackArgs, cancellationToken);
        }

        if (ShouldRetryWithFallbackUsbip(result))
        {
            var fallbackArgs = new List<string> { "--exec", WslUsbipFallbackPath };
            fallbackArgs.AddRange(usbipArgs);
            result = await RunWslCommandAsync(fallbackArgs, cancellationToken);

            if (ShouldRetryWithExecCompatibilityFallback(result))
            {
                var legacyFallbackArgs = new List<string> { "-e", WslUsbipFallbackPath };
                legacyFallbackArgs.AddRange(usbipArgs);
                result = await RunWslCommandAsync(legacyFallbackArgs, cancellationToken);
            }
        }

        return result;
    }

    private static async Task<CommandResult> RunNativeUsbipCommandWithFallbackAsync(IReadOnlyList<string> usbipArgs, CancellationToken cancellationToken)
    {
        CommandResult? lastResult = null;
        foreach (var candidate in GetNativeUsbipCandidates())
        {
            var result = await RunNativeUsbipCommandAsync(candidate, usbipArgs, cancellationToken);
            lastResult = result;

            if (result.ExitCode == 0)
            {
                return result;
            }

            if (!ShouldRetryWithNativeFallback(result))
            {
                return result;
            }
        }

        return lastResult ?? new CommandResult(1, string.Empty, "'usbip.exe' not found");
    }

    private static IEnumerable<string> GetNativeUsbipCandidates()
    {
        yield return Path.Combine(AppContext.BaseDirectory, "usbip.exe");
        yield return NativeUsbipFallbackPath;
        yield return "usbip";
    }

    private static bool ShouldRetryWithNativeFallback(CommandResult result)
    {
        var text = ((result.StandardError ?? string.Empty) + "\n" + (result.StandardOutput ?? string.Empty)).ToLowerInvariant();
        return text.Contains("not recognized")
               || text.Contains("wurde nicht als name eines cmdlet")
               || text.Contains("konnte nicht gefunden werden")
               || text.Contains("file not found")
               || text.Contains("no such file")
               || text.Contains("'usbip' not found")
               || text.Contains("'usbip.exe' not found")
               || text.Contains("argument was not expected: --once")
               || text.Contains("unknown option --once")
               || text.Contains("unrecognized option '--once'")
               || text.Contains("invalid option --once")
               || text.Contains("the following argument was not expected")
               || text.Contains("not found");
    }

    private static async Task<CommandResult> RunNativeUsbipCommandAsync(string fileName, IReadOnlyList<string> args, CancellationToken cancellationToken)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = fileName,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        foreach (var arg in args)
        {
            startInfo.ArgumentList.Add(arg);
        }

        Process process;
        try
        {
            process = Process.Start(startInfo)
                      ?? throw new InvalidOperationException("usbip konnte nicht gestartet werden.");
        }
        catch (Win32Exception)
        {
            return new CommandResult(1, string.Empty, $"'{fileName}' not found");
        }

        using (process)
        {
            using var registration = cancellationToken.Register(() =>
            {
                try
                {
                    if (!process.HasExited)
                    {
                        process.Kill(true);
                    }
                }
                catch
                {
                }
            });

            var stdoutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
            var stderrTask = process.StandardError.ReadToEndAsync(cancellationToken);
            await process.WaitForExitAsync(cancellationToken);

            return new CommandResult(process.ExitCode, await stdoutTask, await stderrTask);
        }
    }

    private static bool ShouldRetryWithExecCompatibilityFallback(CommandResult result)
    {
        if (result.ExitCode == 0)
        {
            return false;
        }

        var text = ((result.StandardError ?? string.Empty) + "\n" + (result.StandardOutput ?? string.Empty)).ToLowerInvariant();
        return text.Contains("nutzung: wsl")
               || text.Contains("usage: wsl")
               || text.Contains("keine installierten komponenten gefunden")
               || text.Contains("no installed distributions")
               || text.Contains("wsl --install")
               || text.Contains("wsl.exe [argumente]");
    }

    private static string NormalizeWslText(string value)
        => string.IsNullOrEmpty(value)
            ? string.Empty
            : value.Replace("\0", string.Empty);

    private static bool ShouldRetryWithFallbackUsbip(CommandResult result)
    {
        if (result.ExitCode == 0)
        {
            return false;
        }

        var text = (result.StandardError ?? string.Empty) + "\n" + (result.StandardOutput ?? string.Empty);
        return text.IndexOf("execvpe(usbip) failed", StringComparison.OrdinalIgnoreCase) >= 0
               || text.IndexOf("usbip: not found", StringComparison.OrdinalIgnoreCase) >= 0
               || text.IndexOf("No such file or directory", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static async Task<int> RunCommandElevatedAsync(IReadOnlyList<string> args, CancellationToken cancellationToken)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = ResolveExecutablePath(),
            UseShellExecute = true,
            Verb = "runas",
            Arguments = BuildArguments(args),
            CreateNoWindow = false
        };

        Process process;
        try
        {
            process = Process.Start(startInfo)
                      ?? throw new InvalidOperationException("usbipd konnte nicht mit Administratorrechten gestartet werden.");
        }
        catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
        {
            throw new InvalidOperationException("UAC-Bestätigung abgebrochen. Aktion wurde nicht ausgeführt.", ex);
        }
        catch (Win32Exception ex)
        {
            throw new InvalidOperationException(
                "usbipd.exe wurde nicht gefunden. Bitte usbipd-win installieren und erneut versuchen.", ex);
        }

        using (process)
        {
            using var registration = cancellationToken.Register(() =>
            {
                try
                {
                    if (!process.HasExited)
                    {
                        process.Kill(true);
                    }
                }
                catch
                {
                }
            });

            await process.WaitForExitAsync(cancellationToken);
            return process.ExitCode;
        }
    }

    private static string BuildArguments(IReadOnlyList<string> args)
    {
        return string.Join(" ", args.Select(QuoteArgument));
    }

    private static string QuoteArgument(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return "\"\"";
        }

        if (!value.Any(ch => char.IsWhiteSpace(ch) || ch == '"'))
        {
            return value;
        }

        return "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
    }

    private static bool IsProcessElevated()
    {
        using var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }

    private static void EnsureSuccess(CommandResult result, string fallbackMessage)
    {
        if (result.ExitCode == 0)
        {
            return;
        }

        var stderr = NormalizeWslText(result.StandardError ?? string.Empty).Trim();
        var stdout = NormalizeWslText(result.StandardOutput ?? string.Empty).Trim();

        var details = string.IsNullOrWhiteSpace(stderr)
            ? stdout
            : stderr;

        if (string.IsNullOrWhiteSpace(details))
        {
            throw new InvalidOperationException(fallbackMessage);
        }

        throw new InvalidOperationException($"{fallbackMessage} {details}");
    }

    private static void EnsureSuccess(int exitCode, string fallbackMessage)
    {
        if (exitCode == 0)
        {
            return;
        }

        throw new InvalidOperationException($"{fallbackMessage} ExitCode={exitCode}.");
    }

    private readonly record struct CommandResult(int ExitCode, string StandardOutput, string StandardError);
}
