using System.Diagnostics;

namespace HyperTool.Guest;

internal static class GuestCli
{
    public static async Task<int> ExecuteAsync(string[] args)
    {
        var configPath = ResolveConfigPath(args);
        var command = ResolveCommand(args);

        GuestConfig config;
        try
        {
            config = GuestConfigService.LoadOrCreate(configPath, out var created);

            if (created)
            {
                GuestLogger.Initialize(new GuestLoggingSettings
                {
                    DirectoryPath = GuestConfigService.DefaultLogDirectory,
                    FileName = "hypertool-guest.ndjson",
                    EchoToConsole = true
                });

                GuestLogger.Info("config.created", "Konfiguration erstellt.", new { configPath });
                GuestLogger.Info("config.hint", "Bitte SharePath/DriveLetter anpassen und erneut starten.");
                return 0;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Konfiguration konnte nicht geladen werden: {ex.Message}");
            return 1;
        }

        GuestLogger.Initialize(config.Logging);
        GuestLogger.Info("startup", "HyperTool.Guest gestartet.", new { command, configPath });

        var driveLetter = GuestConfigService.NormalizeDriveLetter(config.DriveLetter);
        var handshakeState = new GuestHandshakeState
        {
            Command = command,
            ConfigPath = configPath,
            SharePath = config.SharePath,
            DriveLetter = driveLetter,
            PreferredAutostartMode = config.Autostart.PreferredMode
        };

        try
        {
            var result = await ExecuteCommandAsync(command, args, config, configPath, driveLetter, handshakeState);
            handshakeState.LastExitCode = result;
            handshakeState.LastError = result == 0 ? string.Empty : "Befehl endete mit Fehlercode.";
            await UpdateHandshakeStateAsync(config, driveLetter, handshakeState);
            GuestHandshakeWriter.TryWrite(config, handshakeState);
            return result;
        }
        catch (Exception ex)
        {
            GuestLogger.Error("runtime.exception", ex.Message, new { command, ex.StackTrace });
            handshakeState.LastExitCode = 1;
            handshakeState.LastError = ex.Message;
            await UpdateHandshakeStateAsync(config, driveLetter, handshakeState);
            GuestHandshakeWriter.TryWrite(config, handshakeState);
            return 1;
        }
    }

    private static string ResolveCommand(string[] args)
    {
        var positional = new List<string>();
        for (var index = 0; index < args.Length; index++)
        {
            if (string.Equals(args[index], "--config", StringComparison.OrdinalIgnoreCase))
            {
                index++;
                continue;
            }

            positional.Add(args[index]);
        }

        return positional.FirstOrDefault()?.Trim().ToLowerInvariant() ?? "status";
    }

    private static string ResolveConfigPath(string[] args)
    {
        var index = Array.FindIndex(args, item => string.Equals(item, "--config", StringComparison.OrdinalIgnoreCase));
        if (index >= 0 && index + 1 < args.Length)
        {
            return Path.GetFullPath(args[index + 1]);
        }

        return GuestConfigService.DefaultConfigPath;
    }

    private static async Task<int> ExecuteCommandAsync(
        string command,
        string[] args,
        GuestConfig config,
        string configPath,
        string driveLetter,
        GuestHandshakeState handshakeState)
    {
        if (command is "install-autostart" or "remove-autostart" or "autostart-status")
        {
            return await ExecuteAutostartCommandAsync(command, args, config, configPath, handshakeState);
        }

        if (!config.Enabled)
        {
            GuestLogger.Info("agent.disabled", "HyperTool.Guest ist deaktiviert (enabled=false).");
            return 0;
        }

        if (string.IsNullOrWhiteSpace(config.SharePath))
        {
            GuestLogger.Error("config.invalid", "SharePath ist leer.");
            return 1;
        }

        return command switch
        {
            "status" => await PrintStatusAsync(driveLetter),
            "once" => await EnsureMappedAsync(config, driveLetter),
            "unmap" => await UnmapDriveAsync(driveLetter),
            "run" => await RunLoopAsync(config, driveLetter, handshakeState),
            "handshake" => await WriteHandshakeNowAsync(config, driveLetter, handshakeState),
            _ => UnknownCommand(command)
        };
    }

    private static int UnknownCommand(string command)
    {
        GuestLogger.Error("command.invalid", "Unbekannter Befehl.", new
        {
            command,
            supported = new[]
            {
                "status",
                "once",
                "run",
                "unmap",
                "handshake",
                "install-autostart",
                "remove-autostart",
                "autostart-status"
            }
        });

        return 1;
    }

    private static async Task<int> ExecuteAutostartCommandAsync(
        string command,
        string[] args,
        GuestConfig config,
        string configPath,
        GuestHandshakeState handshakeState)
    {
        if (command == "autostart-status")
        {
            var status = await GuestAutostartService.QueryStatusAsync(config);
            GuestLogger.Info("autostart.status", "Autostart-Status geladen.", status);
            handshakeState.RunRegistryInstalled = status.RunRegistryInstalled;
            handshakeState.TaskInstalled = status.TaskInstalled;
            return 0;
        }

        if (command == "remove-autostart")
        {
            var removeCode = await GuestAutostartService.RemoveAsync(config);
            if (removeCode == 0)
            {
                GuestLogger.Info("autostart.remove", "Autostart-Einträge entfernt.");
            }
            else
            {
                GuestLogger.Warn("autostart.remove", "Autostart-Einträge konnten nicht vollständig entfernt werden.", new { removeCode });
            }

            return removeCode;
        }

        var overrideMode = ResolveAutostartModeOverride(args);
        var mode = string.IsNullOrWhiteSpace(overrideMode) ? config.Autostart.PreferredMode : overrideMode;
        var installCode = await GuestAutostartService.InstallAsync(config, configPath, mode);

        if (installCode == 0)
        {
            GuestLogger.Info("autostart.install", "Autostart installiert.", new { mode });
        }
        else
        {
            GuestLogger.Error("autostart.install", "Autostart-Installation fehlgeschlagen.", new { mode, installCode });
        }

        return installCode;
    }

    private static string? ResolveAutostartModeOverride(string[] args)
    {
        foreach (var arg in args)
        {
            var value = arg.Trim().ToLowerInvariant();
            if (value is "run-registry" or "task")
            {
                return value;
            }
        }

        return null;
    }

    private static async Task<int> RunLoopAsync(GuestConfig config, string driveLetter, GuestHandshakeState handshakeState)
    {
        using var cts = new CancellationTokenSource();
        Console.CancelKeyPress += (_, eventArgs) =>
        {
            eventArgs.Cancel = true;
            cts.Cancel();
        };

        GuestLogger.Info("loop.start", "Run-Loop gestartet.", new { config.PollIntervalSeconds });

        while (!cts.IsCancellationRequested)
        {
            var ensureCode = await EnsureMappedAsync(config, driveLetter);
            handshakeState.LastExitCode = ensureCode;
            handshakeState.LastError = ensureCode == 0 ? string.Empty : "Mapping fehlgeschlagen.";
            await UpdateHandshakeStateAsync(config, driveLetter, handshakeState);
            GuestHandshakeWriter.TryWrite(config, handshakeState);

            try
            {
                await Task.Delay(TimeSpan.FromSeconds(config.PollIntervalSeconds), cts.Token);
            }
            catch (TaskCanceledException)
            {
                break;
            }
        }

        GuestLogger.Info("loop.stop", "Run-Loop beendet.");
        return 0;
    }

    private static async Task<int> PrintStatusAsync(string driveLetter)
    {
        var status = await QueryDriveMappingAsync(driveLetter);
        if (!status.Exists)
        {
            GuestLogger.Info("mapping.status", $"{driveLetter}: nicht verbunden.");
            return 0;
        }

        GuestLogger.Info("mapping.status", $"{driveLetter}: verbunden.", new { status.RemotePath });
        return 0;
    }

    private static async Task<int> WriteHandshakeNowAsync(GuestConfig config, string driveLetter, GuestHandshakeState handshakeState)
    {
        await UpdateHandshakeStateAsync(config, driveLetter, handshakeState);
        GuestHandshakeWriter.TryWrite(config, handshakeState);
        GuestLogger.Info("handshake.write", "Handshake-Datei aktualisiert.");
        return 0;
    }

    private static async Task<int> EnsureMappedAsync(GuestConfig config, string driveLetter)
    {
        var targetPath = ResolveEffectiveSharePath(config);
        var mapping = await QueryDriveMappingAsync(driveLetter);

        if (mapping.Exists
            && string.Equals(NormalizeUnc(mapping.RemotePath), NormalizeUnc(targetPath), StringComparison.OrdinalIgnoreCase))
        {
            GuestLogger.Info("mapping.ok", $"{driveLetter}: bereits korrekt gemappt.", new { targetPath });
            return 0;
        }

        if (mapping.Exists)
        {
            GuestLogger.Warn("mapping.repair", $"{driveLetter}: Mapping wird ersetzt.", new { current = mapping.RemotePath, targetPath });
            var unmapCode = await RunProcessAsync("net", $"use {driveLetter}: /delete /y");
            if (unmapCode != 0)
            {
                GuestLogger.Error("mapping.unmap_failed", "Bestehendes Mapping konnte nicht entfernt werden.", new { unmapCode });
                return 1;
            }
        }

        var persistentText = config.Persistent ? "yes" : "no";
        var mapArgs = string.IsNullOrWhiteSpace(config.Credential.Username)
            ? $"use {driveLetter}: \"{targetPath}\" /persistent:{persistentText}"
            : $"use {driveLetter}: \"{targetPath}\" \"{config.Credential.Password}\" /user:\"{config.Credential.Username}\" /persistent:{persistentText}";

        GuestLogger.Info("mapping.start", "Laufwerk wird gemappt.", new { driveLetter, targetPath, config.Persistent });
        var mapCode = await RunProcessAsync("net", mapArgs);
        if (mapCode != 0)
        {
            GuestLogger.Error("mapping.failed", "Mapping fehlgeschlagen.", new { mapCode, driveLetter, targetPath });
            return 1;
        }

        GuestLogger.Info("mapping.success", "Mapping erfolgreich.", new { driveLetter, targetPath });
        return 0;
    }

    private static string ResolveEffectiveSharePath(GuestConfig config)
    {
        var sharePath = (config.SharePath ?? string.Empty).Trim();
        if (!sharePath.StartsWith("\\\\", StringComparison.Ordinal))
        {
            return sharePath;
        }

        var hostAddress = (config.Usb?.HostAddress ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(hostAddress))
        {
            return sharePath;
        }

        var withoutPrefix = sharePath[2..];
        var separatorIndex = withoutPrefix.IndexOf('\\');
        if (separatorIndex <= 0)
        {
            return sharePath;
        }

        var uncHost = withoutPrefix[..separatorIndex];
        if (!string.Equals(uncHost, "HOST", StringComparison.OrdinalIgnoreCase))
        {
            return sharePath;
        }

        var suffix = withoutPrefix[separatorIndex..];
        var resolved = $"\\\\{hostAddress}{suffix}";

        GuestLogger.Info("mapping.sharepath.resolve", "SharePath-Hostplatzhalter aufgelöst.", new
        {
            configuredSharePath = sharePath,
            resolvedSharePath = resolved,
            hostAddress
        });

        return resolved;
    }

    private static async Task<int> UnmapDriveAsync(string driveLetter)
    {
        var mapping = await QueryDriveMappingAsync(driveLetter);
        if (!mapping.Exists)
        {
            GuestLogger.Info("unmap.noop", $"{driveLetter}: bereits nicht verbunden.");
            return 0;
        }

        var code = await RunProcessAsync("net", $"use {driveLetter}: /delete /y");
        if (code != 0)
        {
            GuestLogger.Error("unmap.failed", "Unmap fehlgeschlagen.", new { code, driveLetter });
            return 1;
        }

        GuestLogger.Info("unmap.success", "Laufwerk getrennt.", new { driveLetter });
        return 0;
    }

    private static async Task UpdateHandshakeStateAsync(GuestConfig config, string driveLetter, GuestHandshakeState handshakeState)
    {
        var mapping = await QueryDriveMappingAsync(driveLetter);
        var autostart = await GuestAutostartService.QueryStatusAsync(config);

        handshakeState.TimestampUtc = DateTime.UtcNow;
        handshakeState.IsMapped = mapping.Exists;
        handshakeState.MappedRemotePath = mapping.RemotePath;
        handshakeState.RunRegistryInstalled = autostart.RunRegistryInstalled;
        handshakeState.TaskInstalled = autostart.TaskInstalled;
    }

    private static async Task<DriveMappingStatus> QueryDriveMappingAsync(string driveLetter)
    {
        var result = await RunProcessWithOutputAsync("net", $"use {driveLetter}:");
        if (result.ExitCode != 0)
        {
            return new DriveMappingStatus(false, string.Empty);
        }

        var lines = result.Output
            .Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries)
            .Select(line => line.Trim())
            .ToList();

        var remotePath = lines
            .Select(line => line.Split(' ', StringSplitOptions.RemoveEmptyEntries))
            .FirstOrDefault(parts => parts.Length >= 3 && parts[1].Equals($"{driveLetter}:", StringComparison.OrdinalIgnoreCase))?
            .ElementAtOrDefault(2)
            ?? string.Empty;

        return new DriveMappingStatus(!string.IsNullOrWhiteSpace(remotePath), remotePath);
    }

    private static async Task<int> RunProcessAsync(string fileName, string arguments)
    {
        var result = await RunProcessWithOutputAsync(fileName, arguments);
        if (!string.IsNullOrWhiteSpace(result.Output))
        {
            GuestLogger.Info("process.output", "Prozessausgabe.", new { fileName, arguments, result.Output });
        }

        return result.ExitCode;
    }

    private static async Task<ProcessResult> RunProcessWithOutputAsync(string fileName, string arguments)
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

        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();

        await process.WaitForExitAsync();

        var stdout = await stdoutTask;
        var stderr = await stderrTask;
        var output = string.IsNullOrWhiteSpace(stderr) ? stdout : $"{stdout}{Environment.NewLine}{stderr}";

        return new ProcessResult(process.ExitCode, output);
    }

    private static string NormalizeUnc(string value)
        => value.Trim().TrimEnd('\\');

    private readonly record struct DriveMappingStatus(bool Exists, string RemotePath);

    private readonly record struct ProcessResult(int ExitCode, string Output);
}
