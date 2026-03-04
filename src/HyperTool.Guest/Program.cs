using System.Diagnostics;
using HyperTool.Services;

namespace HyperTool.Guest;

internal static class GuestCli
{
    public static async Task<int> ExecuteAsync(string[] args)
    {
        _winFspMountService ??= new GuestWinFspMountService();

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
                    FileName = "hypertool-guest.log",
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
        finally
        {
            if (command is "once" or "status" or "unmap" or "handshake")
            {
                _winFspMountService?.Dispose();
                _winFspMountService = null;
            }
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
        _activeConfig = config;

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
            "unmap" => await UnmapDriveAsync(config, driveLetter),
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
        if (_activeConfig is null)
        {
            GuestLogger.Error("mapping.status", "Konfiguration nicht geladen.");
            return 1;
        }

        try
        {
            var mapping = new GuestSharedFolderMapping
            {
                SharePath = _activeConfig.SharePath,
                DriveLetter = driveLetter,
                Enabled = true,
                Persistent = true,
                Label = "CLI"
            };

            await (_winFspMountService ??= new GuestWinFspMountService()).EnsureMountedAsync(mapping, CancellationToken.None);
            GuestLogger.Info("mapping.status", "HyperTool-File-Modus aktiv und erreichbar.", new
            {
                mode = "hypertool-file",
                driveLetter,
                sharePath = _activeConfig.SharePath
            });
            return 0;
        }
        catch (Exception ex)
        {
            GuestLogger.Error("mapping.status", "HyperTool-File-Modus nicht erreichbar.", new
            {
                mode = "hypertool-file",
                driveLetter,
                sharePath = _activeConfig.SharePath,
                error = ex.Message
            });
            return 1;
        }
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
        _activeConfig = config;

        try
        {
            var fileModeMapping = new GuestSharedFolderMapping
            {
                SharePath = config.SharePath,
                DriveLetter = driveLetter,
                Enabled = true,
                Persistent = true,
                Label = "CLI"
            };

            await (_winFspMountService ??= new GuestWinFspMountService()).EnsureMountedAsync(fileModeMapping, CancellationToken.None);
            GuestLogger.Info("mapping.success", "HyperTool-File-Modus bereit.", new
            {
                mode = "hypertool-file",
                driveLetter,
                sharePath = config.SharePath
            });
            return 0;
        }
        catch (Exception ex)
        {
            GuestLogger.Error("mapping.failed", "HyperTool-File-Modus konnte nicht vorbereitet werden.", new
            {
                mode = "hypertool-file",
                driveLetter,
                sharePath = config.SharePath,
                error = ex.Message
            });
            return 1;
        }
    }

    private static async Task<int> UnmapDriveAsync(GuestConfig config, string driveLetter)
    {
        await (_winFspMountService ??= new GuestWinFspMountService()).UnmountAsync(driveLetter, CancellationToken.None);
        GuestLogger.Info("unmap.success", "HyperTool-File-Mount getrennt.", new { driveLetter });
        return 0;
    }

    private static async Task UpdateHandshakeStateAsync(GuestConfig config, string driveLetter, GuestHandshakeState handshakeState)
    {
        try
        {
            await GuestFileServiceModeHelper.EnsureShareAvailableAsync(config.SharePath, CancellationToken.None);
            handshakeState.IsMapped = true;
            handshakeState.MappedRemotePath = $"hypertool-file://{GuestFileServiceModeHelper.ResolveShareName(config.SharePath)}";
        }
        catch
        {
            handshakeState.IsMapped = false;
            handshakeState.MappedRemotePath = string.Empty;
        }

        var autostart = await GuestAutostartService.QueryStatusAsync(config);

        handshakeState.TimestampUtc = DateTime.UtcNow;
        handshakeState.RunRegistryInstalled = autostart.RunRegistryInstalled;
        handshakeState.TaskInstalled = autostart.TaskInstalled;
    }

    private static GuestConfig? _activeConfig;
    private static GuestWinFspMountService? _winFspMountService;

}
