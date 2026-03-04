using HyperTool.Models;
using Microsoft.Win32;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;

namespace HyperTool.Services;

public sealed class HyperVSocketFileHostListener : IDisposable
{
    private const int MaxReadChunkBytes = 1024 * 1024;

    private readonly Guid _serviceId;
    private readonly Func<IReadOnlyList<HostSharedFolderDefinition>> _catalogProvider;
    private readonly Action<DateTimeOffset>? _onRequestServed;
    private Socket? _listener;
    private CancellationTokenSource? _cts;
    private Task? _acceptLoopTask;

    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true
    };

    public HyperVSocketFileHostListener(
        Func<IReadOnlyList<HostSharedFolderDefinition>> catalogProvider,
        Action<DateTimeOffset>? onRequestServed = null,
        Guid? serviceId = null)
    {
        _catalogProvider = catalogProvider ?? throw new ArgumentNullException(nameof(catalogProvider));
        _onRequestServed = onRequestServed;
        _serviceId = serviceId ?? HyperVSocketUsbTunnelDefaults.FileServiceId;
    }

    public bool IsRunning { get; private set; }

    public void Start()
    {
        if (IsRunning)
        {
            return;
        }

        if (!EnsureServiceGuidRegistration())
        {
            throw new InvalidOperationException("Hyper-V Socket File-Dienst ist nicht registriert. Starte HyperTool Host als Administrator, um den Dienst einmalig zu registrieren.");
        }

        var listener = new Socket((AddressFamily)34, SocketType.Stream, (ProtocolType)1);
        listener.Bind(new HyperVSocketEndPoint(HyperVSocketUsbTunnelDefaults.VmIdWildcard, _serviceId));
        listener.Listen(32);

        _listener = listener;
        _cts = new CancellationTokenSource();
        _acceptLoopTask = Task.Run(() => AcceptLoopAsync(_cts.Token));
        IsRunning = true;
    }

    private async Task AcceptLoopAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            Socket? socket = null;
            try
            {
                if (_listener is null)
                {
                    break;
                }

                socket = await _listener.AcceptAsync(cancellationToken);
                _ = Task.Run(() => HandleClientAsync(socket, cancellationToken), cancellationToken);
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch
            {
                socket?.Dispose();
                if (cancellationToken.IsCancellationRequested)
                {
                    break;
                }
            }
        }
    }

    private async Task HandleClientAsync(Socket socket, CancellationToken cancellationToken)
    {
        await using var stream = new NetworkStream(socket, ownsSocket: true);
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 16 * 1024, leaveOpen: true);
        await using var writer = new StreamWriter(stream, Encoding.UTF8, bufferSize: 16 * 1024, leaveOpen: false) { NewLine = "\n" };

        while (!cancellationToken.IsCancellationRequested)
        {
            string? line;
            try
            {
                line = await reader.ReadLineAsync(cancellationToken);
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch
            {
                break;
            }

            if (line is null)
            {
                break;
            }

            HostFileServiceResponse response;
            try
            {
                var request = JsonSerializer.Deserialize<HostFileServiceRequest>(line, SerializerOptions);
                if (request is null)
                {
                    response = CreateErrorResponse(string.Empty, "bad-request", "Ungültige Anfrage.");
                }
                else
                {
                    response = await ProcessRequestAsync(request, cancellationToken);
                }
            }
            catch (Exception ex)
            {
                response = CreateErrorResponse(string.Empty, "server-error", ex.Message);
            }

            var payload = JsonSerializer.Serialize(response, SerializerOptions);
            await writer.WriteLineAsync(payload.AsMemory(), cancellationToken);
            await writer.FlushAsync(cancellationToken);

            try
            {
                _onRequestServed?.Invoke(DateTimeOffset.UtcNow);
            }
            catch
            {
            }
        }
    }

    private async Task<HostFileServiceResponse> ProcessRequestAsync(HostFileServiceRequest request, CancellationToken cancellationToken)
    {
        var op = (request.Operation ?? string.Empty).Trim().ToLowerInvariant();
        return op switch
        {
            "ping" => CreateSuccessResponse(request.RequestId),
            "list-shares" => HandleListShares(request),
            "metadata" => HandleMetadata(request),
            "list-directory" => HandleListDirectory(request),
            "read-file" => await HandleReadFileAsync(request, cancellationToken),
            "write-file" => await HandleWriteFileAsync(request, cancellationToken),
            "set-length" => HandleSetLength(request),
            "create-directory" => HandleCreateDirectory(request),
            "delete" => HandleDelete(request),
            "move" => HandleMove(request),
            _ => CreateErrorResponse(request.RequestId, "unsupported-operation", $"Unbekannte Operation '{request.Operation}'.")
        };
    }

    private HostFileServiceResponse HandleListShares(HostFileServiceRequest request)
    {
        var catalog = _catalogProvider() ?? [];
        var shares = catalog
            .Where(static item => item.Enabled)
            .Select(static item => new HostFileServiceShareInfo
            {
                Id = item.Id,
                Label = item.Label,
                ShareName = item.ShareName,
                ReadOnly = item.ReadOnly
            })
            .ToList();

        return new HostFileServiceResponse
        {
            RequestId = request.RequestId,
            Success = true,
            Shares = shares
        };
    }

    private HostFileServiceResponse HandleMetadata(HostFileServiceRequest request)
    {
        var resolve = ResolveTargetPath(request, requireWrite: false);
        if (!resolve.Success)
        {
            return resolve.Error;
        }

        var physicalPath = resolve.PhysicalPath;
        var response = CreateSuccessResponse(request.RequestId);
        response.ShareId = resolve.Share.Id;
        response.RelativePath = resolve.RelativePath;

        if (Directory.Exists(physicalPath))
        {
            var info = new DirectoryInfo(physicalPath);
            response.Exists = true;
            response.IsDirectory = true;
            response.Size = 0;
            response.LastWriteTimeUtc = info.LastWriteTimeUtc;
            return response;
        }

        if (File.Exists(physicalPath))
        {
            var info = new FileInfo(physicalPath);
            response.Exists = true;
            response.IsDirectory = false;
            response.Size = info.Length;
            response.LastWriteTimeUtc = info.LastWriteTimeUtc;
            return response;
        }

        response.Exists = false;
        return response;
    }

    private HostFileServiceResponse HandleListDirectory(HostFileServiceRequest request)
    {
        var resolve = ResolveTargetPath(request, requireWrite: false);
        if (!resolve.Success)
        {
            return resolve.Error;
        }

        var physicalPath = resolve.PhysicalPath;
        if (!Directory.Exists(physicalPath))
        {
            return CreateErrorResponse(request.RequestId, "not-found", "Verzeichnis wurde nicht gefunden.");
        }

        var entries = new List<HostFileServiceEntryInfo>();

        foreach (var directory in Directory.EnumerateDirectories(physicalPath))
        {
            var info = new DirectoryInfo(directory);
            entries.Add(new HostFileServiceEntryInfo
            {
                Name = info.Name,
                RelativePath = CombineRelative(resolve.RelativePath, info.Name),
                IsDirectory = true,
                Size = 0,
                LastWriteTimeUtc = info.LastWriteTimeUtc,
                ReadOnly = info.Attributes.HasFlag(FileAttributes.ReadOnly)
            });
        }

        foreach (var file in Directory.EnumerateFiles(physicalPath))
        {
            var info = new FileInfo(file);
            entries.Add(new HostFileServiceEntryInfo
            {
                Name = info.Name,
                RelativePath = CombineRelative(resolve.RelativePath, info.Name),
                IsDirectory = false,
                Size = info.Length,
                LastWriteTimeUtc = info.LastWriteTimeUtc,
                ReadOnly = info.IsReadOnly
            });
        }

        return new HostFileServiceResponse
        {
            RequestId = request.RequestId,
            Success = true,
            ShareId = resolve.Share.Id,
            RelativePath = resolve.RelativePath,
            Entries = entries.OrderBy(static item => item.IsDirectory ? 0 : 1).ThenBy(static item => item.Name, StringComparer.OrdinalIgnoreCase).ToList()
        };
    }

    private async Task<HostFileServiceResponse> HandleReadFileAsync(HostFileServiceRequest request, CancellationToken cancellationToken)
    {
        var resolve = ResolveTargetPath(request, requireWrite: false);
        if (!resolve.Success)
        {
            return resolve.Error;
        }

        if (!File.Exists(resolve.PhysicalPath))
        {
            return CreateErrorResponse(request.RequestId, "not-found", "Datei wurde nicht gefunden.");
        }

        var length = request.Length <= 0 ? 64 * 1024 : Math.Min(request.Length, MaxReadChunkBytes);
        var offset = Math.Max(0, request.Offset);

        await using var stream = new FileStream(resolve.PhysicalPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 64 * 1024, useAsync: true);
        if (offset > stream.Length)
        {
            offset = stream.Length;
        }

        stream.Position = offset;
        var buffer = new byte[length];
        var bytesRead = await stream.ReadAsync(buffer.AsMemory(0, length), cancellationToken);

        return new HostFileServiceResponse
        {
            RequestId = request.RequestId,
            Success = true,
            ShareId = resolve.Share.Id,
            RelativePath = resolve.RelativePath,
            Exists = true,
            IsDirectory = false,
            Size = stream.Length,
            LastWriteTimeUtc = File.GetLastWriteTimeUtc(resolve.PhysicalPath),
            BytesRead = bytesRead,
            DataBase64 = bytesRead == 0 ? string.Empty : Convert.ToBase64String(buffer, 0, bytesRead)
        };
    }

    private async Task<HostFileServiceResponse> HandleWriteFileAsync(HostFileServiceRequest request, CancellationToken cancellationToken)
    {
        var resolve = ResolveTargetPath(request, requireWrite: true);
        if (!resolve.Success)
        {
            return resolve.Error;
        }

        var payload = string.IsNullOrWhiteSpace(request.DataBase64)
            ? []
            : Convert.FromBase64String(request.DataBase64);

        var directory = Path.GetDirectoryName(resolve.PhysicalPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var mode = request.Append
            ? FileMode.Append
            : request.Truncate
                ? FileMode.Create
                : File.Exists(resolve.PhysicalPath)
                    ? FileMode.Open
                    : FileMode.CreateNew;

        await using var stream = new FileStream(resolve.PhysicalPath, mode, FileAccess.Write, FileShare.Read, 64 * 1024, useAsync: true);
        if (!request.Append)
        {
            var offset = Math.Max(0, request.Offset);
            stream.Position = offset;
        }

        if (payload.Length > 0)
        {
            await stream.WriteAsync(payload.AsMemory(0, payload.Length), cancellationToken);
        }

        await stream.FlushAsync(cancellationToken);

        return new HostFileServiceResponse
        {
            RequestId = request.RequestId,
            Success = true,
            ShareId = resolve.Share.Id,
            RelativePath = resolve.RelativePath,
            Exists = true,
            IsDirectory = false,
            Size = stream.Length,
            LastWriteTimeUtc = File.GetLastWriteTimeUtc(resolve.PhysicalPath)
        };
    }

    private HostFileServiceResponse HandleCreateDirectory(HostFileServiceRequest request)
    {
        var resolve = ResolveTargetPath(request, requireWrite: true);
        if (!resolve.Success)
        {
            return resolve.Error;
        }

        Directory.CreateDirectory(resolve.PhysicalPath);

        return new HostFileServiceResponse
        {
            RequestId = request.RequestId,
            Success = true,
            ShareId = resolve.Share.Id,
            RelativePath = resolve.RelativePath,
            Exists = true,
            IsDirectory = true,
            LastWriteTimeUtc = Directory.GetLastWriteTimeUtc(resolve.PhysicalPath)
        };
    }

    private HostFileServiceResponse HandleSetLength(HostFileServiceRequest request)
    {
        var resolve = ResolveTargetPath(request, requireWrite: true);
        if (!resolve.Success)
        {
            return resolve.Error;
        }

        if (!File.Exists(resolve.PhysicalPath))
        {
            return CreateErrorResponse(request.RequestId, "not-found", "Datei wurde nicht gefunden.");
        }

        var nextLength = Math.Max(0, request.NewSize);

        using (var stream = new FileStream(resolve.PhysicalPath, FileMode.Open, FileAccess.Write, FileShare.Read))
        {
            stream.SetLength(nextLength);
            stream.Flush();
        }

        var info = new FileInfo(resolve.PhysicalPath);
        return new HostFileServiceResponse
        {
            RequestId = request.RequestId,
            Success = true,
            ShareId = resolve.Share.Id,
            RelativePath = resolve.RelativePath,
            Exists = true,
            IsDirectory = false,
            Size = info.Length,
            LastWriteTimeUtc = info.LastWriteTimeUtc
        };
    }

    private HostFileServiceResponse HandleDelete(HostFileServiceRequest request)
    {
        var resolve = ResolveTargetPath(request, requireWrite: true);
        if (!resolve.Success)
        {
            return resolve.Error;
        }

        if (File.Exists(resolve.PhysicalPath))
        {
            File.Delete(resolve.PhysicalPath);
            return CreateSuccessResponse(request.RequestId);
        }

        if (Directory.Exists(resolve.PhysicalPath))
        {
            Directory.Delete(resolve.PhysicalPath, recursive: request.Recursive);
            return CreateSuccessResponse(request.RequestId);
        }

        return CreateErrorResponse(request.RequestId, "not-found", "Pfad wurde nicht gefunden.");
    }

    private HostFileServiceResponse HandleMove(HostFileServiceRequest request)
    {
        var source = ResolveTargetPath(request, requireWrite: true);
        if (!source.Success)
        {
            return source.Error;
        }

        var moveRequest = new HostFileServiceRequest
        {
            RequestId = request.RequestId,
            ShareId = request.ShareId,
            RelativePath = request.TargetRelativePath
        };

        var target = ResolveTargetPath(moveRequest, requireWrite: true);
        if (!target.Success)
        {
            return target.Error;
        }

        var targetDirectory = Path.GetDirectoryName(target.PhysicalPath);
        if (!string.IsNullOrWhiteSpace(targetDirectory))
        {
            Directory.CreateDirectory(targetDirectory);
        }

        if (File.Exists(source.PhysicalPath))
        {
            if (File.Exists(target.PhysicalPath))
            {
                if (!request.Overwrite)
                {
                    return CreateErrorResponse(request.RequestId, "already-exists", "Ziel existiert bereits.");
                }

                File.Delete(target.PhysicalPath);
            }

            File.Move(source.PhysicalPath, target.PhysicalPath);
            return CreateSuccessResponse(request.RequestId);
        }

        if (Directory.Exists(source.PhysicalPath))
        {
            if (Directory.Exists(target.PhysicalPath))
            {
                if (!request.Overwrite)
                {
                    return CreateErrorResponse(request.RequestId, "already-exists", "Zielverzeichnis existiert bereits.");
                }

                Directory.Delete(target.PhysicalPath, recursive: true);
            }

            Directory.Move(source.PhysicalPath, target.PhysicalPath);
            return CreateSuccessResponse(request.RequestId);
        }

        return CreateErrorResponse(request.RequestId, "not-found", "Quellpfad wurde nicht gefunden.");
    }

    private (bool Success, HostFileServiceResponse Error, HostSharedFolderDefinition Share, string RelativePath, string PhysicalPath) ResolveTargetPath(HostFileServiceRequest request, bool requireWrite)
    {
        var catalog = _catalogProvider() ?? [];
        var share = catalog.FirstOrDefault(item =>
            item.Enabled
            && (!string.IsNullOrWhiteSpace(request.ShareId)
                && (string.Equals(item.Id, request.ShareId, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(item.ShareName, request.ShareId, StringComparison.OrdinalIgnoreCase))));

        if (share is null)
        {
            return (false, CreateErrorResponse(request.RequestId, "share-not-found", "Freigabe wurde nicht gefunden oder ist deaktiviert."), new HostSharedFolderDefinition(), string.Empty, string.Empty);
        }

        if (requireWrite && share.ReadOnly)
        {
            return (false, CreateErrorResponse(request.RequestId, "read-only", "Freigabe ist schreibgeschützt."), share, string.Empty, string.Empty);
        }

        var rootPath = (share.LocalPath ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(rootPath) || !Directory.Exists(rootPath))
        {
            return (false, CreateErrorResponse(request.RequestId, "share-path-missing", "Lokaler Freigabe-Pfad ist nicht verfügbar."), share, string.Empty, string.Empty);
        }

        var relativePath = NormalizeRelativePath(request.RelativePath);
        if (relativePath is null)
        {
            return (false, CreateErrorResponse(request.RequestId, "invalid-path", "Ungültiger relativer Pfad."), share, string.Empty, string.Empty);
        }

        var fullRoot = Path.GetFullPath(rootPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var fullPath = string.IsNullOrWhiteSpace(relativePath)
            ? fullRoot
            : Path.GetFullPath(Path.Combine(fullRoot, relativePath));

        if (!string.Equals(fullPath, fullRoot, StringComparison.OrdinalIgnoreCase)
            && !fullPath.StartsWith(fullRoot + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
        {
            return (false, CreateErrorResponse(request.RequestId, "invalid-path", "Pfad liegt außerhalb der Freigabe."), share, string.Empty, string.Empty);
        }

        return (true, CreateSuccessResponse(request.RequestId), share, relativePath, fullPath);
    }

    private static string? NormalizeRelativePath(string? value)
    {
        var normalized = (value ?? string.Empty).Trim().Replace('/', '\\').TrimStart('\\');
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return string.Empty;
        }

        if (Path.IsPathRooted(normalized))
        {
            return null;
        }

        var segments = normalized.Split('\\', StringSplitOptions.RemoveEmptyEntries);
        if (segments.Any(segment => segment == "." || segment == ".."))
        {
            return null;
        }

        return string.Join(Path.DirectorySeparatorChar, segments);
    }

    private static string CombineRelative(string basePath, string name)
    {
        if (string.IsNullOrWhiteSpace(basePath))
        {
            return name;
        }

        return $"{basePath.TrimEnd(Path.DirectorySeparatorChar)}{Path.DirectorySeparatorChar}{name}";
    }

    private bool EnsureServiceGuidRegistration()
    {
        if (IsServiceGuidRegistered())
        {
            return true;
        }

        if (TryRegisterServiceGuid())
        {
            return true;
        }

        return IsServiceGuidRegistered();
    }

    private bool TryRegisterServiceGuid()
    {
        try
        {
            const string rootPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices";

            using var rootKey = Registry.LocalMachine.CreateSubKey(rootPath, writable: true);
            if (rootKey is null)
            {
                return false;
            }

            using var serviceKey = rootKey.CreateSubKey(_serviceId.ToString("D"), writable: true);
            serviceKey?.SetValue("ElementName", "HyperTool Hyper-V Socket File Service", RegistryValueKind.String);
            return serviceKey is not null;
        }
        catch
        {
            return false;
        }
    }

    private bool IsServiceGuidRegistered()
    {
        try
        {
            const string rootPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices";

            using var rootKey = Registry.LocalMachine.OpenSubKey(rootPath, writable: false);
            if (rootKey is null)
            {
                return false;
            }

            using var serviceKey = rootKey.OpenSubKey(_serviceId.ToString("D"), writable: false);
            return serviceKey is not null;
        }
        catch
        {
            return false;
        }
    }

    private static HostFileServiceResponse CreateSuccessResponse(string requestId)
    {
        return new HostFileServiceResponse
        {
            RequestId = requestId,
            Success = true
        };
    }

    private static HostFileServiceResponse CreateErrorResponse(string requestId, string code, string message)
    {
        return new HostFileServiceResponse
        {
            RequestId = requestId,
            Success = false,
            ErrorCode = code,
            ErrorMessage = message
        };
    }

    public void Dispose()
    {
        if (!IsRunning)
        {
            return;
        }

        IsRunning = false;

        try
        {
            _cts?.Cancel();
        }
        catch
        {
        }

        try
        {
            _listener?.Dispose();
        }
        catch
        {
        }

        try
        {
            _acceptLoopTask?.Wait(TimeSpan.FromMilliseconds(250));
        }
        catch
        {
        }

        _cts?.Dispose();
        _cts = null;
        _listener = null;
        _acceptLoopTask = null;
    }
}
