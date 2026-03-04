using Fsp;
using Fsp.Interop;
using HyperTool.Models;
using HyperTool.Services;
using System.Collections.Concurrent;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using FspFileInfo = Fsp.Interop.FileInfo;
using FspVolumeInfo = Fsp.Interop.VolumeInfo;

namespace HyperTool.Guest;

internal static class GuestWinFspMountRegistry
{
    public static GuestWinFspMountService Instance { get; } = new();
}

internal sealed class GuestWinFspMountService : IDisposable
{
    private readonly ConcurrentDictionary<string, MountedFileSystem> _mountedByDrive = new(StringComparer.OrdinalIgnoreCase);

    public async Task EnsureCatalogMountedAsync(string driveLetter, IReadOnlyList<GuestSharedFolderMapping> mappings, CancellationToken cancellationToken)
    {
        var normalizedDrive = GuestConfigService.NormalizeDriveLetter(driveLetter);
        var enabledMappings = (mappings ?? [])
            .Where(mapping => mapping is not null
                              && mapping.Enabled
                              && !string.IsNullOrWhiteSpace(mapping.SharePath))
            .ToList();

        var catalogEntries = new List<CatalogShareEntry>(enabledMappings.Count);
        var usedRootNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var mapping in enabledMappings)
        {
            var shareId = GuestFileServiceModeHelper.ResolveShareName(mapping.SharePath);
            if (string.IsNullOrWhiteSpace(shareId))
            {
                continue;
            }

            await GuestFileServiceModeHelper.EnsureShareAvailableAsync(mapping.SharePath, cancellationToken);

            var preferredRootName = (mapping.Label ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(preferredRootName))
            {
                preferredRootName = shareId;
            }

            var rootName = BuildUniqueCatalogRootName(preferredRootName, usedRootNames);
            catalogEntries.Add(new CatalogShareEntry(rootName, shareId));
        }

        if (_mountedByDrive.TryGetValue(normalizedDrive, out var existing) && existing.IsCatalog)
        {
            existing.UpdateCatalog(catalogEntries, "HyperTool");
            return;
        }

        await UnmountAsync(normalizedDrive, cancellationToken);

        MountedFileSystem mounted;
        try
        {
            mounted = MountCatalogInternal(normalizedDrive, catalogEntries, "HyperTool");
        }
        catch (DllNotFoundException ex)
        {
            throw new InvalidOperationException("WinFsp Runtime nicht gefunden. Bitte WinFsp auf dem Guest installieren.", ex);
        }
        catch (EntryPointNotFoundException ex)
        {
            throw new InvalidOperationException("WinFsp Runtime ist inkompatibel oder unvollständig installiert.", ex);
        }

        if (!_mountedByDrive.TryAdd(normalizedDrive, mounted))
        {
            mounted.Dispose();
            throw new InvalidOperationException($"WinFsp-Katalog-Mount für Laufwerk '{normalizedDrive}:' konnte nicht registriert werden.");
        }
    }

    public async Task EnsureMountedAsync(GuestSharedFolderMapping mapping, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(mapping);

        var driveLetter = GuestConfigService.NormalizeDriveLetter(mapping.DriveLetter);
        var shareId = GuestFileServiceModeHelper.ResolveShareName(mapping.SharePath);
        var volumeLabel = ResolveVolumeLabel(mapping, shareId);

        if (string.IsNullOrWhiteSpace(shareId))
        {
            throw new InvalidOperationException("Ungültiger SharePath für WinFsp-Mount.");
        }

        await GuestFileServiceModeHelper.EnsureShareAvailableAsync(mapping.SharePath, cancellationToken);

        if (_mountedByDrive.TryGetValue(driveLetter, out var existing)
            && string.Equals(existing.ShareId, shareId, StringComparison.OrdinalIgnoreCase)
            && string.Equals(existing.VolumeLabel, volumeLabel, StringComparison.Ordinal))
        {
            return;
        }

        await UnmountAsync(driveLetter, cancellationToken);

        MountedFileSystem mounted;
        try
        {
            mounted = MountInternal(driveLetter, shareId, volumeLabel);
        }
        catch (DllNotFoundException ex)
        {
            throw new InvalidOperationException("WinFsp Runtime nicht gefunden. Bitte WinFsp auf dem Guest installieren.", ex);
        }
        catch (EntryPointNotFoundException ex)
        {
            throw new InvalidOperationException("WinFsp Runtime ist inkompatibel oder unvollständig installiert.", ex);
        }

        if (!_mountedByDrive.TryAdd(driveLetter, mounted))
        {
            mounted.Dispose();
            throw new InvalidOperationException($"WinFsp-Mount für Laufwerk '{driveLetter}:' konnte nicht registriert werden.");
        }
    }

    public Task UnmountAsync(string driveLetter, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var normalizedDrive = GuestConfigService.NormalizeDriveLetter(driveLetter);
        if (_mountedByDrive.TryRemove(normalizedDrive, out var mounted))
        {
            mounted.Dispose();
        }

        return Task.CompletedTask;
    }

    public async Task UnmountManyAsync(IEnumerable<string> driveLetters, CancellationToken cancellationToken)
    {
        foreach (var driveLetter in driveLetters ?? [])
        {
            await UnmountAsync(driveLetter, cancellationToken);
        }
    }

    public bool IsMounted(string driveLetter)
    {
        var normalizedDrive = GuestConfigService.NormalizeDriveLetter(driveLetter);
        return _mountedByDrive.ContainsKey(normalizedDrive);
    }

    public void Dispose()
    {
        foreach (var pair in _mountedByDrive)
        {
            if (_mountedByDrive.TryRemove(pair.Key, out var mounted))
            {
                mounted.Dispose();
            }
        }
    }

    private static string ResolveVolumeLabel(GuestSharedFolderMapping mapping, string shareId)
    {
        var candidate = (mapping.Label ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(candidate))
        {
            candidate = shareId;
        }

        var sanitized = new string(candidate
            .Where(ch => ch >= 32 && ch != '\\' && ch != '/' && ch != ':' && ch != '*' && ch != '?' && ch != '"' && ch != '<' && ch != '>' && ch != '|')
            .ToArray())
            .Trim();

        if (string.IsNullOrWhiteSpace(sanitized))
        {
            sanitized = "HyperTool";
        }

        return sanitized.Length <= 32 ? sanitized : sanitized[..32];
    }

    private static string BuildUniqueCatalogRootName(string preferredName, ISet<string> usedRootNames)
    {
        var baseName = new string((preferredName ?? string.Empty)
            .Where(ch => ch >= 32 && ch != '\\' && ch != '/' && ch != ':' && ch != '*' && ch != '?' && ch != '"' && ch != '<' && ch != '>' && ch != '|')
            .ToArray())
            .Trim();

        if (string.IsNullOrWhiteSpace(baseName))
        {
            baseName = "Share";
        }

        if (baseName.Length > 64)
        {
            baseName = baseName[..64].Trim();
        }

        var candidate = baseName;
        var suffix = 2;
        while (!usedRootNames.Add(candidate))
        {
            candidate = $"{baseName} ({suffix})";
            suffix++;
        }

        return candidate;
    }

    private static MountedFileSystem MountInternal(string driveLetter, string shareId, string volumeLabel)
    {
        var fileSystem = new HyperToolRpcFileSystem(shareId, volumeLabel);
        var host = new FileSystemHost(fileSystem)
        {
            CaseSensitiveSearch = false,
            CasePreservedNames = true,
            UnicodeOnDisk = true,
            PersistentAcls = false,
            ReparsePoints = false,
            NamedStreams = false,
            ExtendedAttributes = false,
            FileSystemName = "HyperToolFS"
        };

        var mountPoint = $"{driveLetter}:";
        var preflight = host.Preflight(mountPoint);
        if (preflight != FileSystemBase.STATUS_SUCCESS)
        {
            host.Dispose();
            throw new InvalidOperationException($"WinFsp Preflight fehlgeschlagen für {mountPoint} (NTSTATUS={preflight}).");
        }

        var mountStatus = host.Mount(mountPoint, SecurityDescriptor: null, Synchronized: false, DebugLog: 0);
        if (mountStatus != FileSystemBase.STATUS_SUCCESS)
        {
            host.Dispose();
            throw new InvalidOperationException($"WinFsp Mount fehlgeschlagen für {mountPoint} (NTSTATUS={mountStatus}).");
        }

        return new MountedFileSystem(driveLetter, shareId, volumeLabel, host, fileSystem, isCatalog: false);
    }

    private static MountedFileSystem MountCatalogInternal(string driveLetter, IReadOnlyList<CatalogShareEntry> catalogEntries, string volumeLabel)
    {
        var catalogByRoot = catalogEntries
            .Where(entry => !string.IsNullOrWhiteSpace(entry.RootName) && !string.IsNullOrWhiteSpace(entry.ShareId))
            .GroupBy(entry => entry.RootName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First().ShareId, StringComparer.OrdinalIgnoreCase);

        var fileSystem = new HyperToolRpcFileSystem(defaultShareId: null, volumeLabel, catalogByRoot);
        var host = new FileSystemHost(fileSystem)
        {
            CaseSensitiveSearch = false,
            CasePreservedNames = true,
            UnicodeOnDisk = true,
            PersistentAcls = false,
            ReparsePoints = false,
            NamedStreams = false,
            ExtendedAttributes = false,
            FileSystemName = "HyperToolFS"
        };

        var mountPoint = $"{driveLetter}:";
        var preflight = host.Preflight(mountPoint);
        if (preflight != FileSystemBase.STATUS_SUCCESS)
        {
            host.Dispose();
            throw new InvalidOperationException($"WinFsp Preflight fehlgeschlagen für {mountPoint} (NTSTATUS={preflight}).");
        }

        var mountStatus = host.Mount(mountPoint, SecurityDescriptor: null, Synchronized: false, DebugLog: 0);
        if (mountStatus != FileSystemBase.STATUS_SUCCESS)
        {
            host.Dispose();
            throw new InvalidOperationException($"WinFsp Mount fehlgeschlagen für {mountPoint} (NTSTATUS={mountStatus}).");
        }

        return new MountedFileSystem(driveLetter, "__catalog__", volumeLabel, host, fileSystem, isCatalog: true);
    }

    private sealed class MountedFileSystem : IDisposable
    {
        public MountedFileSystem(string driveLetter, string shareId, string volumeLabel, FileSystemHost host, HyperToolRpcFileSystem fileSystem, bool isCatalog)
        {
            DriveLetter = driveLetter;
            ShareId = shareId;
            VolumeLabel = volumeLabel;
            Host = host;
            FileSystem = fileSystem;
            IsCatalog = isCatalog;
        }

        public string DriveLetter { get; }

        public string ShareId { get; }

        public string VolumeLabel { get; }

        public FileSystemHost Host { get; }

        public HyperToolRpcFileSystem FileSystem { get; }

        public bool IsCatalog { get; }

        public void UpdateCatalog(IReadOnlyList<CatalogShareEntry> catalogEntries, string volumeLabel)
        {
            if (!IsCatalog)
            {
                return;
            }

            var catalogByRoot = catalogEntries
                .Where(entry => !string.IsNullOrWhiteSpace(entry.RootName) && !string.IsNullOrWhiteSpace(entry.ShareId))
                .GroupBy(entry => entry.RootName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First().ShareId, StringComparer.OrdinalIgnoreCase);

            FileSystem.UpdateCatalog(catalogByRoot, volumeLabel);
        }

        public void Dispose()
        {
            Host.Dispose();
        }
    }

    private sealed record CatalogShareEntry(string RootName, string ShareId);
}

internal sealed class HyperToolRpcFileSystem : FileSystemBase
{
    private readonly string? _defaultShareId;
    private string _volumeLabel;
    private Dictionary<string, string> _catalogByRoot;
    private readonly byte[] _defaultSecurityDescriptor;
    private readonly object _catalogSync = new();

    private const ulong VirtualTotalSize = 512UL * 1024UL * 1024UL * 1024UL;
    private const ulong VirtualFreeSize = 256UL * 1024UL * 1024UL * 1024UL;

    public HyperToolRpcFileSystem(string? defaultShareId, string volumeLabel, IReadOnlyDictionary<string, string>? catalogByRoot = null)
    {
        _defaultShareId = string.IsNullOrWhiteSpace(defaultShareId)
            ? null
            : defaultShareId.Trim();
        _catalogByRoot = catalogByRoot is null
            ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, string>(catalogByRoot, StringComparer.OrdinalIgnoreCase);

        _volumeLabel = string.IsNullOrWhiteSpace(volumeLabel) ? "HyperTool" : volumeLabel.Trim();

        _defaultSecurityDescriptor = BuildDefaultSecurityDescriptor();
    }

    public void UpdateCatalog(IReadOnlyDictionary<string, string>? catalogByRoot, string? volumeLabel)
    {
        lock (_catalogSync)
        {
            _catalogByRoot = catalogByRoot is null
                ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                : new Dictionary<string, string>(catalogByRoot, StringComparer.OrdinalIgnoreCase);

            _volumeLabel = string.IsNullOrWhiteSpace(volumeLabel) ? "HyperTool" : volumeLabel.Trim();
        }
    }

    public override int Init(object Host)
    {
        if (Host is FileSystemHost fileSystemHost)
        {
            fileSystemHost.SectorSize = 4096;
            fileSystemHost.SectorsPerAllocationUnit = 1;
            fileSystemHost.MaxComponentLength = 255;
            fileSystemHost.CaseSensitiveSearch = false;
            fileSystemHost.CasePreservedNames = true;
            fileSystemHost.UnicodeOnDisk = true;
            fileSystemHost.PersistentAcls = false;
            fileSystemHost.FileSystemName = "HyperToolFS";
            fileSystemHost.Prefix = string.Empty;
        }

        return STATUS_SUCCESS;
    }

    public override int GetVolumeInfo(out FspVolumeInfo VolumeInfo)
    {
        VolumeInfo = new FspVolumeInfo();
        VolumeInfo.TotalSize = VirtualTotalSize;
        VolumeInfo.FreeSize = VirtualFreeSize;
        VolumeInfo.SetVolumeLabel(_volumeLabel);
        return STATUS_SUCCESS;
    }

    public override int GetSecurityByName(string FileName, out uint FileAttributes, ref byte[] SecurityDescriptor)
    {
        FileAttributes = (uint)System.IO.FileAttributes.Normal;

        var relativePath = NormalizeToRelativePath(FileName);
        if (string.IsNullOrWhiteSpace(relativePath))
        {
            FileAttributes = (uint)System.IO.FileAttributes.Directory;
            SecurityDescriptor = _defaultSecurityDescriptor;
            return STATUS_SUCCESS;
        }

        if (!TryGetMetadata(relativePath, out var metadata, out var errorStatus))
        {
            return errorStatus;
        }

        FileAttributes = ToFileAttributes(metadata.IsDirectory, readOnly: false);
        SecurityDescriptor = _defaultSecurityDescriptor;
        return STATUS_SUCCESS;
    }

    public override int Create(
        string FileName,
        uint CreateOptions,
        uint GrantedAccess,
        uint FileAttributes,
        byte[] SecurityDescriptor,
        ulong AllocationSize,
        out object FileNode,
        out object FileDesc,
        out FspFileInfo FileInfo,
        out string NormalizedName)
    {
        FileNode = new RpcPathHandle(string.Empty, isDirectory: true);
        FileDesc = new RpcPathHandle(string.Empty, isDirectory: true);
        FileInfo = new FspFileInfo();
        NormalizedName = "\\";

        var relativePath = NormalizeToRelativePath(FileName);
        var isDirectory = (CreateOptions & FILE_DIRECTORY_FILE) != 0;

        var request = new HostFileServiceRequest
        {
            RelativePath = relativePath,
            Operation = isDirectory ? "create-directory" : "write-file",
            Truncate = !isDirectory,
            DataBase64 = string.Empty,
            Length = 0,
            Offset = 0
        };

        var response = SendRequest(request);
        if (!response.Success)
        {
            return MapErrorCode(response.ErrorCode);
        }

        if (!TryGetMetadata(relativePath, out var metadata, out var status))
        {
            return status;
        }

        var handle = new RpcPathHandle(relativePath, metadata.IsDirectory);
        FileNode = handle;
        FileDesc = handle;
        FillFileInfo(metadata, ref FileInfo);
        NormalizedName = ToNormalizedName(relativePath);
        return STATUS_SUCCESS;
    }

    public override int Open(
        string FileName,
        uint CreateOptions,
        uint GrantedAccess,
        out object FileNode,
        out object FileDesc,
        out FspFileInfo FileInfo,
        out string NormalizedName)
    {
        FileNode = new RpcPathHandle(string.Empty, isDirectory: true);
        FileDesc = new RpcPathHandle(string.Empty, isDirectory: true);
        FileInfo = new FspFileInfo();
        NormalizedName = "\\";

        var relativePath = NormalizeToRelativePath(FileName);

        if (!TryGetMetadata(relativePath, out var metadata, out var status))
        {
            return status;
        }

        var requiresDirectory = (CreateOptions & FILE_DIRECTORY_FILE) != 0;
        var requiresFile = (CreateOptions & FILE_NON_DIRECTORY_FILE) != 0;

        if (requiresDirectory && !metadata.IsDirectory)
        {
            return STATUS_NOT_A_DIRECTORY;
        }

        if (requiresFile && metadata.IsDirectory)
        {
            return STATUS_FILE_IS_A_DIRECTORY;
        }

        var handle = new RpcPathHandle(relativePath, metadata.IsDirectory);
        FileNode = handle;
        FileDesc = handle;
        FillFileInfo(metadata, ref FileInfo);
        NormalizedName = ToNormalizedName(relativePath);
        return STATUS_SUCCESS;
    }

    public override int GetFileInfo(object FileNode, object FileDesc, out FspFileInfo FileInfo)
    {
        FileInfo = new FspFileInfo();

        if (!TryGetMetadataFromNode(FileNode, out var metadata, out var status))
        {
            return status;
        }

        FillFileInfo(metadata, ref FileInfo);
        return STATUS_SUCCESS;
    }

    public override int Read(object FileNode, object FileDesc, IntPtr Buffer, ulong Offset, uint Length, out uint BytesTransferred)
    {
        if (!TryGetNode(FileNode, out var node))
        {
            BytesTransferred = 0;
            return STATUS_INVALID_PARAMETER;
        }

        if (node.IsDirectory)
        {
            BytesTransferred = 0;
            return STATUS_FILE_IS_A_DIRECTORY;
        }

        var response = SendRequest(new HostFileServiceRequest
        {
            Operation = "read-file",
            RelativePath = node.RelativePath,
            Offset = checked((long)Offset),
            Length = checked((int)Math.Min(Length, int.MaxValue))
        });

        if (!response.Success)
        {
            BytesTransferred = 0;
            return MapErrorCode(response.ErrorCode);
        }

        var bytes = string.IsNullOrWhiteSpace(response.DataBase64)
            ? []
            : Convert.FromBase64String(response.DataBase64);

        if (bytes.Length > 0)
        {
            Marshal.Copy(bytes, 0, Buffer, bytes.Length);
        }

        BytesTransferred = checked((uint)bytes.Length);
        return STATUS_SUCCESS;
    }

    public override int Write(
        object FileNode,
        object FileDesc,
        IntPtr Buffer,
        ulong Offset,
        uint Length,
        bool WriteToEndOfFile,
        bool ConstrainedIo,
        out uint BytesTransferred,
        out FspFileInfo FileInfo)
    {
        FileInfo = new FspFileInfo();

        if (!TryGetNode(FileNode, out var node))
        {
            BytesTransferred = 0;
            return STATUS_INVALID_PARAMETER;
        }

        if (node.IsDirectory)
        {
            BytesTransferred = 0;
            return STATUS_FILE_IS_A_DIRECTORY;
        }

        var payload = Length == 0
            ? []
            : new byte[Length];

        if (payload.Length > 0)
        {
            Marshal.Copy(Buffer, payload, 0, payload.Length);
        }

        var response = SendRequest(new HostFileServiceRequest
        {
            Operation = "write-file",
            RelativePath = node.RelativePath,
            Offset = checked((long)Offset),
            Append = WriteToEndOfFile,
            Truncate = false,
            DataBase64 = payload.Length == 0 ? string.Empty : Convert.ToBase64String(payload)
        });

        if (!response.Success)
        {
            BytesTransferred = 0;
            return MapErrorCode(response.ErrorCode);
        }

        var metadataResponse = SendRequest(new HostFileServiceRequest
        {
            Operation = "metadata",
            RelativePath = node.RelativePath
        });

        if (metadataResponse.Success)
        {
            FillFileInfo(metadataResponse, ref FileInfo);
        }

        BytesTransferred = Length;
        return STATUS_SUCCESS;
    }

    public override int SetFileSize(object FileNode, object FileDesc, ulong NewSize, bool SetAllocationSize, out FspFileInfo FileInfo)
    {
        FileInfo = new FspFileInfo();

        if (!TryGetNode(FileNode, out var node))
        {
            return STATUS_INVALID_PARAMETER;
        }

        if (node.IsDirectory)
        {
            return STATUS_FILE_IS_A_DIRECTORY;
        }

        var response = SendRequest(new HostFileServiceRequest
        {
            Operation = "set-length",
            RelativePath = node.RelativePath,
            NewSize = checked((long)NewSize)
        });

        if (!response.Success)
        {
            return MapErrorCode(response.ErrorCode);
        }

        FillFileInfo(response, ref FileInfo);
        return STATUS_SUCCESS;
    }

    public override int SetBasicInfo(
        object FileNode,
        object FileDesc,
        uint FileAttributes,
        ulong CreationTime,
        ulong LastAccessTime,
        ulong LastWriteTime,
        ulong ChangeTime,
        out FspFileInfo FileInfo)
    {
        FileInfo = new FspFileInfo();

        if (!TryGetMetadataFromNode(FileNode, out var metadata, out var status))
        {
            return status;
        }

        FillFileInfo(metadata, ref FileInfo);
        return STATUS_SUCCESS;
    }

    public override int Flush(object FileNode, object FileDesc, out FspFileInfo FileInfo)
    {
        FileInfo = new FspFileInfo();

        if (!TryGetMetadataFromNode(FileNode, out var metadata, out var status))
        {
            if (status == STATUS_INVALID_PARAMETER)
            {
                return STATUS_SUCCESS;
            }

            return status;
        }

        FillFileInfo(metadata, ref FileInfo);
        return STATUS_SUCCESS;
    }

    public override int ReadDirectory(
        object FileNode,
        object FileDesc,
        string Pattern,
        string Marker,
        IntPtr Buffer,
        uint Length,
        out uint BytesTransferred)
    {
        return SeekableReadDirectory(FileNode, FileDesc, Pattern, Marker, Buffer, Length, out BytesTransferred);
    }

    public override bool ReadDirectoryEntry(
        object FileNode,
        object FileDesc,
        string Pattern,
        string Marker,
        ref object Context,
        out string FileName,
        out FspFileInfo FileInfo)
    {
        FileName = string.Empty;
        FileInfo = new FspFileInfo();

        if (!TryGetNode(FileNode, out var node) || !node.IsDirectory)
        {
            return false;
        }

        if (Context is not DirectoryEnumerationContext directoryContext)
        {
            if (TryGetCatalogRootEntries(node.RelativePath, out var catalogEntries))
            {
                var entries = catalogEntries
                    .OrderBy(static item => item.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (!string.IsNullOrWhiteSpace(Marker))
                {
                    entries = entries
                        .Where(item => string.Compare(item.Name, Marker, StringComparison.OrdinalIgnoreCase) > 0)
                        .ToList();
                }

                directoryContext = new DirectoryEnumerationContext(entries);
                Context = directoryContext;
            }
            else
            {
            var response = SendRequest(new HostFileServiceRequest
            {
                Operation = "list-directory",
                RelativePath = node.RelativePath
            });

            if (!response.Success)
            {
                return false;
            }

            var entries = response.Entries
                .Where(static item => item is not null)
                .OrderBy(static item => item.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (!string.IsNullOrWhiteSpace(Marker))
            {
                entries = entries
                    .Where(item => string.Compare(item.Name, Marker, StringComparison.OrdinalIgnoreCase) > 0)
                    .ToList();
            }

            directoryContext = new DirectoryEnumerationContext(entries);
            Context = directoryContext;
            }
        }

        if (!directoryContext.TryMoveNext(out var entry))
        {
            return false;
        }

        FileName = entry.Name;
        FillFileInfo(entry, ref FileInfo);
        return true;
    }

    public override int GetDirInfoByName(object FileNode, object FileDesc, string FileName, out string NormalizedName, out FspFileInfo FileInfo)
    {
        NormalizedName = "\\";
        FileInfo = new FspFileInfo();

        if (!TryGetNode(FileNode, out var node) || !node.IsDirectory)
        {
            return STATUS_NOT_A_DIRECTORY;
        }

        var childRelativePath = CombineRelativePath(node.RelativePath, FileName);
        if (!TryGetMetadata(childRelativePath, out var metadata, out var status))
        {
            return status;
        }

        NormalizedName = ToNormalizedName(childRelativePath);
        FillFileInfo(metadata, ref FileInfo);
        return STATUS_SUCCESS;
    }

    public override int CanDelete(object FileNode, object FileDesc, string FileName)
    {
        if (!TryGetNode(FileNode, out var node))
        {
            return STATUS_INVALID_PARAMETER;
        }

        if (!node.IsDirectory)
        {
            return STATUS_SUCCESS;
        }

        var response = SendRequest(new HostFileServiceRequest
        {
            Operation = "list-directory",
            RelativePath = node.RelativePath
        });

        if (!response.Success)
        {
            return MapErrorCode(response.ErrorCode);
        }

        return response.Entries.Count == 0
            ? STATUS_SUCCESS
            : STATUS_DIRECTORY_NOT_EMPTY;
    }

    public override int SetDelete(object FileNode, object FileDesc, string FileName, bool DeleteFile)
    {
        return STATUS_SUCCESS;
    }

    public override int Rename(object FileNode, object FileDesc, string FileName, string NewFileName, bool ReplaceIfExists)
    {
        if (!TryGetNode(FileNode, out var node))
        {
            return STATUS_INVALID_PARAMETER;
        }

        var source = NormalizeToRelativePath(FileName);
        var target = NormalizeToRelativePath(NewFileName);

        var response = SendRequest(new HostFileServiceRequest
        {
            Operation = "move",
            RelativePath = source,
            TargetRelativePath = target,
            Overwrite = ReplaceIfExists
        });

        if (!response.Success)
        {
            return MapErrorCode(response.ErrorCode);
        }

        node.RelativePath = target;
        return STATUS_SUCCESS;
    }

    public override void Cleanup(object FileNode, object FileDesc, string FileName, uint Flags)
    {
        if (!TryGetNode(FileNode, out var node))
        {
            return;
        }

        if ((Flags & CleanupDelete) == 0)
        {
            return;
        }

        var response = SendRequest(new HostFileServiceRequest
        {
            Operation = "delete",
            RelativePath = node.RelativePath,
            Recursive = false
        });

        if (!response.Success)
        {
            return;
        }
    }

    public override void Close(object FileNode, object FileDesc)
    {
    }

    public override int ExceptionHandler(Exception ex)
    {
        return ex switch
        {
            SocketException => STATUS_UNEXPECTED_NETWORK_ERROR,
            TimeoutException => STATUS_IO_TIMEOUT,
            DllNotFoundException => STATUS_DLL_NOT_FOUND,
            _ => STATUS_UNSUCCESSFUL
        };
    }

    private bool TryGetMetadataFromNode(object fileNode, out HostFileServiceResponse metadata, out int status)
    {
        metadata = new HostFileServiceResponse();
        status = STATUS_INVALID_PARAMETER;

        if (!TryGetNode(fileNode, out var node))
        {
            return false;
        }

        return TryGetMetadata(node.RelativePath, out metadata, out status);
    }

    private bool TryGetMetadata(string relativePath, out HostFileServiceResponse metadata, out int status)
    {
        if (TryGetVirtualMetadata(relativePath, out metadata, out status))
        {
            return true;
        }

        metadata = SendRequest(new HostFileServiceRequest
        {
            Operation = "metadata",
            RelativePath = relativePath
        });

        if (!metadata.Success)
        {
            status = MapErrorCode(metadata.ErrorCode);
            return false;
        }

        if (!metadata.Exists)
        {
            status = STATUS_OBJECT_NAME_NOT_FOUND;
            return false;
        }

        status = STATUS_SUCCESS;
        return true;
    }

    private HostFileServiceResponse SendRequest(HostFileServiceRequest request)
    {
        if (!TryResolveShareRequestPath(request.RelativePath, out var shareId, out var translatedRelativePath, out var statusCode))
        {
            return new HostFileServiceResponse
            {
                RequestId = request.RequestId,
                Success = false,
                ErrorCode = statusCode == STATUS_OBJECT_NAME_NOT_FOUND ? "not-found" : "bad-request",
                ErrorMessage = "Pfad konnte keinem Share zugeordnet werden."
            };
        }

        request.ShareId = shareId;
        request.RelativePath = translatedRelativePath;
        if (!string.IsNullOrWhiteSpace(request.TargetRelativePath)
            && TryResolveShareRequestPath(request.TargetRelativePath, out _, out var translatedTarget, out _))
        {
            request.TargetRelativePath = translatedTarget;
        }
        request.RequestId = string.IsNullOrWhiteSpace(request.RequestId)
            ? Guid.NewGuid().ToString("N")
            : request.RequestId;

        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
        var client = new HyperVSocketFileGuestClient();
        return client.SendAsync(request, cts.Token).GetAwaiter().GetResult();
    }

    private static bool TryGetNode(object fileNode, out RpcPathHandle node)
    {
        if (fileNode is RpcPathHandle typed)
        {
            node = typed;
            return true;
        }

        node = new RpcPathHandle(string.Empty, isDirectory: true);
        return false;
    }

    private static string NormalizeToRelativePath(string value)
    {
        var normalized = (value ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalized)
            || string.Equals(normalized, "\\", StringComparison.Ordinal))
        {
            return string.Empty;
        }

        normalized = normalized.Replace('/', '\\').TrimStart('\\');
        return normalized;
    }

    private static string CombineRelativePath(string parent, string child)
    {
        var normalizedChild = NormalizeToRelativePath(child);
        if (string.IsNullOrWhiteSpace(parent))
        {
            return normalizedChild;
        }

        if (string.IsNullOrWhiteSpace(normalizedChild))
        {
            return parent;
        }

        return $"{parent.TrimEnd('\\')}\\{normalizedChild}";
    }

    private static string ToNormalizedName(string relativePath)
    {
        if (string.IsNullOrWhiteSpace(relativePath))
        {
            return "\\";
        }

        return "\\" + relativePath.Replace('/', '\\').TrimStart('\\');
    }

    private static uint ToFileAttributes(bool isDirectory, bool readOnly)
    {
        var attributes = isDirectory
            ? FileAttributes.Directory
            : FileAttributes.Archive;

        if (readOnly)
        {
            attributes |= FileAttributes.ReadOnly;
        }

        return unchecked((uint)attributes);
    }

    private static void FillFileInfo(HostFileServiceResponse source, ref FspFileInfo target)
    {
        var writeTime = source.LastWriteTimeUtc ?? DateTimeOffset.UtcNow;

        target.FileAttributes = ToFileAttributes(source.IsDirectory, readOnly: false);
        target.AllocationSize = source.IsDirectory
            ? 0
            : unchecked((ulong)Math.Max(0, source.Size));
        target.FileSize = source.IsDirectory
            ? 0
            : unchecked((ulong)Math.Max(0, source.Size));
        target.CreationTime = unchecked((ulong)writeTime.UtcDateTime.ToFileTimeUtc());
        target.LastAccessTime = unchecked((ulong)writeTime.UtcDateTime.ToFileTimeUtc());
        target.LastWriteTime = unchecked((ulong)writeTime.UtcDateTime.ToFileTimeUtc());
        target.ChangeTime = unchecked((ulong)writeTime.UtcDateTime.ToFileTimeUtc());
        target.ReparseTag = 0;
        target.HardLinks = 0;
    }

    private static void FillFileInfo(HostFileServiceEntryInfo source, ref FspFileInfo target)
    {
        var pseudo = new HostFileServiceResponse
        {
            IsDirectory = source.IsDirectory,
            Size = source.Size,
            LastWriteTimeUtc = source.LastWriteTimeUtc
        };

        FillFileInfo(pseudo, ref target);
    }

    private static int MapErrorCode(string errorCode)
    {
        var normalized = (errorCode ?? string.Empty).Trim().ToLowerInvariant();
        return normalized switch
        {
            "not-found" => STATUS_OBJECT_NAME_NOT_FOUND,
            "share-not-found" => STATUS_OBJECT_PATH_NOT_FOUND,
            "share-path-missing" => STATUS_OBJECT_PATH_NOT_FOUND,
            "already-exists" => STATUS_OBJECT_NAME_COLLISION,
            "invalid-path" => STATUS_OBJECT_PATH_NOT_FOUND,
            "read-only" => STATUS_MEDIA_WRITE_PROTECTED,
            "unsupported-operation" => STATUS_NOT_IMPLEMENTED,
            "bad-request" => STATUS_INVALID_PARAMETER,
            _ => STATUS_UNSUCCESSFUL
        };
    }

    private static byte[] BuildDefaultSecurityDescriptor()
    {
        try
        {
            var tempPath = Path.GetTempPath();
            var security = new DirectoryInfo(tempPath).GetAccessControl();
            return security.GetSecurityDescriptorBinaryForm();
        }
        catch
        {
            return [];
        }
    }

    private bool TryResolveShareRequestPath(string relativePath, out string shareId, out string translatedRelativePath, out int status)
    {
        shareId = string.Empty;
        translatedRelativePath = string.Empty;
        status = STATUS_SUCCESS;

        var normalized = NormalizeToRelativePath(relativePath);
        var catalogSnapshot = GetCatalogSnapshot();
        if (catalogSnapshot.Count == 0)
        {
            if (string.IsNullOrWhiteSpace(_defaultShareId))
            {
                status = string.IsNullOrWhiteSpace(normalized)
                    ? STATUS_SUCCESS
                    : STATUS_OBJECT_NAME_NOT_FOUND;
                return false;
            }

            shareId = _defaultShareId;
            translatedRelativePath = normalized;
            return true;
        }

        if (string.IsNullOrWhiteSpace(normalized))
        {
            status = STATUS_OBJECT_NAME_NOT_FOUND;
            return false;
        }

        var separatorIndex = normalized.IndexOf('\\');
        var rootName = separatorIndex < 0 ? normalized : normalized[..separatorIndex];
        var childPath = separatorIndex < 0 ? string.Empty : normalized[(separatorIndex + 1)..];

        if (!catalogSnapshot.TryGetValue(rootName, out var mappedShareId))
        {
            status = STATUS_OBJECT_NAME_NOT_FOUND;
            return false;
        }

        shareId = mappedShareId;
        translatedRelativePath = childPath;
        return true;
    }

    private bool TryGetVirtualMetadata(string relativePath, out HostFileServiceResponse metadata, out int status)
    {
        metadata = new HostFileServiceResponse();
        status = STATUS_SUCCESS;

        var catalogSnapshot = GetCatalogSnapshot();
        var normalized = NormalizeToRelativePath(relativePath);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            metadata.Success = true;
            metadata.Exists = true;
            metadata.IsDirectory = true;
            metadata.Size = 0;
            metadata.LastWriteTimeUtc = DateTimeOffset.UtcNow;
            return true;
        }

        if (catalogSnapshot.ContainsKey(normalized))
        {
            metadata.Success = true;
            metadata.Exists = true;
            metadata.IsDirectory = true;
            metadata.Size = 0;
            metadata.LastWriteTimeUtc = DateTimeOffset.UtcNow;
            return true;
        }

        return false;
    }

    private bool TryGetCatalogRootEntries(string nodeRelativePath, out List<HostFileServiceEntryInfo> entries)
    {
        entries = [];
        var catalogSnapshot = GetCatalogSnapshot();

        var normalized = NormalizeToRelativePath(nodeRelativePath);
        if (!string.IsNullOrWhiteSpace(normalized))
        {
            return false;
        }

        entries = catalogSnapshot.Keys
            .OrderBy(static key => key, StringComparer.OrdinalIgnoreCase)
            .Select(static key => new HostFileServiceEntryInfo
            {
                Name = key,
                RelativePath = key,
                IsDirectory = true,
                Size = 0,
                LastWriteTimeUtc = DateTimeOffset.UtcNow
            })
            .ToList();

        return true;
    }

    private Dictionary<string, string> GetCatalogSnapshot()
    {
        lock (_catalogSync)
        {
            return new Dictionary<string, string>(_catalogByRoot, StringComparer.OrdinalIgnoreCase);
        }
    }

    private sealed class RpcPathHandle
    {
        public RpcPathHandle(string relativePath, bool isDirectory)
        {
            RelativePath = relativePath ?? string.Empty;
            IsDirectory = isDirectory;
        }

        public string RelativePath { get; set; }

        public bool IsDirectory { get; }
    }

    private sealed class DirectoryEnumerationContext
    {
        private readonly IReadOnlyList<HostFileServiceEntryInfo> _entries;
        private int _index;

        public DirectoryEnumerationContext(IReadOnlyList<HostFileServiceEntryInfo> entries)
        {
            _entries = entries;
            _index = 0;
        }

        public bool TryMoveNext(out HostFileServiceEntryInfo entry)
        {
            if (_index >= _entries.Count)
            {
                entry = new HostFileServiceEntryInfo();
                return false;
            }

            entry = _entries[_index];
            _index++;
            return true;
        }
    }
}
