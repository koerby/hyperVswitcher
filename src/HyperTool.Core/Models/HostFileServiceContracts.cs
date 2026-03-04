using System.Text.Json.Serialization;

namespace HyperTool.Models;

public sealed class HostFileServiceRequest
{
    public string RequestId { get; set; } = Guid.NewGuid().ToString("N");

    public string Operation { get; set; } = "ping";

    public string ShareId { get; set; } = string.Empty;

    public string RelativePath { get; set; } = string.Empty;

    public string TargetRelativePath { get; set; } = string.Empty;

    public long Offset { get; set; }

    public int Length { get; set; }

    public long NewSize { get; set; } = -1;

    public bool Recursive { get; set; }

    public bool Overwrite { get; set; }

    public bool Truncate { get; set; }

    public bool Append { get; set; }

    public string DataBase64 { get; set; } = string.Empty;
}

public sealed class HostFileServiceResponse
{
    public string RequestId { get; set; } = string.Empty;

    public bool Success { get; set; }

    public string ErrorCode { get; set; } = string.Empty;

    public string ErrorMessage { get; set; } = string.Empty;

    public string ShareId { get; set; } = string.Empty;

    public string RelativePath { get; set; } = string.Empty;

    public bool Exists { get; set; }

    public bool IsDirectory { get; set; }

    public long Size { get; set; }

    public DateTimeOffset? LastWriteTimeUtc { get; set; }

    public int BytesRead { get; set; }

    public string DataBase64 { get; set; } = string.Empty;

    public List<HostFileServiceShareInfo> Shares { get; set; } = [];

    public List<HostFileServiceEntryInfo> Entries { get; set; } = [];
}

public sealed class HostFileServiceShareInfo
{
    public string Id { get; set; } = string.Empty;

    public string Label { get; set; } = string.Empty;

    public string ShareName { get; set; } = string.Empty;

    public bool ReadOnly { get; set; }
}

public sealed class HostFileServiceEntryInfo
{
    public string Name { get; set; } = string.Empty;

    public string RelativePath { get; set; } = string.Empty;

    public bool IsDirectory { get; set; }

    public long Size { get; set; }

    public DateTimeOffset? LastWriteTimeUtc { get; set; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool ReadOnly { get; set; }
}
