using Microsoft.Win32;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;

namespace HyperTool.Services;

public sealed class HyperVSocketDiagnosticsHostListener : IDisposable
{
    private static readonly JsonSerializerOptions AckJsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private readonly Guid _serviceId;
    private readonly Action<HyperVSocketDiagnosticsAck> _onDiagnosticsAck;
    private Socket? _listener;
    private CancellationTokenSource? _cts;
    private Task? _acceptLoopTask;

    public HyperVSocketDiagnosticsHostListener(Action<HyperVSocketDiagnosticsAck> onDiagnosticsAck, Guid? serviceId = null)
    {
        _onDiagnosticsAck = onDiagnosticsAck ?? throw new ArgumentNullException(nameof(onDiagnosticsAck));
        _serviceId = serviceId ?? HyperVSocketUsbTunnelDefaults.DiagnosticsServiceId;
    }

    public bool IsRunning { get; private set; }

    public void Start()
    {
        if (IsRunning)
        {
            return;
        }

        TryRegisterServiceGuid();

        var listener = new Socket((AddressFamily)34, SocketType.Stream, (ProtocolType)1);
        listener.Bind(new HyperVSocketEndPoint(HyperVSocketUsbTunnelDefaults.VmIdWildcard, _serviceId));
        listener.Listen(16);

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
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 512, leaveOpen: false);

        string? line;
        try
        {
            line = await reader.ReadLineAsync(cancellationToken);
        }
        catch
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(line))
        {
            return;
        }

        var payload = line.Trim();
        if (payload.Length == 0)
        {
            return;
        }

        var ack = ParseAckPayload(payload);
        _onDiagnosticsAck(ack);
    }

    private static HyperVSocketDiagnosticsAck ParseAckPayload(string payload)
    {
        var normalizedPayload = payload.TrimStart('\uFEFF', ' ', '\t', '\r', '\n');

        try
        {
            var parsed = JsonSerializer.Deserialize<HyperVSocketDiagnosticsAck>(normalizedPayload, AckJsonOptions);
            if (parsed is not null && !string.IsNullOrWhiteSpace(parsed.GuestComputerName))
            {
                return parsed;
            }
        }
        catch
        {
        }

        try
        {
            using var doc = JsonDocument.Parse(normalizedPayload);
            if (doc.RootElement.ValueKind == JsonValueKind.Object)
            {
                var root = doc.RootElement;
                var guestComputerName = GetJsonString(root, "guestComputerName");
                if (!string.IsNullOrWhiteSpace(guestComputerName))
                {
                    return new HyperVSocketDiagnosticsAck
                    {
                        GuestComputerName = guestComputerName,
                        HyperVSocketActive = GetJsonBool(root, "hyperVSocketActive"),
                        RegistryServiceOk = GetJsonBool(root, "registryServiceOk"),
                        BusId = GetJsonString(root, "busId"),
                        EventType = GetJsonString(root, "eventType"),
                        SentAtUtc = GetJsonString(root, "sentAtUtc")
                    };
                }
            }
        }
        catch
        {
        }

        return new HyperVSocketDiagnosticsAck
        {
            GuestComputerName = normalizedPayload,
            HyperVSocketActive = null,
            RegistryServiceOk = null,
            SentAtUtc = null
        };
    }

    private static string? GetJsonString(JsonElement root, string propertyName)
    {
        foreach (var property in root.EnumerateObject())
        {
            if (!string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            return property.Value.ValueKind switch
            {
                JsonValueKind.String => property.Value.GetString(),
                JsonValueKind.Null => null,
                _ => property.Value.ToString()
            };
        }

        return null;
    }

    private static bool? GetJsonBool(JsonElement root, string propertyName)
    {
        foreach (var property in root.EnumerateObject())
        {
            if (!string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (property.Value.ValueKind == JsonValueKind.True)
            {
                return true;
            }

            if (property.Value.ValueKind == JsonValueKind.False)
            {
                return false;
            }

            if (property.Value.ValueKind == JsonValueKind.String
                && bool.TryParse(property.Value.GetString(), out var parsed))
            {
                return parsed;
            }
        }

        return null;
    }

    private void TryRegisterServiceGuid()
    {
        try
        {
            const string rootPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices";

            using var rootKey = Registry.LocalMachine.CreateSubKey(rootPath, writable: true);
            if (rootKey is null)
            {
                return;
            }

            using var serviceKey = rootKey.CreateSubKey(_serviceId.ToString("D"), writable: true);
            serviceKey?.SetValue("ElementName", "HyperTool Hyper-V Socket Diagnostics", RegistryValueKind.String);
        }
        catch
        {
        }
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

public sealed class HyperVSocketDiagnosticsAck
{
    public string GuestComputerName { get; set; } = string.Empty;

    public bool? HyperVSocketActive { get; set; }

    public bool? RegistryServiceOk { get; set; }

    public string? BusId { get; set; }

    public string? EventType { get; set; }

    public string? SentAtUtc { get; set; }
}
