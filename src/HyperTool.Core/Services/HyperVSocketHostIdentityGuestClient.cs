using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using HyperTool.Models;

namespace HyperTool.Services;

public sealed class HyperVSocketHostIdentityGuestClient
{
    private readonly Guid _serviceId;

    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private sealed class HostIdentityPayload
    {
        public string HostName { get; set; } = string.Empty;
        public string Fqdn { get; set; } = string.Empty;
        public HostFeatureAvailability? Features { get; set; }
    }

    public HyperVSocketHostIdentityGuestClient(Guid? serviceId = null)
    {
        _serviceId = serviceId ?? HyperVSocketUsbTunnelDefaults.HostIdentityServiceId;
    }

    public async Task<string?> FetchHostNameAsync(CancellationToken cancellationToken)
    {
        var identity = await FetchHostIdentityAsync(cancellationToken);
        if (identity is null)
        {
            return null;
        }

        if (!string.IsNullOrWhiteSpace(identity.HostName))
        {
            return identity.HostName.Trim();
        }

        if (!string.IsNullOrWhiteSpace(identity.Fqdn))
        {
            return identity.Fqdn.Trim();
        }

        return null;
    }

    public async Task<HostIdentityInfo?> FetchHostIdentityAsync(CancellationToken cancellationToken)
    {
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        linkedCts.CancelAfter(TimeSpan.FromMilliseconds(2000));

        using var socket = new Socket((AddressFamily)34, SocketType.Stream, (ProtocolType)1);
        linkedCts.Token.ThrowIfCancellationRequested();
        socket.Connect(new HyperVSocketEndPoint(HyperVSocketUsbTunnelDefaults.VmIdParent, _serviceId));

        await using var stream = new NetworkStream(socket, ownsSocket: true);
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: false);

        var payloadText = await reader.ReadLineAsync(linkedCts.Token);
        if (string.IsNullOrWhiteSpace(payloadText))
        {
            return null;
        }

        var payload = JsonSerializer.Deserialize<HostIdentityPayload>(payloadText, SerializerOptions);
        if (payload is null)
        {
            return null;
        }

        var hostName = payload.HostName?.Trim() ?? string.Empty;
        var fqdn = payload.Fqdn?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(hostName) && string.IsNullOrWhiteSpace(fqdn))
        {
            return null;
        }

        return new HostIdentityInfo
        {
            HostName = hostName,
            Fqdn = fqdn,
            Features = payload.Features ?? new HostFeatureAvailability()
        };
    }
}
