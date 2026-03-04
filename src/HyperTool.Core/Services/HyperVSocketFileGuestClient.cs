using HyperTool.Models;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;

namespace HyperTool.Services;

public sealed class HyperVSocketFileGuestClient
{
    private readonly Guid _serviceId;

    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true
    };

    public HyperVSocketFileGuestClient(Guid? serviceId = null)
    {
        _serviceId = serviceId ?? HyperVSocketUsbTunnelDefaults.FileServiceId;
    }

    public Task<HostFileServiceResponse> PingAsync(CancellationToken cancellationToken)
    {
        return SendAsync(new HostFileServiceRequest
        {
            Operation = "ping"
        }, cancellationToken);
    }

    public Task<HostFileServiceResponse> ListSharesAsync(CancellationToken cancellationToken)
    {
        return SendAsync(new HostFileServiceRequest
        {
            Operation = "list-shares"
        }, cancellationToken);
    }

    public Task<HostFileServiceResponse> SendAsync(HostFileServiceRequest request, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (string.IsNullOrWhiteSpace(request.RequestId))
        {
            request.RequestId = Guid.NewGuid().ToString("N");
        }

        return SendCoreAsync(request, cancellationToken);
    }

    private async Task<HostFileServiceResponse> SendCoreAsync(HostFileServiceRequest request, CancellationToken cancellationToken)
    {
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        linkedCts.CancelAfter(TimeSpan.FromMilliseconds(6000));

        using var socket = new Socket((AddressFamily)34, SocketType.Stream, (ProtocolType)1);
        linkedCts.Token.ThrowIfCancellationRequested();
        socket.Connect(new HyperVSocketEndPoint(HyperVSocketUsbTunnelDefaults.VmIdParent, _serviceId));

        await using var stream = new NetworkStream(socket, ownsSocket: true);
        await using var writer = new StreamWriter(stream, Encoding.UTF8, bufferSize: 16 * 1024, leaveOpen: true)
        {
            NewLine = "\n"
        };
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 16 * 1024, leaveOpen: false);

        var payload = JsonSerializer.Serialize(request, SerializerOptions);
        await writer.WriteLineAsync(payload.AsMemory(), linkedCts.Token);
        await writer.FlushAsync(linkedCts.Token);

        var responsePayload = await reader.ReadLineAsync(linkedCts.Token);
        if (string.IsNullOrWhiteSpace(responsePayload))
        {
            throw new InvalidOperationException("Leere Antwort vom HyperTool File-Dienst.");
        }

        var response = JsonSerializer.Deserialize<HostFileServiceResponse>(responsePayload, SerializerOptions)
            ?? new HostFileServiceResponse();

        response.RequestId = string.IsNullOrWhiteSpace(response.RequestId)
            ? request.RequestId
            : response.RequestId;

        return response;
    }
}
