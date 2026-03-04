using Microsoft.Win32;
using HyperTool.Models;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;

namespace HyperTool.Services;

public sealed class HyperVSocketHostIdentityHostListener : IDisposable
{
    private readonly Guid _serviceId;
    private readonly Func<HostFeatureAvailability>? _featureAvailabilityProvider;
    private Socket? _listener;
    private CancellationTokenSource? _cts;
    private Task? _acceptLoopTask;

    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public HyperVSocketHostIdentityHostListener(Guid? serviceId = null, Func<HostFeatureAvailability>? featureAvailabilityProvider = null)
    {
        _serviceId = serviceId ?? HyperVSocketUsbTunnelDefaults.HostIdentityServiceId;
        _featureAvailabilityProvider = featureAvailabilityProvider;
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
            throw new InvalidOperationException(
                "Hyper-V Socket Host-Identity-Dienst ist nicht registriert. Starte HyperTool Host als Administrator, um den Dienst einmalig zu registrieren.");
        }

        var listener = new Socket((AddressFamily)34, SocketType.Stream, (ProtocolType)1);
        listener.Bind(new HyperVSocketEndPoint(HyperVSocketUsbTunnelDefaults.VmIdWildcard, _serviceId));
        listener.Listen(16);

        _listener = listener;
        _cts = new CancellationTokenSource();
        _acceptLoopTask = Task.Run(() => AcceptLoopAsync(_cts.Token));
        IsRunning = true;
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

        HostFeatureAvailability featureAvailability;
        try
        {
            featureAvailability = _featureAvailabilityProvider?.Invoke() ?? new HostFeatureAvailability();
        }
        catch
        {
            featureAvailability = new HostFeatureAvailability();
        }

        var payload = JsonSerializer.Serialize(new
        {
            hostName = Environment.MachineName,
            fqdn = ResolveFqdn(),
            features = new
            {
                usbSharingEnabled = featureAvailability.UsbSharingEnabled,
                sharedFoldersEnabled = featureAvailability.SharedFoldersEnabled
            },
            timestampUtc = DateTime.UtcNow
        }, SerializerOptions);

        await using var writer = new StreamWriter(stream, Encoding.UTF8, bufferSize: 1024, leaveOpen: false)
        {
            NewLine = "\n"
        };

        await writer.WriteLineAsync(payload.AsMemory(), cancellationToken);
        await writer.FlushAsync(cancellationToken);
    }

    private static string ResolveFqdn()
    {
        try
        {
            var host = Dns.GetHostEntry(Environment.MachineName);
            return host.HostName ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
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
            serviceKey?.SetValue("ElementName", "HyperTool Hyper-V Socket Host Identity", RegistryValueKind.String);
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
