namespace HyperTool.Models;

public sealed class UsbIpDeviceInfo
{
    public string BusId { get; set; } = string.Empty;

    public string Description { get; set; } = string.Empty;

    public string HardwareId { get; set; } = string.Empty;

    public string InstanceId { get; set; } = string.Empty;

    public string PersistedGuid { get; set; } = string.Empty;

    public string ClientIpAddress { get; set; } = string.Empty;

    public bool IsConnected => !string.IsNullOrWhiteSpace(InstanceId);

    public bool IsRemoteAvailable => !string.IsNullOrWhiteSpace(BusId)
        && string.IsNullOrWhiteSpace(InstanceId)
        && string.IsNullOrWhiteSpace(PersistedGuid)
        && string.IsNullOrWhiteSpace(ClientIpAddress);

    public bool IsShared => !string.IsNullOrWhiteSpace(PersistedGuid);

    public bool IsAttached => !string.IsNullOrWhiteSpace(ClientIpAddress);

    public string StateText
    {
        get
        {
            if (!IsConnected && IsShared)
            {
                return "Persisted";
            }

            if (IsAttached)
            {
                return "Attached";
            }

            if (IsShared)
            {
                return "Shared";
            }

            if (IsRemoteAvailable)
            {
                return "Available";
            }

            if (IsConnected)
            {
                return "Connected";
            }

            return "Unknown";
        }
    }

    public string DisplayName
    {
        get
        {
            var bus = string.IsNullOrWhiteSpace(BusId) ? "-" : BusId;
            var hardware = string.IsNullOrWhiteSpace(HardwareId) ? "----:----" : HardwareId;
            var description = string.IsNullOrWhiteSpace(Description) ? "USB Device" : Description;
            return $"{bus}  {hardware}  {description}  [{StateText}]";
        }
    }
}
