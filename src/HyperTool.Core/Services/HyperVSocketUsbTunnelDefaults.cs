namespace HyperTool.Services;

public static class HyperVSocketUsbTunnelDefaults
{
    public const string ServiceIdString = "6c4eb1be-40e8-4c8b-a4d6-5b0f67d7e40f";
    public static readonly Guid ServiceId = Guid.Parse(ServiceIdString);
    public const string DiagnosticsServiceIdString = "67c53bca-3f3d-4628-98e4-e45be5d6d1ad";
    public static readonly Guid DiagnosticsServiceId = Guid.Parse(DiagnosticsServiceIdString);
    public const string SharedFolderCatalogServiceIdString = "e7db04df-0e32-4f30-a4dc-c6cbc31a8792";
    public static readonly Guid SharedFolderCatalogServiceId = Guid.Parse(SharedFolderCatalogServiceIdString);
    public const string HostIdentityServiceIdString = "54b2c423-6f79-47d8-a77d-8cab14e3f041";
    public static readonly Guid HostIdentityServiceId = Guid.Parse(HostIdentityServiceIdString);
    public const string FileServiceIdString = "91df7cec-c5ba-452a-b072-42e5f672d5f9";
    public static readonly Guid FileServiceId = Guid.Parse(FileServiceIdString);
    public static readonly Guid VmIdParent = Guid.Parse("a42e7cda-d03f-480c-9cc2-a4de20abb878");
    public static readonly Guid VmIdWildcard = Guid.Empty;
    public const int UsbIpTcpPort = 3240;
    public const string LoopbackAddress = "127.0.0.1";
}