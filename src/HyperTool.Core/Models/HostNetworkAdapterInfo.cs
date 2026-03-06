namespace HyperTool.Models;

public sealed class HostNetworkAdapterInfo
{
    public string AdapterName { get; set; } = string.Empty;

    public string InterfaceDescription { get; set; } = string.Empty;

    public string IpAddresses { get; set; } = string.Empty;

    public string Subnets { get; set; } = string.Empty;

    public string Gateway { get; set; } = string.Empty;

    public string DnsServers { get; set; } = string.Empty;

    public string NetworkProfileCategory { get; set; } = string.Empty;

    public bool HasGateway => !string.IsNullOrWhiteSpace(Gateway);

    public bool IsDefaultSwitch { get; set; }
}