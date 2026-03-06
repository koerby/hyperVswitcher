using HyperTool.Models;
using Serilog;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Json;

namespace HyperTool.Services;

public sealed class HyperVPowerShellService : IHyperVService
{
    public async Task<IReadOnlyList<HyperVVmInfo>> GetVmsAsync(CancellationToken cancellationToken)
    {
        const string script = """
            @(
                Get-VM | ForEach-Object {
                    $adapter = Get-VMNetworkAdapter -VMName $_.Name -ErrorAction SilentlyContinue | Select-Object -First 1

                    [pscustomobject]@{
                        Name = $_.Name
                        State = $_.State.ToString()
                        Status = $_.Status
                        CurrentSwitchName = if ($null -ne $adapter -and $null -ne $adapter.SwitchName) { $adapter.SwitchName } else { '' }
                    }
                }
            ) | ConvertTo-Json -Depth 4 -Compress
            """;

        var rows = await InvokeJsonArrayAsync(script, cancellationToken);
        return rows.Select(row => new HyperVVmInfo
        {
            Name = GetString(row, "Name"),
            State = GetString(row, "State"),
            Status = GetString(row, "Status"),
            CurrentSwitchName = GetString(row, "CurrentSwitchName")
        }).ToList();
    }

    public Task StartVmAsync(string vmName, CancellationToken cancellationToken) =>
        InvokeNonQueryAsync($"Start-VM -VMName {ToPsSingleQuoted(vmName)} -Confirm:$false", cancellationToken);

    public Task StopVmGracefulAsync(string vmName, CancellationToken cancellationToken) =>
        InvokeNonQueryAsync($"Stop-VM -VMName {ToPsSingleQuoted(vmName)} -Confirm:$false", cancellationToken);

    public Task TurnOffVmAsync(string vmName, CancellationToken cancellationToken) =>
        InvokeNonQueryAsync($"Stop-VM -VMName {ToPsSingleQuoted(vmName)} -TurnOff -Confirm:$false", cancellationToken);

    public Task RestartVmAsync(string vmName, CancellationToken cancellationToken) =>
        InvokeNonQueryAsync($"Restart-VM -VMName {ToPsSingleQuoted(vmName)} -Force -Confirm:$false", cancellationToken);

    public async Task<IReadOnlyList<HyperVSwitchInfo>> GetVmSwitchesAsync(CancellationToken cancellationToken)
    {
        const string script = """
            @(Get-VMSwitch | Select-Object Name, SwitchType) | ConvertTo-Json -Depth 3 -Compress
            """;

        var rows = await InvokeJsonArrayAsync(script, cancellationToken);
        return rows.Select(row => new HyperVSwitchInfo
        {
            Name = GetString(row, "Name"),
            SwitchType = GetString(row, "SwitchType")
        }).ToList();
    }

    public async Task<IReadOnlyList<HyperVVmNetworkAdapterInfo>> GetVmNetworkAdaptersAsync(string vmName, CancellationToken cancellationToken)
    {
        var script =
            $"@(Get-VMNetworkAdapter -VMName {ToPsSingleQuoted(vmName)} -ErrorAction SilentlyContinue | " +
            "ForEach-Object { [pscustomobject]@{ Name = if ($null -ne $_.Name) { $_.Name } else { '' }; " +
            "SwitchName = if ($null -ne $_.SwitchName) { $_.SwitchName } else { '' }; " +
            "MacAddress = if ($null -ne $_.MacAddress) { $_.MacAddress } else { '' } } }) | ConvertTo-Json -Depth 4 -Compress";

        var rows = await InvokeJsonArrayAsync(script, cancellationToken);
        return rows.Select(row => new HyperVVmNetworkAdapterInfo
        {
            Name = GetString(row, "Name"),
            SwitchName = GetString(row, "SwitchName"),
            MacAddress = GetString(row, "MacAddress")
        }).ToList();
    }

    public async Task<string?> GetVmCurrentSwitchNameAsync(string vmName, CancellationToken cancellationToken)
    {
        var script = $"$adapter = Get-VMNetworkAdapter -VMName {ToPsSingleQuoted(vmName)} -ErrorAction SilentlyContinue | Select-Object -First 1; if ($null -eq $adapter -or $null -eq $adapter.SwitchName) {{ '' }} else {{ $adapter.SwitchName }}";

        var value = await InvokePowerShellAsync(script, cancellationToken);
        return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
    }

    public Task ConnectVmNetworkAdapterAsync(string vmName, string switchName, string? adapterName, CancellationToken cancellationToken)
    {
        var script = string.IsNullOrWhiteSpace(adapterName)
            ? $"Connect-VMNetworkAdapter -VMName {ToPsSingleQuoted(vmName)} -SwitchName {ToPsSingleQuoted(switchName)}"
            : $"Connect-VMNetworkAdapter -VMName {ToPsSingleQuoted(vmName)} -Name {ToPsSingleQuoted(adapterName)} -SwitchName {ToPsSingleQuoted(switchName)}";

        return InvokeNonQueryAsync(script, cancellationToken);
    }

    public Task DisconnectVmNetworkAdapterAsync(string vmName, string? adapterName, CancellationToken cancellationToken)
    {
        var script = string.IsNullOrWhiteSpace(adapterName)
            ? $"Disconnect-VMNetworkAdapter -VMName {ToPsSingleQuoted(vmName)}"
            : $"Disconnect-VMNetworkAdapter -VMName {ToPsSingleQuoted(vmName)} -Name {ToPsSingleQuoted(adapterName)}";

        return InvokeNonQueryAsync(script, cancellationToken);
    }

    public Task RenameVmNetworkAdapterAsync(string vmName, string adapterName, string newAdapterName, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(adapterName))
        {
            throw new ArgumentException("Adaptername darf nicht leer sein.", nameof(adapterName));
        }

        if (string.IsNullOrWhiteSpace(newAdapterName))
        {
            throw new ArgumentException("Neuer Adaptername darf nicht leer sein.", nameof(newAdapterName));
        }

        var script = $"Rename-VMNetworkAdapter -VMName {ToPsSingleQuoted(vmName)} -Name {ToPsSingleQuoted(adapterName)} -NewName {ToPsSingleQuoted(newAdapterName)}";
        return InvokeNonQueryAsync(script, cancellationToken);
    }

    public async Task<string> GetHostNetworkProfileCategoryAsync(CancellationToken cancellationToken)
    {
        const string script = """
            $profiles = @(
                Get-NetConnectionProfile -ErrorAction SilentlyContinue |
                    Where-Object {
                        ($_.IPv4Connectivity -ne 'Disconnected') -or ($_.IPv6Connectivity -ne 'Disconnected')
                    }
            )

            if ($profiles.Count -eq 0)
            {
                $profiles = @(Get-NetConnectionProfile -ErrorAction SilentlyContinue)
            }

            $categories = @($profiles | ForEach-Object {
                if ($null -ne $_.NetworkCategory) { $_.NetworkCategory.ToString() } else { '' }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

            if ($categories -contains 'Public')
            {
                'Public'
                return
            }

            if ($categories -contains 'DomainAuthenticated')
            {
                'DomainAuthenticated'
                return
            }

            if ($categories -contains 'Private')
            {
                'Private'
                return
            }

            if ($categories.Count -gt 0)
            {
                $categories[0]
                return
            }

            'Unknown'
            """;

        var output = await InvokePowerShellAsync(script, cancellationToken);
        return string.IsNullOrWhiteSpace(output) ? "Unknown" : output.Trim();
    }

    public async Task SetHostNetworkProfileCategoryAsync(string adapterName, string networkCategory, CancellationToken cancellationToken)
    {
        var script =
            $"$adapterName = {ToPsSingleQuoted(adapterName ?? string.Empty)}; " +
            $"$networkCategory = {ToPsSingleQuoted(networkCategory ?? string.Empty)}; " +
            "$networkCategory = if ([string]::IsNullOrWhiteSpace($networkCategory)) { '' } else { $networkCategory.Trim() }; " +
            "if ($networkCategory -notin @('Public','Private')) { throw \"Ungültige Netzprofil-Kategorie. Erlaubt: Public oder Private.\" }; " +
            "$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator); " +
            "if (-not $isAdmin) { throw \"Zum Ändern des Host-Netzprofils sind Administratorrechte erforderlich.\" }; " +
            "$profiles = @(); " +
            "if (-not [string]::IsNullOrWhiteSpace($adapterName)) { $profiles = @(Get-NetConnectionProfile -InterfaceAlias $adapterName -ErrorAction SilentlyContinue) }; " +
            "if ($profiles.Count -eq 0 -and [string]::IsNullOrWhiteSpace($adapterName)) { " +
            "  $profiles = @(Get-NetConnectionProfile -ErrorAction SilentlyContinue | Where-Object { ($_.IPv4Connectivity -ne 'Disconnected') -or ($_.IPv6Connectivity -ne 'Disconnected') }); " +
            "  if ($profiles.Count -eq 0) { $profiles = @(Get-NetConnectionProfile -ErrorAction SilentlyContinue) }; " +
            "}; " +
            "if ($profiles.Count -eq 0) { throw \"Kein passendes Netzprofil gefunden.\" }; " +
            "$domainSkipped = 0; " +
            "$changed = 0; " +
            "foreach ($profile in $profiles) { " +
            "  if ($null -eq $profile.InterfaceIndex) { continue }; " +
            "  $currentCategory = if ($null -ne $profile.NetworkCategory) { $profile.NetworkCategory.ToString() } else { '' }; " +
            "  if ($currentCategory -eq 'DomainAuthenticated') { $domainSkipped++; continue }; " +
            "  if ($currentCategory -eq $networkCategory) { continue }; " +
            "  try { " +
            "    Set-NetConnectionProfile -InterfaceIndex $profile.InterfaceIndex -NetworkCategory $networkCategory -ErrorAction Stop; " +
            "    $changed++; " +
            "  } catch { " +
            "    $message = $_.Exception.Message; " +
            "    if ($message -match 'Network List Manager Policies|Group Policy') { throw \"Netzprofiländerung ist durch Gruppenrichtlinie blockiert.\" }; " +
            "    throw; " +
            "  }; " +
            "}; " +
            "if ($changed -eq 0 -and $domainSkipped -gt 0) { throw \"Domänenprofile können nicht manuell auf Privat/Öffentlich umgestellt werden.\" }; " +
            "Write-Output 'HT_OK'";

        try
        {
            await InvokePowerShellElevatedNonQueryAsync(script, cancellationToken);
        }
        catch (UnauthorizedAccessException ex) when (ex.Message.Contains("UAC", StringComparison.OrdinalIgnoreCase))
        {
            throw new UnauthorizedAccessException("UAC-Bestätigung abgebrochen. Netzprofil wurde nicht geändert.");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("not running PowerShell elevated", StringComparison.OrdinalIgnoreCase)
                                                 || ex.Message.Contains("Administratorrechte", StringComparison.OrdinalIgnoreCase)
                                                 || ex.Message.Contains("PermissionDenied", StringComparison.OrdinalIgnoreCase))
        {
            throw new UnauthorizedAccessException("Administratorrechte wurden nicht erteilt. Netzprofil wurde nicht geändert.");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("DomainAuthenticated", StringComparison.OrdinalIgnoreCase)
                                                 || ex.Message.Contains("Domänenprofile", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Domänenprofile können nicht manuell geändert werden.");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("Network List Manager Policies", StringComparison.OrdinalIgnoreCase)
                                                 || ex.Message.Contains("Group Policy", StringComparison.OrdinalIgnoreCase)
                                                 || ex.Message.Contains("Gruppenrichtlinie", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Netzprofiländerung ist durch Gruppenrichtlinie blockiert.");
        }
    }

    public async Task<IReadOnlyList<HostNetworkAdapterInfo>> GetHostNetworkAdaptersWithUplinkAsync(CancellationToken cancellationToken)
    {
        const string script = """
            function Resolve-NetworkCategory {
                param(
                    [string]$InterfaceAlias,
                    [int]$InterfaceIndex
                )

                $profiles = @()

                if (-not [string]::IsNullOrWhiteSpace($InterfaceAlias))
                {
                    $profiles = @(Get-NetConnectionProfile -InterfaceAlias $InterfaceAlias -ErrorAction SilentlyContinue)
                }

                if ($profiles.Count -eq 0 -and $InterfaceIndex -gt 0)
                {
                    $profiles = @(Get-NetConnectionProfile -InterfaceIndex $InterfaceIndex -ErrorAction SilentlyContinue)
                }

                $categories = @($profiles | ForEach-Object {
                    if ($null -ne $_.NetworkCategory) { $_.NetworkCategory.ToString() } else { '' }
                } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

                if ($categories -contains 'Public') { return 'Public' }
                if ($categories -contains 'DomainAuthenticated') { return 'DomainAuthenticated' }
                if ($categories -contains 'Private') { return 'Private' }

                if ($categories.Count -gt 0)
                {
                    return $categories[0]
                }

                return ''
            }

            $items = @(
                Get-NetIPConfiguration -Detailed -ErrorAction SilentlyContinue |
                    Where-Object {
                        $adapter = $_.NetAdapter
                        if ($null -eq $adapter) { return $false }

                        $statusText = if ($null -ne $adapter.Status) { $adapter.Status.ToString() } else { '' }
                        if ([string]::IsNullOrWhiteSpace($statusText)) { $statusText = if ($null -ne $adapter.AdminStatus) { $adapter.AdminStatus.ToString() } else { '' } }
                        if ($statusText -match 'Disabled|Not Present') { return $false }

                        return $true
                    } |
                    ForEach-Object {
                        $adapter = $_.NetAdapter
                        $ipv4Addresses = @($_.IPv4Address | Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace($_.IPAddress) })
                        $ipv6Addresses = @($_.IPv6Address | Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace($_.IPAddress) -and $_.IPAddress -notlike 'fe80:*' })
                        $ipAddresses = @($ipv4Addresses | ForEach-Object { $_.IPAddress }) + @($ipv6Addresses | ForEach-Object { $_.IPAddress })
                        $prefixes = @($ipv4Addresses | ForEach-Object { '/' + $_.PrefixLength }) + @($ipv6Addresses | ForEach-Object { '/' + $_.PrefixLength })
                        $dnsServers = @($_.DnsServer.ServerAddresses | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

                        $gatewayCandidates = @()
                        if ($null -ne $_.IPv4DefaultGateway -and -not [string]::IsNullOrWhiteSpace($_.IPv4DefaultGateway.NextHop)) { $gatewayCandidates += $_.IPv4DefaultGateway.NextHop }
                        if ($null -ne $_.IPv6DefaultGateway -and -not [string]::IsNullOrWhiteSpace($_.IPv6DefaultGateway.NextHop)) { $gatewayCandidates += $_.IPv6DefaultGateway.NextHop }
                        $gatewayCandidates = @($gatewayCandidates | Select-Object -Unique)

                        [pscustomobject]@{
                            AdapterName = if ($null -ne $adapter -and $null -ne $adapter.Name) { $adapter.Name } else { '' }
                            InterfaceDescription = if ($null -ne $adapter -and $null -ne $adapter.InterfaceDescription) { $adapter.InterfaceDescription } else { '' }
                            IpAddresses = if ($ipAddresses.Count -gt 0) { $ipAddresses -join ', ' } else { '' }
                            Subnets = if ($prefixes.Count -gt 0) { $prefixes -join ', ' } else { '' }
                            Gateway = if ($gatewayCandidates.Count -gt 0) { $gatewayCandidates -join ', ' } else { '' }
                            DnsServers = if ($dnsServers.Count -gt 0) { $dnsServers -join ', ' } else { '' }
                            NetworkProfileCategory = Resolve-NetworkCategory -InterfaceAlias $adapter.Name -InterfaceIndex $adapter.ifIndex
                            IsDefaultSwitch = ($adapter.Name -like 'vEthernet (*Default Switch*)')
                        }
                    }
            )

            $defaultSwitchAdded = $false
            $defaultSwitchAdapters = @(Get-NetAdapter -IncludeHidden -Name 'vEthernet (Default Switch)' -ErrorAction SilentlyContinue)

            if ($defaultSwitchAdapters.Count -eq 0)
            {
                $defaultSwitchAdapters = @(Get-NetAdapter -IncludeHidden -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'vEthernet (*Default Switch*)' })
            }

            foreach ($adapter in $defaultSwitchAdapters)
            {
                if ($null -eq $adapter -or [string]::IsNullOrWhiteSpace($adapter.Name))
                {
                    continue
                }

                if ($items | Where-Object { $_.AdapterName -ceq $adapter.Name })
                {
                    $defaultSwitchAdded = $true
                    continue
                }

                $ipConfig = Get-NetIPConfiguration -InterfaceIndex $adapter.ifIndex -Detailed -ErrorAction SilentlyContinue
                $ipv4Addresses = @($ipConfig.IPv4Address | Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace($_.IPAddress) })
                $ipv6Addresses = @($ipConfig.IPv6Address | Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace($_.IPAddress) -and $_.IPAddress -notlike 'fe80:*' })
                $ipAddresses = @($ipv4Addresses | ForEach-Object { $_.IPAddress }) + @($ipv6Addresses | ForEach-Object { $_.IPAddress })
                $prefixes = @($ipv4Addresses | ForEach-Object { '/' + $_.PrefixLength }) + @($ipv6Addresses | ForEach-Object { '/' + $_.PrefixLength })
                $dnsServers = @($ipConfig.DnsServer.ServerAddresses | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

                $gatewayCandidates = @()
                if ($null -ne $ipConfig.IPv4DefaultGateway -and -not [string]::IsNullOrWhiteSpace($ipConfig.IPv4DefaultGateway.NextHop)) { $gatewayCandidates += $ipConfig.IPv4DefaultGateway.NextHop }
                if ($null -ne $ipConfig.IPv6DefaultGateway -and -not [string]::IsNullOrWhiteSpace($ipConfig.IPv6DefaultGateway.NextHop)) { $gatewayCandidates += $ipConfig.IPv6DefaultGateway.NextHop }
                $gatewayCandidates = @($gatewayCandidates | Select-Object -Unique)

                $items += [pscustomobject]@{
                    AdapterName = $adapter.Name
                    InterfaceDescription = if ($null -ne $adapter.InterfaceDescription) { $adapter.InterfaceDescription } else { 'Hyper-V Default Switch (ICS)' }
                    IpAddresses = if ($ipAddresses.Count -gt 0) { $ipAddresses -join ', ' } else { '' }
                    Subnets = if ($prefixes.Count -gt 0) { $prefixes -join ', ' } else { '' }
                    Gateway = if ($gatewayCandidates.Count -gt 0) { $gatewayCandidates -join ', ' } else { '' }
                    DnsServers = if ($dnsServers.Count -gt 0) { $dnsServers -join ', ' } else { '' }
                    NetworkProfileCategory = Resolve-NetworkCategory -InterfaceAlias $adapter.Name -InterfaceIndex $adapter.ifIndex
                    IsDefaultSwitch = $true
                }

                $defaultSwitchAdded = $true
            }

            if (-not $defaultSwitchAdded)
            {
                $defaultSwitch = Get-VMSwitch -SwitchType Internal -ErrorAction SilentlyContinue | Where-Object { $_.Name -ceq 'Default Switch' } | Select-Object -First 1
                if ($null -ne $defaultSwitch)
                {
                    $defaultAlias = "vEthernet ($($defaultSwitch.Name))"
                    $defaultAdapter = Get-NetAdapter -IncludeHidden -ErrorAction SilentlyContinue | Where-Object { $_.Name -ceq $defaultAlias -or $_.Name -like 'vEthernet (*Default Switch*)' } | Select-Object -First 1

                    $ipConfig = $null
                    if ($null -ne $defaultAdapter)
                    {
                        $ipConfig = Get-NetIPConfiguration -InterfaceIndex $defaultAdapter.ifIndex -Detailed -ErrorAction SilentlyContinue
                    }

                    if ($null -eq $ipConfig)
                    {
                        $ipConfig = Get-NetIPConfiguration -InterfaceAlias $defaultAlias -Detailed -ErrorAction SilentlyContinue
                    }

                    $ipv4Addresses = @($ipConfig.IPv4Address | Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace($_.IPAddress) })
                    $ipv6Addresses = @($ipConfig.IPv6Address | Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace($_.IPAddress) -and $_.IPAddress -notlike 'fe80:*' })

                    if ($ipv4Addresses.Count -eq 0 -and $ipv6Addresses.Count -eq 0)
                    {
                        $fallbackIpRows = @(
                            Get-NetIPAddress -AddressFamily IPv4,IPv6 -ErrorAction SilentlyContinue |
                                Where-Object {
                                    -not [string]::IsNullOrWhiteSpace($_.IPAddress) -and
                                    $_.IPAddress -notlike 'fe80:*' -and
                                    ($_.InterfaceAlias -ceq $defaultAlias -or $_.InterfaceAlias -like '*Default Switch*')
                                }
                        )

                        $ipv4Addresses = @($fallbackIpRows | Where-Object { $_.AddressFamily -eq 'IPv4' })
                        $ipv6Addresses = @($fallbackIpRows | Where-Object { $_.AddressFamily -eq 'IPv6' })
                    }

                    $ipAddresses = @($ipv4Addresses | ForEach-Object { $_.IPAddress }) + @($ipv6Addresses | ForEach-Object { $_.IPAddress })
                    $prefixes = @($ipv4Addresses | ForEach-Object { '/' + $_.PrefixLength }) + @($ipv6Addresses | ForEach-Object { '/' + $_.PrefixLength })

                    $dnsServers = @()
                    if ($null -ne $ipConfig -and $null -ne $ipConfig.DnsServer)
                    {
                        $dnsServers = @($ipConfig.DnsServer.ServerAddresses | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                    }

                    if ($dnsServers.Count -eq 0)
                    {
                        $dnsRows = @(
                            Get-DnsClientServerAddress -AddressFamily IPv4,IPv6 -ErrorAction SilentlyContinue |
                                Where-Object { $_.InterfaceAlias -ceq $defaultAlias -or $_.InterfaceAlias -like '*Default Switch*' }
                        )
                        $dnsServers = @($dnsRows | ForEach-Object { $_.ServerAddresses } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                    }

                    $gatewayCandidates = @()
                    if ($null -ne $ipConfig)
                    {
                        if ($null -ne $ipConfig.IPv4DefaultGateway -and -not [string]::IsNullOrWhiteSpace($ipConfig.IPv4DefaultGateway.NextHop)) { $gatewayCandidates += $ipConfig.IPv4DefaultGateway.NextHop }
                        if ($null -ne $ipConfig.IPv6DefaultGateway -and -not [string]::IsNullOrWhiteSpace($ipConfig.IPv6DefaultGateway.NextHop)) { $gatewayCandidates += $ipConfig.IPv6DefaultGateway.NextHop }
                    }

                    if ($gatewayCandidates.Count -eq 0)
                    {
                        $routeRows = @(
                            Get-NetRoute -AddressFamily IPv4,IPv6 -ErrorAction SilentlyContinue |
                                Where-Object {
                                    ($_.InterfaceAlias -ceq $defaultAlias -or $_.InterfaceAlias -like '*Default Switch*') -and
                                    ($_.DestinationPrefix -eq '0.0.0.0/0' -or $_.DestinationPrefix -eq '::/0') -and
                                    -not [string]::IsNullOrWhiteSpace($_.NextHop) -and
                                    $_.NextHop -ne '0.0.0.0' -and
                                    $_.NextHop -ne '::'
                                }
                        )
                        $gatewayCandidates = @($routeRows | ForEach-Object { $_.NextHop })
                    }

                    $gatewayCandidates = @($gatewayCandidates | Select-Object -Unique)
                    $dnsServers = @($dnsServers | Select-Object -Unique)

                    $items += [pscustomobject]@{
                        AdapterName = if (-not [string]::IsNullOrWhiteSpace($defaultAlias)) { $defaultAlias } else { 'Default Switch (ICS)' }
                        InterfaceDescription = if ($null -ne $defaultAdapter -and -not [string]::IsNullOrWhiteSpace($defaultAdapter.InterfaceDescription)) { $defaultAdapter.InterfaceDescription } else { 'Hyper-V Default Switch (ICS)' }
                        IpAddresses = if ($ipAddresses.Count -gt 0) { $ipAddresses -join ', ' } else { '' }
                        Subnets = if ($prefixes.Count -gt 0) { $prefixes -join ', ' } else { '' }
                        Gateway = if ($gatewayCandidates.Count -gt 0) { $gatewayCandidates -join ', ' } else { '' }
                        DnsServers = if ($dnsServers.Count -gt 0) { $dnsServers -join ', ' } else { '' }
                        NetworkProfileCategory = Resolve-NetworkCategory -InterfaceAlias $defaultAlias -InterfaceIndex $(if ($null -ne $defaultAdapter) { $defaultAdapter.ifIndex } else { 0 })
                        IsDefaultSwitch = $true
                    }
                }
            }

            $items | Sort-Object AdapterName -Unique | ConvertTo-Json -Depth 4 -Compress
            """;

        var rows = await InvokeJsonArrayAsync(script, cancellationToken);
        return rows.Select(row => new HostNetworkAdapterInfo
        {
            AdapterName = GetString(row, "AdapterName"),
            InterfaceDescription = GetString(row, "InterfaceDescription"),
            IpAddresses = GetString(row, "IpAddresses"),
            Subnets = GetString(row, "Subnets"),
            Gateway = GetString(row, "Gateway"),
            DnsServers = GetString(row, "DnsServers"),
            NetworkProfileCategory = GetString(row, "NetworkProfileCategory"),
            IsDefaultSwitch = GetBoolean(row, "IsDefaultSwitch")
        }).ToList();
    }

    public async Task<IReadOnlyList<HyperVCheckpointInfo>> GetCheckpointsAsync(string vmName, CancellationToken cancellationToken)
    {
        var script = $"$vm = Get-VM -Name {ToPsSingleQuoted(vmName)} -ErrorAction SilentlyContinue; $currentSnapshotId = ''; if ($null -ne $vm -and $null -ne $vm.ParentSnapshotId) {{ $currentSnapshotId = $vm.ParentSnapshotId.ToString() }}; @(Get-VMCheckpoint -VMName {ToPsSingleQuoted(vmName)} | ForEach-Object {{ $id = if ($null -ne $_.VMCheckpointId) {{ $_.VMCheckpointId.ToString() }} elseif ($null -ne $_.Id) {{ $_.Id.ToString() }} else {{ '' }}; $parentId = ''; if ($null -ne $_.ParentCheckpointId) {{ $parentId = $_.ParentCheckpointId.ToString() }} elseif ($null -ne $_.Parent -and $null -ne $_.Parent.VMCheckpointId) {{ $parentId = $_.Parent.VMCheckpointId.ToString() }}; $isCurrent = $false; if ($null -ne $_.IsCurrentSnapshot) {{ $isCurrent = [bool]$_.IsCurrentSnapshot }}; if (-not $isCurrent -and -not [string]::IsNullOrWhiteSpace($currentSnapshotId) -and -not [string]::IsNullOrWhiteSpace($id) -and $id -ceq $currentSnapshotId) {{ $isCurrent = $true }}; [pscustomobject]@{{ Id = $id; ParentId = $parentId; IsCurrent = $isCurrent; Name = if ($null -ne $_.Name) {{ $_.Name }} else {{ '' }}; CreationTime = if ($null -ne $_.CreationTime) {{ $_.CreationTime.ToString('o') }} else {{ '' }}; CheckpointType = if ($null -ne $_.CheckpointType) {{ $_.CheckpointType.ToString() }} else {{ '' }} }} }}) | ConvertTo-Json -Depth 4 -Compress";

        var rows = await InvokeJsonArrayAsync(script, cancellationToken);
        return rows.Select(row => new HyperVCheckpointInfo
        {
            Id = GetString(row, "Id"),
            ParentId = GetString(row, "ParentId"),
            IsCurrent = GetBoolean(row, "IsCurrent"),
            Name = GetString(row, "Name"),
            Created = GetDateTime(row, "CreationTime"),
            Type = GetString(row, "CheckpointType")
        }).ToList();
    }

    public async Task CreateCheckpointAsync(string vmName, string checkpointName, string? description, CancellationToken cancellationToken)
    {
        if (!string.IsNullOrWhiteSpace(description))
        {
            Log.Information("Checkpoint description currently informational only for VM {VmName}: {Description}", vmName, description);
        }

        try
        {
            await InvokeNonQueryAsync(
                $"Checkpoint-VM -VMName {ToPsSingleQuoted(vmName)} -SnapshotName {ToPsSingleQuoted(checkpointName)} -Confirm:$false",
                cancellationToken);
        }
        catch (InvalidOperationException ex) when (IsProductionCheckpointError(ex.Message))
        {
            Log.Warning(ex,
                "Production checkpoint creation failed for VM {VmName}. Retrying once with temporary Standard checkpoint type.",
                vmName);

            var vmNameQuoted = ToPsSingleQuoted(vmName);
            var checkpointNameQuoted = ToPsSingleQuoted(checkpointName);
            var fallbackScript = $"$vmName = {vmNameQuoted}; " +
                                 "$vm = Get-VM -Name $vmName; " +
                                 "$originalType = $vm.CheckpointType; " +
                                 "try { " +
                                 "Set-VM -Name $vmName -CheckpointType Standard; " +
                                 $"Checkpoint-VM -VMName $vmName -SnapshotName {checkpointNameQuoted} -Confirm:$false; " +
                                 "} finally { " +
                                 "if ($null -ne $originalType) { Set-VM -Name $vmName -CheckpointType $originalType } " +
                                 "}";

            await InvokeNonQueryAsync(fallbackScript, cancellationToken);
        }
    }

    public Task ApplyCheckpointAsync(string vmName, string checkpointName, string? checkpointId, CancellationToken cancellationToken)
    {
        var script = $"$vmName = {ToPsSingleQuoted(vmName)}; " +
                     $"$checkpointName = {ToPsSingleQuoted(checkpointName)}; " +
                     $"$checkpointId = {ToPsSingleQuoted(checkpointId ?? string.Empty)}; " +
                     "$checkpoint = $null; " +
                     "if (-not [string]::IsNullOrWhiteSpace($checkpointId)) { $checkpoint = Get-VMCheckpoint -VMName $vmName | ForEach-Object { $currentId = ''; if ($null -ne $_.VMCheckpointId) { $currentId = $_.VMCheckpointId.ToString() } elseif ($null -ne $_.Id) { $currentId = $_.Id.ToString() }; if ($currentId -ceq $checkpointId) { $_ } } | Select-Object -First 1 }; " +
                     "if ($null -eq $checkpoint) { $checkpoint = Get-VMCheckpoint -VMName $vmName | Where-Object { $_.Name -ceq $checkpointName } | Select-Object -First 1 }; " +
                     "if ($null -eq $checkpoint) { throw \"Checkpoint '$checkpointName' wurde auf VM '$vmName' nicht gefunden.\" }; " +
                     "Restore-VMCheckpoint -VMCheckpoint $checkpoint -Confirm:$false";

        return InvokeNonQueryAsync(script, cancellationToken);
    }

    public Task RemoveCheckpointAsync(string vmName, string checkpointName, string? checkpointId, CancellationToken cancellationToken)
    {
        var script = $"$vmName = {ToPsSingleQuoted(vmName)}; " +
                     $"$checkpointName = {ToPsSingleQuoted(checkpointName)}; " +
                     $"$checkpointId = {ToPsSingleQuoted(checkpointId ?? string.Empty)}; " +
                     "$checkpoint = $null; " +
                     "if (-not [string]::IsNullOrWhiteSpace($checkpointId)) { $checkpoint = Get-VMCheckpoint -VMName $vmName | ForEach-Object { $currentId = ''; if ($null -ne $_.VMCheckpointId) { $currentId = $_.VMCheckpointId.ToString() } elseif ($null -ne $_.Id) { $currentId = $_.Id.ToString() }; if ($currentId -ceq $checkpointId) { $_ } } | Select-Object -First 1 }; " +
                     "if ($null -eq $checkpoint) { $checkpoint = Get-VMCheckpoint -VMName $vmName | Where-Object { $_.Name -ceq $checkpointName } | Select-Object -First 1 }; " +
                     "if ($null -eq $checkpoint) { throw \"Checkpoint '$checkpointName' wurde auf VM '$vmName' nicht gefunden.\" }; " +
                     "Remove-VMCheckpoint -VMCheckpoint $checkpoint -Confirm:$false";

        return InvokeNonQueryAsync(script, cancellationToken);
    }

    public Task OpenVmConnectAsync(string vmName, string computerName, bool openWithSessionEdit, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var processStartInfo = new ProcessStartInfo("vmconnect.exe")
        {
            UseShellExecute = true
        };

        processStartInfo.ArgumentList.Add(computerName);
        processStartInfo.ArgumentList.Add(vmName);

        if (openWithSessionEdit)
        {
            processStartInfo.ArgumentList.Add("/edit");
        }

        try
        {
            Process.Start(processStartInfo);
            Log.Information("vmconnect started for VM {VmName} on host {ComputerName} (SessionEdit: {SessionEdit})", vmName, computerName, openWithSessionEdit);
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to start vmconnect for VM {VmName}", vmName);
            throw;
        }

        return Task.CompletedTask;
    }

    public Task ReopenVmConnectWithSessionEditAsync(string vmName, string computerName, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        try
        {
            CloseExistingVmConnectWindows(vmName, computerName);
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Existing vmconnect windows could not be fully closed for VM {VmName}", vmName);
        }

        return OpenVmConnectAsync(vmName, computerName, openWithSessionEdit: true, cancellationToken);
    }

    private static void CloseExistingVmConnectWindows(string vmName, string computerName)
    {
        var titleNeedles = new[]
        {
            vmName?.Trim() ?? string.Empty,
            computerName?.Trim() ?? string.Empty
        }
        .Where(item => !string.IsNullOrWhiteSpace(item))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToArray();

        if (titleNeedles.Length == 0)
        {
            return;
        }

        var candidates = Process.GetProcessesByName("vmconnect");
        foreach (var process in candidates)
        {
            try
            {
                var windowTitle = process.MainWindowTitle ?? string.Empty;
                var isMatching = titleNeedles.Any(needle => windowTitle.Contains(needle, StringComparison.OrdinalIgnoreCase));
                if (!isMatching)
                {
                    continue;
                }

                if (!process.CloseMainWindow())
                {
                    process.Kill(entireProcessTree: false);
                    continue;
                }

                if (!process.WaitForExit(1200))
                {
                    process.Kill(entireProcessTree: false);
                }
            }
            catch
            {
            }
        }
    }

    public async Task<(bool HasEnoughSpace, long RequiredBytes, long AvailableBytes, string TargetDrive)> CheckExportDiskSpaceAsync(
        string vmName,
        string destinationPath,
        CancellationToken cancellationToken)
    {
        var script =
            $"$vmName = {ToPsSingleQuoted(vmName)}; " +
            $"$destinationPath = {ToPsSingleQuoted(destinationPath)}; " +
            "$required = 0; " +
            "$measure = Measure-VM -VMName $vmName -ErrorAction SilentlyContinue; " +
            "if ($null -ne $measure -and $null -ne $measure.TotalDiskAllocation) { $required = [int64]$measure.TotalDiskAllocation }; " +
            "if ($required -le 0) { " +
            "$diskBytes = @(Get-VMHardDiskDrive -VMName $vmName -ErrorAction SilentlyContinue | ForEach-Object { if ($null -ne $_.Path -and (Test-Path -LiteralPath $_.Path -PathType Leaf)) { (Get-Item -LiteralPath $_.Path).Length } else { 0 } }); " +
            "$required = [int64](($diskBytes | Measure-Object -Sum).Sum) " +
            "}; " +
            "if ($required -le 0) { $required = 1 }; " +
            "$available = -1; " +
            "$targetDrive = ''; " +
            "$uncRoot = $null; " +
            "if ($destinationPath -match '^(\\\\[^\\\\]+\\\\[^\\\\]+)') { $uncRoot = $Matches[1]; $targetDrive = $uncRoot }; " +
            "if ([string]::IsNullOrWhiteSpace($targetDrive)) { $targetDrive = [System.IO.Path]::GetPathRoot($destinationPath) }; " +
            "if ([string]::IsNullOrWhiteSpace($targetDrive)) { $targetDrive = [System.IO.Path]::GetPathRoot((Get-Location).Path) }; " +
            "if (-not [string]::IsNullOrWhiteSpace($uncRoot)) { " +
            "$tempDriveName = 'HT' + [System.Guid]::NewGuid().ToString('N').Substring(0, 8); " +
            "try { " +
            "$uncDrive = New-PSDrive -Name $tempDriveName -PSProvider FileSystem -Root $uncRoot -ErrorAction Stop; " +
            "if ($null -ne $uncDrive -and $null -ne $uncDrive.Free) { $available = [int64]$uncDrive.Free } " +
            "} catch { $available = -1 } " +
            "finally { Remove-PSDrive -Name $tempDriveName -Force -ErrorAction SilentlyContinue } " +
            "} else { " +
            "if (-not [string]::IsNullOrWhiteSpace($targetDrive)) { " +
            "try { $available = [int64](New-Object System.IO.DriveInfo($targetDrive)).AvailableFreeSpace } catch { $available = -1 } " +
            "} " +
            "}; " +
            "$hasEnoughSpace = if ($available -ge 0) { $available -ge $required } else { $true }; " +
            "[pscustomobject]@{ HasEnoughSpace = $hasEnoughSpace; RequiredBytes = $required; AvailableBytes = $available; TargetDrive = $targetDrive } | ConvertTo-Json -Depth 3 -Compress";

        var rows = await InvokeJsonArrayAsync(script, cancellationToken);
        var row = rows.FirstOrDefault();
        if (row.ValueKind == JsonValueKind.Undefined)
        {
            return (false, 0, 0, string.Empty);
        }

        return (
            HasEnoughSpace: GetBoolean(row, "HasEnoughSpace"),
            RequiredBytes: GetInt64(row, "RequiredBytes"),
            AvailableBytes: GetInt64(row, "AvailableBytes"),
            TargetDrive: GetString(row, "TargetDrive"));
    }

    public async Task ExportVmAsync(string vmName, string destinationPath, IProgress<int>? progress, CancellationToken cancellationToken)
    {
        var script =
            $"$vmName = {ToPsSingleQuoted(vmName)}; " +
            $"$destinationPath = {ToPsSingleQuoted(destinationPath)}; " +
            "$job = Export-VM -Name $vmName -Path $destinationPath -Confirm:$false -AsJob; " +
            "if ($null -eq $job) { throw 'Export-Job konnte nicht gestartet werden.' }; " +
            "$lastProgress = -1; " +
            "while ($true) { " +
            "$job = Get-Job -Id $job.Id -ErrorAction SilentlyContinue; " +
            "if ($null -eq $job) { break }; " +
            "if ($job.State -ne 'Running' -and $job.State -ne 'NotStarted') { break }; " +
            "$percent = $null; " +
            "if ($job.ChildJobs.Count -gt 0) { " +
            "$childJob = $job.ChildJobs[0]; " +
            "$progressRecord = $childJob.Progress | Select-Object -Last 1; " +
            "if ($null -ne $progressRecord -and $null -ne $progressRecord.PercentComplete) { $percent = [int]$progressRecord.PercentComplete }; " +
            "if ($null -eq $percent -and $null -ne $childJob.PercentComplete) { $percent = [int]$childJob.PercentComplete }; " +
            "if ($null -eq $percent -and $null -ne $childJob.StatusMessage -and $childJob.StatusMessage -match '(?<p>\\d{1,3})\\s*%') { $percent = [int]$Matches['p'] } " +
            "}; " +
            "if ($null -eq $percent) { Start-Sleep -Milliseconds 500; continue }; " +
            "if ($lastProgress -ge 0 -and $percent -lt $lastProgress) { $percent = $lastProgress }; " +
            "if ($percent -gt 99) { $percent = 99 }; " +
            "if ($percent -ne $lastProgress) { Write-Output ('HT_PROGRESS:' + $percent); $lastProgress = $percent }; " +
            "Start-Sleep -Milliseconds 500 " +
            "}; " +
            "Wait-Job -Id $job.Id | Out-Null; " +
            "$job = Get-Job -Id $job.Id; " +
            "if ($job.State -ne 'Completed') { " +
            "$reason = ($job.ChildJobs | ForEach-Object { if ($null -ne $_.JobStateInfo -and $null -ne $_.JobStateInfo.Reason) { $_.JobStateInfo.Reason.Message } }) -join '; '; " +
            "if ([string]::IsNullOrWhiteSpace($reason)) { $reason = (Receive-Job -Id $job.Id -Keep 2>&1 | Out-String) }; " +
            "Remove-Job -Id $job.Id -Force -ErrorAction SilentlyContinue; " +
            "throw $reason " +
            "}; " +
            "Receive-Job -Id $job.Id -Keep | Out-Null; " +
            "Write-Output 'HT_PROGRESS:100'; " +
            "Remove-Job -Id $job.Id -Force -ErrorAction SilentlyContinue";

        progress?.Report(0);
        _ = await InvokePowerShellWithProgressAsync(script, progress, cancellationToken);
        progress?.Report(100);
    }

    public async Task<ImportVmResult> ImportVmAsync(
        string importPath,
        string destinationPath,
        string? requestedVmName,
        string? requestedFolderName,
        string importMode,
        IProgress<int>? progress,
        CancellationToken cancellationToken)
    {
        var script = $"$importPath = {ToPsSingleQuoted(importPath)}; " +
                     $"$destinationPath = {ToPsSingleQuoted(destinationPath)}; " +
                     $"$requestedVmName = {ToPsSingleQuoted(requestedVmName ?? string.Empty)}; " +
                     $"$requestedFolderName = {ToPsSingleQuoted(requestedFolderName ?? string.Empty)}; " +
                     $"$importMode = {ToPsSingleQuoted(importMode ?? string.Empty)}.ToLowerInvariant(); " +
                     "if (-not (Test-Path -LiteralPath $importPath)) { throw \"Import-Pfad nicht gefunden: $importPath\" }; " +
                     "if (@('copy','register','restore') -notcontains $importMode) { $importMode = 'copy' }; " +
                     "if ($importMode -eq 'copy' -and [string]::IsNullOrWhiteSpace($destinationPath)) { throw \"Zielpfad für den Import fehlt.\" }; " +
                     "if ($importMode -eq 'copy' -and -not (Test-Path -LiteralPath $destinationPath -PathType Container)) { New-Item -Path $destinationPath -ItemType Directory -Force | Out-Null }; " +
                     "$existingVms = @(Get-VM -ErrorAction SilentlyContinue); " +
                     "$existingVmNames = @($existingVms | Select-Object -ExpandProperty Name); " +
                     "$existingVmIds = @($existingVms | Select-Object -ExpandProperty Id); " +
                     "function Get-UniqueVmName([string]$baseName, [string[]]$reservedNames) { " +
                     "if ([string]::IsNullOrWhiteSpace($baseName)) { $baseName = 'Imported-VM' }; " +
                     "$candidate = $baseName; $index = 1; " +
                     "while ($reservedNames | Where-Object { $_ -ceq $candidate }) { $candidate = $baseName + '-' + $index; $index++ }; " +
                     "return $candidate }; " +
                     "function Normalize-FolderName([string]$value) { " +
                     "$name = if ([string]::IsNullOrWhiteSpace($value)) { '' } else { $value.Trim() }; " +
                     "if ([string]::IsNullOrWhiteSpace($name)) { return '' }; " +
                     "$name = [regex]::Replace($name, '[<>:/\\|?*]', '_'); " +
                     "$name = $name.Trim('. ').Trim(); " +
                     "if ([string]::IsNullOrWhiteSpace($name)) { return '' }; return $name }; " +
                     "function Get-UniqueFolderPath([string]$rootPath, [string]$baseName) { " +
                     "$cleanBase = Normalize-FolderName $baseName; if ([string]::IsNullOrWhiteSpace($cleanBase)) { $cleanBase = 'Imported-VM' }; " +
                     "$candidate = Join-Path $rootPath $cleanBase; $index = 1; " +
                     "while (Test-Path -LiteralPath $candidate -PathType Container) { $candidate = Join-Path $rootPath ($cleanBase + '-' + $index); $index++ }; " +
                     "return $candidate }; " +
                     "$destinationRoot = $destinationPath; " +
                     "if ([System.IO.Path]::GetFullPath($importPath).TrimEnd('\\') -eq [System.IO.Path]::GetFullPath($destinationRoot).TrimEnd('\\')) { $destinationRoot = Join-Path $destinationRoot ('Imported-' + (Get-Date -Format 'yyyyMMdd-HHmmss')) }; " +
                     "$vmStoragePath = ''; " +
                     "$configPath = $importPath; " +
                     "if (Test-Path -LiteralPath $importPath -PathType Container) { " +
                     "$configFile = Get-ChildItem -LiteralPath $importPath -Recurse -File | Where-Object { $_.Extension -in '.vmcx', '.xml' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1; " +
                     "if ($null -eq $configFile) { throw \"Keine VM-Konfigurationsdatei (.vmcx/.xml) im Ordner gefunden.\" }; " +
                     "$configPath = $configFile.FullName; }; " +
                     "if ($importMode -eq 'copy') { " +
                     "$baseFolderName = if (-not [string]::IsNullOrWhiteSpace($requestedFolderName)) { $requestedFolderName } elseif (-not [string]::IsNullOrWhiteSpace($requestedVmName)) { $requestedVmName } else { [System.IO.Path]::GetFileNameWithoutExtension($configPath) }; " +
                     "$vmStoragePath = Get-UniqueFolderPath $destinationRoot $baseFolderName; " +
                     "$virtualMachinePath = Join-Path $vmStoragePath 'Virtual Machines'; " +
                     "$snapshotPath = Join-Path $vmStoragePath 'Snapshots'; " +
                     "$vhdPath = Join-Path $vmStoragePath 'Virtual Disks'; " +
                     "$smartPagingPath = Join-Path $vmStoragePath 'Smart Paging'; " +
                     "foreach ($path in @($vmStoragePath, $virtualMachinePath, $snapshotPath, $vhdPath, $smartPagingPath)) { if (-not (Test-Path -LiteralPath $path -PathType Container)) { New-Item -Path $path -ItemType Directory -Force | Out-Null } }; " +
                     "$job = Import-VM -Path $configPath -Copy -GenerateNewId -VirtualMachinePath $virtualMachinePath -VhdDestinationPath $vhdPath -SnapshotFilePath $snapshotPath -SmartPagingFilePath $smartPagingPath -Confirm:$false -AsJob; " +
                     "} elseif ($importMode -eq 'register') { $job = Import-VM -Path $configPath -Register -Confirm:$false -AsJob; } " +
                     "else { $job = Import-VM -Path $configPath -Confirm:$false -AsJob; }; " +
                     "if ($null -eq $job) { throw 'Import-Job konnte nicht gestartet werden.' }; " +
                     "$lastProgress = -1; " +
                     "while ($true) { " +
                     "$job = Get-Job -Id $job.Id -ErrorAction SilentlyContinue; if ($null -eq $job) { break }; " +
                     "if ($job.State -ne 'Running' -and $job.State -ne 'NotStarted') { break }; " +
                     "$percent = $null; " +
                     "if ($job.ChildJobs.Count -gt 0) { $progressRecord = $job.ChildJobs[0].Progress | Select-Object -Last 1; if ($null -ne $progressRecord -and $null -ne $progressRecord.PercentComplete) { $percent = [int]$progressRecord.PercentComplete } }; " +
                     "if ($null -eq $percent) { Start-Sleep -Milliseconds 500; continue }; " +
                     "if ($percent -gt 99) { $percent = 99 }; " +
                     "if ($percent -ne $lastProgress) { Write-Output ('HT_PROGRESS:' + $percent); $lastProgress = $percent }; " +
                     "Start-Sleep -Milliseconds 500 }; " +
                     "Wait-Job -Id $job.Id | Out-Null; $job = Get-Job -Id $job.Id; " +
                     "if ($job.State -ne 'Completed') { " +
                     "$reason = ($job.ChildJobs | ForEach-Object { if ($null -ne $_.JobStateInfo -and $null -ne $_.JobStateInfo.Reason) { $_.JobStateInfo.Reason.Message } }) -join '; '; " +
                     "if ([string]::IsNullOrWhiteSpace($reason)) { $reason = (Receive-Job -Id $job.Id -Keep 2>&1 | Out-String) }; " +
                     "Remove-Job -Id $job.Id -Force -ErrorAction SilentlyContinue; throw $reason }; " +
                     "$jobOutput = @(Receive-Job -Id $job.Id -Keep 2>&1); " +
                     "$importedVm = $jobOutput | Where-Object { $_ -is [Microsoft.HyperV.PowerShell.VirtualMachine] } | Select-Object -First 1; " +
                     "if ($null -eq $importedVm) { $newVmCandidates = @(Get-VM -ErrorAction SilentlyContinue | Where-Object { $existingVmIds -notcontains $_.Id }); if ($newVmCandidates.Count -gt 0) { $importedVm = $newVmCandidates | Select-Object -First 1 } }; " +
                     "Remove-Job -Id $job.Id -Force -ErrorAction SilentlyContinue; " +
                     "if ($null -eq $importedVm) { throw \"Import-VM hat keine VM zurückgegeben.\" }; " +
                     "$importedName = if ($null -ne $importedVm.Name) { $importedVm.Name } else { '' }; " +
                     "$renamed = $false; $originalName = $importedName; " +
                     "if (-not [string]::IsNullOrWhiteSpace($requestedVmName)) { " +
                     "$targetName = Get-UniqueVmName $requestedVmName @($existingVmNames + $importedName); " +
                     "if (-not [string]::IsNullOrWhiteSpace($targetName) -and $targetName -cne $importedName) { Rename-VM -VM $importedVm -NewName $targetName -Confirm:$false -ErrorAction Stop; $importedName = $targetName; $renamed = $true } " +
                     "} elseif (-not [string]::IsNullOrWhiteSpace($importedName) -and ($existingVmNames | Where-Object { $_ -ceq $importedName })) { " +
                     "$targetName = Get-UniqueVmName ($importedName + '-import') @($existingVmNames + $importedName); " +
                     "Rename-VM -VM $importedVm -NewName $targetName -Confirm:$false -ErrorAction Stop; $importedName = $targetName; $renamed = $true }; " +
                     "Write-Output 'HT_PROGRESS:100'; " +
                     "Write-Output ('HT_RESULT:' + $importedName); " +
                     "Write-Output ('HT_ORIGINAL:' + $originalName); " +
                     "Write-Output ('HT_RENAMED:' + $renamed); " +
                     "Write-Output ('HT_DESTINATION:' + $vmStoragePath); " +
                     "Write-Output ('HT_MODE:' + $importMode)";

        progress?.Report(0);
        var output = await InvokePowerShellWithProgressAsync(script, progress, cancellationToken);
        progress?.Report(100);

        static string? GetResultToken(string source, string prefix)
        {
            return source
                .Split(["\r\n", "\n"], StringSplitOptions.RemoveEmptyEntries)
                .Select(line => line.Trim())
                .LastOrDefault(line => line.StartsWith(prefix, StringComparison.Ordinal));
        }

        var importedName = GetResultToken(output, "HT_RESULT:");
        var originalName = GetResultToken(output, "HT_ORIGINAL:");
        var renamedFlag = GetResultToken(output, "HT_RENAMED:");
        var destinationFolder = GetResultToken(output, "HT_DESTINATION:");
        var importModeResult = GetResultToken(output, "HT_MODE:");

        if (string.IsNullOrWhiteSpace(importedName))
        {
            throw new InvalidOperationException("Import-VM hat keinen VM-Namen zurückgegeben.");
        }

        var resolvedImportedName = importedName["HT_RESULT:".Length..].Trim();
        var resolvedOriginalName = string.IsNullOrWhiteSpace(originalName)
            ? resolvedImportedName
            : originalName["HT_ORIGINAL:".Length..].Trim();
        var renamedDueToConflict = !string.IsNullOrWhiteSpace(renamedFlag)
            && bool.TryParse(renamedFlag["HT_RENAMED:".Length..].Trim(), out var parsedRenamed)
            && parsedRenamed;

        return new ImportVmResult
        {
            VmName = resolvedImportedName,
            OriginalName = resolvedOriginalName,
            RenamedDueToConflict = renamedDueToConflict,
            DestinationFolderPath = string.IsNullOrWhiteSpace(destinationFolder) ? string.Empty : destinationFolder["HT_DESTINATION:".Length..].Trim(),
            ImportMode = string.IsNullOrWhiteSpace(importModeResult) ? string.Empty : importModeResult["HT_MODE:".Length..].Trim()
        };
    }

    private async Task InvokeNonQueryAsync(string script, CancellationToken cancellationToken)
    {
        _ = await InvokePowerShellAsync(script, cancellationToken);
    }

    private static async Task<IReadOnlyList<JsonElement>> InvokeJsonArrayAsync(string script, CancellationToken cancellationToken)
    {
        var json = await InvokePowerShellAsync(script, cancellationToken);
        if (string.IsNullOrWhiteSpace(json))
        {
            return [];
        }

        using var document = JsonDocument.Parse(json);
        if (document.RootElement.ValueKind == JsonValueKind.Array)
        {
            return document.RootElement.EnumerateArray().Select(element => element.Clone()).ToList();
        }

        if (document.RootElement.ValueKind == JsonValueKind.Object)
        {
            return [document.RootElement.Clone()];
        }

        return [];
    }

    private static async Task<string> InvokePowerShellAsync(string script, CancellationToken cancellationToken)
    {
        var wrappedScript = "$ErrorActionPreference = 'Stop'; " +
                            "[Console]::InputEncoding = [System.Text.Encoding]::UTF8; " +
                            "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; " +
                            "$OutputEncoding = [System.Text.Encoding]::UTF8; " +
                            $"Import-Module Hyper-V -ErrorAction Stop; {script}";
        var processStartInfo = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        processStartInfo.ArgumentList.Add("-NoProfile");
        processStartInfo.ArgumentList.Add("-NonInteractive");
        processStartInfo.ArgumentList.Add("-ExecutionPolicy");
        processStartInfo.ArgumentList.Add("Bypass");
        processStartInfo.ArgumentList.Add("-Command");
        processStartInfo.ArgumentList.Add(wrappedScript);
        processStartInfo.StandardOutputEncoding = Encoding.UTF8;
        processStartInfo.StandardErrorEncoding = Encoding.UTF8;

        using var process = Process.Start(processStartInfo);
        if (process is null)
        {
            throw new InvalidOperationException("PowerShell konnte nicht gestartet werden.");
        }

        var standardOutputTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var standardErrorTask = process.StandardError.ReadToEndAsync(cancellationToken);

        try
        {
            await process.WaitForExitAsync(cancellationToken);
        }
        catch (OperationCanceledException)
        {
            try
            {
                if (!process.HasExited)
                {
                    process.Kill(entireProcessTree: true);
                }
            }
            catch
            {
            }

            throw;
        }

        var standardOutput = await standardOutputTask;
        var standardError = await standardErrorTask;

        if (process.ExitCode != 0)
        {
            var message = string.IsNullOrWhiteSpace(standardError) ? standardOutput : standardError;

            if (IsHyperVPermissionError(message))
            {
                throw new UnauthorizedAccessException(
                    "Keine Berechtigung für Hyper-V. Bitte HyperTool als Administrator starten oder den Benutzer zur Gruppe 'Hyper-V-Administratoren' hinzufügen.");
            }

            throw new InvalidOperationException($"Hyper-V PowerShell command failed:{Environment.NewLine}{message.Trim()}");
        }

        return standardOutput.Trim();
    }

    private static async Task InvokePowerShellElevatedNonQueryAsync(string script, CancellationToken cancellationToken)
    {
        var statusFilePath = Path.Combine(Path.GetTempPath(), $"hypertool-netprofile-{Guid.NewGuid():N}.txt");
        var statusFilePathPs = ToPsSingleQuoted(statusFilePath);

        var wrappedScript =
            "$ErrorActionPreference = 'Stop'; " +
            "[Console]::InputEncoding = [System.Text.Encoding]::UTF8; " +
            "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; " +
            "$OutputEncoding = [System.Text.Encoding]::UTF8; " +
            "$statusFilePath = " + statusFilePathPs + "; " +
            "$statusDir = Split-Path -Parent -Path $statusFilePath; " +
            "if (-not [string]::IsNullOrWhiteSpace($statusDir) -and -not (Test-Path -LiteralPath $statusDir)) { New-Item -Path $statusDir -ItemType Directory -Force | Out-Null }; " +
            "Import-Module Hyper-V -ErrorAction Stop; " +
            "try { " +
            script + "; " +
            "Set-Content -LiteralPath $statusFilePath -Value 'OK' -Encoding UTF8 -Force; " +
            "} catch { " +
            "$msg = if ($null -ne $_.Exception -and -not [string]::IsNullOrWhiteSpace($_.Exception.Message)) { $_.Exception.Message } else { ($_ | Out-String) }; " +
            "Set-Content -LiteralPath $statusFilePath -Value ('ERROR:' + $msg) -Encoding UTF8 -Force; " +
            "exit 1; " +
            "}";

        var encodedScript = Convert.ToBase64String(Encoding.Unicode.GetBytes(wrappedScript));
        var args = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand {encodedScript}";

        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = args,
                UseShellExecute = true,
                Verb = "runas",
                WorkingDirectory = Environment.CurrentDirectory
            };

            using var process = Process.Start(startInfo);
            if (process is null)
            {
                throw new InvalidOperationException("Erhöhter PowerShell-Prozess konnte nicht gestartet werden.");
            }

            await process.WaitForExitAsync(cancellationToken);

            string? statusText = null;
            if (File.Exists(statusFilePath))
            {
                statusText = await File.ReadAllTextAsync(statusFilePath, cancellationToken);
            }

            if (!string.IsNullOrWhiteSpace(statusText))
            {
                var trimmed = statusText.Trim();
                if (trimmed.StartsWith("ERROR:", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(trimmed["ERROR:".Length..].Trim());
                }

                if (trimmed.Equals("OK", StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }

            if (process.ExitCode != 0)
            {
                throw new InvalidOperationException("Erhöhte Ausführung fehlgeschlagen.");
            }
        }
        catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
        {
            throw new UnauthorizedAccessException("UAC-Abfrage wurde vom Benutzer abgebrochen.", ex);
        }
        finally
        {
            try
            {
                if (File.Exists(statusFilePath))
                {
                    File.Delete(statusFilePath);
                }
            }
            catch
            {
            }
        }
    }

    private static async Task<string> InvokePowerShellWithProgressAsync(string script, IProgress<int>? progress, CancellationToken cancellationToken)
    {
        var wrappedScript = "$ErrorActionPreference = 'Stop'; " +
                            "[Console]::InputEncoding = [System.Text.Encoding]::UTF8; " +
                            "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; " +
                            "$OutputEncoding = [System.Text.Encoding]::UTF8; " +
                            $"Import-Module Hyper-V -ErrorAction Stop; {script}";

        var processStartInfo = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        processStartInfo.ArgumentList.Add("-NoProfile");
        processStartInfo.ArgumentList.Add("-NonInteractive");
        processStartInfo.ArgumentList.Add("-ExecutionPolicy");
        processStartInfo.ArgumentList.Add("Bypass");
        processStartInfo.ArgumentList.Add("-Command");
        processStartInfo.ArgumentList.Add(wrappedScript);
        processStartInfo.StandardOutputEncoding = Encoding.UTF8;
        processStartInfo.StandardErrorEncoding = Encoding.UTF8;

        using var process = Process.Start(processStartInfo);
        if (process is null)
        {
            throw new InvalidOperationException("PowerShell konnte nicht gestartet werden.");
        }

        var outputBuilder = new StringBuilder();
        var errorBuilder = new StringBuilder();

        process.OutputDataReceived += (_, args) =>
        {
            if (string.IsNullOrWhiteSpace(args.Data))
            {
                return;
            }

            if (TryParseProgressLine(args.Data, out var percent))
            {
                progress?.Report(percent);
                return;
            }

            outputBuilder.AppendLine(args.Data);
        };

        process.ErrorDataReceived += (_, args) =>
        {
            if (!string.IsNullOrWhiteSpace(args.Data))
            {
                errorBuilder.AppendLine(args.Data);
            }
        };

        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        try
        {
            await process.WaitForExitAsync(cancellationToken);
        }
        catch (OperationCanceledException)
        {
            try
            {
                if (!process.HasExited)
                {
                    process.Kill(entireProcessTree: true);
                }
            }
            catch
            {
            }

            throw;
        }

        if (process.ExitCode != 0)
        {
            var message = errorBuilder.Length == 0 ? outputBuilder.ToString() : errorBuilder.ToString();

            if (IsHyperVPermissionError(message))
            {
                throw new UnauthorizedAccessException(
                    "Keine Berechtigung für Hyper-V. Bitte HyperTool als Administrator starten oder den Benutzer zur Gruppe 'Hyper-V-Administratoren' hinzufügen.");
            }

            throw new InvalidOperationException($"Hyper-V PowerShell command failed:{Environment.NewLine}{message.Trim()}");
        }

        return outputBuilder.ToString().Trim();
    }

    private static string ToPsSingleQuoted(string value)
    {
        var escaped = (value ?? string.Empty).Replace("'", "''", StringComparison.Ordinal);
        return $"'{escaped}'";
    }

    private static string GetString(JsonElement source, string propertyName)
    {
        if (!source.TryGetProperty(propertyName, out var value))
        {
            return string.Empty;
        }

        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString() ?? string.Empty,
            JsonValueKind.Null => string.Empty,
            _ => value.ToString()
        };
    }

    private static DateTime GetDateTime(JsonElement source, string propertyName)
    {
        if (!source.TryGetProperty(propertyName, out var value))
        {
            return DateTime.MinValue;
        }

        return value.ValueKind switch
        {
            JsonValueKind.String when DateTime.TryParse(value.GetString(), CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var parsed) => parsed,
            _ => DateTime.MinValue
        };
    }

    private static long GetInt64(JsonElement source, string propertyName)
    {
        if (!source.TryGetProperty(propertyName, out var value))
        {
            return 0;
        }

        return value.ValueKind switch
        {
            JsonValueKind.Number when value.TryGetInt64(out var parsed) => parsed,
            JsonValueKind.String when long.TryParse(value.GetString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) => parsed,
            _ => 0
        };
    }

    private static bool GetBoolean(JsonElement source, string propertyName)
    {
        if (!source.TryGetProperty(propertyName, out var value))
        {
            return false;
        }

        return value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.String when bool.TryParse(value.GetString(), out var parsed) => parsed,
            _ => false
        };
    }

    private static bool TryParseProgressLine(string line, out int percent)
    {
        percent = 0;
        const string prefix = "HT_PROGRESS:";
        if (!line.StartsWith(prefix, StringComparison.Ordinal))
        {
            return false;
        }

        var value = line[prefix.Length..].Trim();
        if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
        {
            return false;
        }

        percent = Math.Clamp(parsed, 0, 100);
        return true;
    }

    private static bool IsHyperVPermissionError(string? message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return false;
        }

        return message.Contains("erforderliche Berechtigung", StringComparison.OrdinalIgnoreCase)
               || message.Contains("authorization policy", StringComparison.OrdinalIgnoreCase)
               || message.Contains("required permission", StringComparison.OrdinalIgnoreCase)
               || message.Contains("access is denied", StringComparison.OrdinalIgnoreCase)
               || message.Contains("virtualizationexception", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsProductionCheckpointError(string? message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return false;
        }

        return message.Contains("production checkpoint", StringComparison.OrdinalIgnoreCase)
               || message.Contains("Produktionsprüfpunkt", StringComparison.OrdinalIgnoreCase)
               || message.Contains("VSS", StringComparison.OrdinalIgnoreCase)
               || message.Contains("Die Erstellung eines Prüfpunkts ist fehlgeschlagen", StringComparison.OrdinalIgnoreCase)
               || message.Contains("failed to create checkpoint", StringComparison.OrdinalIgnoreCase);
    }
}