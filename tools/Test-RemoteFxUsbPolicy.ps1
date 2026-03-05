param(
    [string]$Action,

    # Bei Get zuerst Computer-Richtlinien neu anwenden (gpupdate /target:computer /force)
    [switch]$ApplyGpUpdate,

    # Fuer Local Computer Policy: LocalMachine
    [ValidateSet("LocalMachine", "CurrentUser")]
    [string]$Scope = "LocalMachine"
)

function Resolve-Action([string]$rawAction) {
    if ([string]::IsNullOrWhiteSpace($rawAction)) {
        Write-Host "Waehle Aktion:" -ForegroundColor Cyan
        Write-Host "  1 = Status lesen (Get)"
        Write-Host "  2 = Aktivieren (Enable)"
        Write-Host "  3 = Deaktivieren (Disable)"
        $rawAction = Read-Host "Eingabe"
    }

    $normalizedAction = if ($null -eq $rawAction) { "" } else { $rawAction }
    $normalizedAction = $normalizedAction.Trim().ToLowerInvariant()

    switch ($normalizedAction) {
        "1" { return "Get" }
        "2" { return "Enable" }
        "3" { return "Disable" }
        "get" { return "Get" }
        "enable" { return "Enable" }
        "disable" { return "Disable" }
        default {
            throw "Ungueltige Action '$rawAction'. Erlaubt: Get/Enable/Disable oder 1/2/3."
        }
    }
}

$Action = Resolve-Action $Action

$baseHive = if ($Scope -eq "LocalMachine") { "HKLM:" } else { "HKCU:" }

# Beide Pfade, damit aeltere/abweichende Setups auch abgedeckt sind
$policyPaths = @(
    "$baseHive\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\Client",
    "$baseHive\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services"
)

# Relevante Werte fuer RemoteFX USB Redirection
$policyValues = @(
    "fEnableUsbRedirection",
    "fEnableUsbBlockDeviceBySetupClass",
    "fEnableUsbSelectDeviceByInterface",
    "fEnableUsbNoAckIsochWriteToDevice"
)

function Get-PolicyState {
    $found = $false
    $active = $false
    $rows = @()

    foreach ($path in $policyPaths) {
        foreach ($name in $policyValues) {
            $val = $null
            try {
                $val = (Get-ItemProperty -Path $path -Name $name -ErrorAction Stop).$name
                $found = $true
                if ([int]$val -gt 0) { $active = $true }
            }
            catch {
                # Wert/Pfad nicht vorhanden -> ignorieren
            }

            $rows += [pscustomobject]@{
                Scope = $Scope
                Path = $path
                Name = $name
                Value = if ($null -eq $val) { "<not set>" } else { [int]$val }
            }
        }
    }

    [pscustomobject]@{
        FoundAnyPolicyValue = $found
        IsActive = $active
        HasOnlyZeroValues = ($found -and -not $active)
        Details = $rows
    }
}

function Invoke-ComputerPolicyRefresh {
    Write-Host "Wende Computer-Richtlinien neu an (gpupdate /target:computer /force)..." -ForegroundColor Cyan
    $gpupdateOutput = & gpupdate /target:computer /force 2>&1
    $exitCode = $LASTEXITCODE

    if ($gpupdateOutput) {
        $gpupdateOutput | ForEach-Object { Write-Host $_ }
    }

    if ($exitCode -ne 0) {
        throw "gpupdate fehlgeschlagen (ExitCode $exitCode)."
    }
}

function Set-PolicyState([bool]$enabled) {
    $mainVal = if ($enabled) { 1 } else { 0 }
    $isoVal = if ($enabled) { 80 } else { 0 }

    foreach ($path in $policyPaths) {
        New-Item -Path $path -Force | Out-Null

        New-ItemProperty -Path $path -Name "fEnableUsbRedirection" -PropertyType DWord -Value $mainVal -Force | Out-Null
        New-ItemProperty -Path $path -Name "fEnableUsbBlockDeviceBySetupClass" -PropertyType DWord -Value $mainVal -Force | Out-Null
        New-ItemProperty -Path $path -Name "fEnableUsbSelectDeviceByInterface" -PropertyType DWord -Value $mainVal -Force | Out-Null
        New-ItemProperty -Path $path -Name "fEnableUsbNoAckIsochWriteToDevice" -PropertyType DWord -Value $isoVal -Force | Out-Null
    }
}

if ($Scope -eq "LocalMachine") {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).
        IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin -and $Action -ne "Get") {
        throw "Fuer Enable/Disable mit Scope=LocalMachine PowerShell als Administrator starten."
    }
}

switch ($Action) {
    "Get" {
        if ($ApplyGpUpdate) {
            if ($Scope -ne "LocalMachine") {
                Write-Host "Hinweis: -ApplyGpUpdate wirkt auf Computer-Richtlinien (LocalMachine)." -ForegroundColor Yellow
            }

            try {
                Invoke-ComputerPolicyRefresh
            }
            catch {
                Write-Host "Warnung: $_" -ForegroundColor Yellow
            }

            Write-Host ""
        }

        $state = Get-PolicyState
        Write-Host "FoundAnyPolicyValue: $($state.FoundAnyPolicyValue)"
        Write-Host "IsActive:            $($state.IsActive)"
        Write-Host "HasOnlyZeroValues:   $($state.HasOnlyZeroValues)"

        if ($state.IsActive) {
            Write-Host "Hinweis: Richtlinie ist effektiv noch AKTIV (mindestens ein Wert > 0)." -ForegroundColor Yellow
            Write-Host "Wenn du in gpedit auf 'Deaktiviert' gestellt hast, fuehre zuerst mit Adminrechten aus:" -ForegroundColor Yellow
            Write-Host "  powershell -ExecutionPolicy Bypass -File .\Test-RemoteFxUsbPolicy.ps1 -Action Get -Scope LocalMachine -ApplyGpUpdate" -ForegroundColor Yellow
            Write-Host "Bleibt es danach aktiv, wird die Einstellung sehr wahrscheinlich von einer anderen GPO ueberschrieben (z.B. Domain-GPO)." -ForegroundColor Yellow
        }

        Write-Host ""
        Write-Host "Tabellarische Ansicht (mit Umbruch):"
        ($state.Details | Format-Table Scope, Path, Name, Value -Wrap | Out-String -Width 4096).TrimEnd() | Write-Host

        Write-Host ""
        Write-Host "Volle Detailansicht (ungekuerzt):"
        foreach ($entry in $state.Details) {
            Write-Host "----------------------------------------"
            Write-Host "Scope : $($entry.Scope)"
            Write-Host "Path  : $($entry.Path)"
            Write-Host "Name  : $($entry.Name)"
            Write-Host "Value : $($entry.Value)"
        }
    }
    "Enable" {
        Set-PolicyState -enabled $true
        Write-Host "RemoteFX USB Device Redirection wurde aktiviert."
        Write-Host "Hinweis: gpupdate /force und ggf. Neustart erforderlich."
        Get-PolicyState | Select-Object FoundAnyPolicyValue, IsActive
    }
    "Disable" {
        Set-PolicyState -enabled $false
        Write-Host "RemoteFX USB Device Redirection wurde deaktiviert."
        Write-Host "Hinweis: gpupdate /force und ggf. Neustart erforderlich."
        Get-PolicyState | Select-Object FoundAnyPolicyValue, IsActive
    }
}
