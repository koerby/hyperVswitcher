param(
    [ValidateSet('remove', 'reset', 'install', 'check')]
    [string]$Mode = '',
    [switch]$PurgeOnly,
    [switch]$ReinstallOnly,
    [switch]$NoPause
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
    $OutputEncoding = [System.Text.UTF8Encoding]::new($false)
}
catch {
}

function Test-IsAdministrator {
    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Assert-Administrator {
    if (Test-IsAdministrator) {
        return
    }

    throw 'Bitte als Administrator starten.'
}

$rootPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices'
$services = @(
    @{ Id = '6c4eb1be-40e8-4c8b-a4d6-5b0f67d7e40f'; Name = 'HyperTool Hyper-V Socket USB Tunnel' },
    @{ Id = '67c53bca-3f3d-4628-98e4-e45be5d6d1ad'; Name = 'HyperTool Hyper-V Socket Diagnostics' },
    @{ Id = 'e7db04df-0e32-4f30-a4dc-c6cbc31a8792'; Name = 'HyperTool Hyper-V Socket Shared Folder Catalog' },
    @{ Id = '54b2c423-6f79-47d8-a77d-8cab14e3f041'; Name = 'HyperTool Hyper-V Socket Host Identity' },
    @{ Id = '91df7cec-c5ba-452a-b072-42e5f672d5f9'; Name = 'HyperTool Hyper-V Socket File Service' }
)

$legacyServiceIds = @(
    '0f9db05a-531f-4fd8-9b4d-675f5f06f0d8' # Deprecated: Shared Folder Credential Service
)

function Resolve-Mode {
    if ($PurgeOnly -and $ReinstallOnly) {
        throw 'PurgeOnly und ReinstallOnly koennen nicht gleichzeitig verwendet werden.'
    }

    if (-not [string]::IsNullOrWhiteSpace($Mode)) {
        return $Mode
    }

    if ($PurgeOnly) {
        return 'remove'
    }

    if ($ReinstallOnly) {
        return 'install'
    }

    Write-Host ''
    Write-Host 'Waehle Aktion fuer HyperTool Hyper-V Socket Registry-Services:'
    Write-Host '  [1] Nur entfernen'
    Write-Host '  [2] Entfernen und neu setzen (Reset)'
    Write-Host '  [3] Nur neu setzen'
    Write-Host '  [4] Nur pruefen (keine Aenderungen)'

    while ($true) {
        $choice = (Read-Host 'Eingabe (1/2/3/4)').Trim()
        switch ($choice) {
            '1' { return 'remove' }
            '2' { return 'reset' }
            '3' { return 'install' }
            '4' { return 'check' }
            default {
                Write-Host "[WARN] Ungueltige Eingabe '$choice'. Bitte 1, 2, 3 oder 4 eingeben."
            }
        }
    }
}

function Remove-LegacyServiceKeys {
    foreach ($legacyId in $legacyServiceIds) {
        $legacyPath = Join-Path $rootPath $legacyId
        if (Test-Path -LiteralPath $legacyPath) {
            Remove-Item -LiteralPath $legacyPath -Recurse -Force
            Write-Host "[OK] Legacy entfernt: $legacyId"
        }
    }
}

function Remove-UnknownHyperToolKeys {
    if (-not (Test-Path -LiteralPath $rootPath)) {
        return
    }

    $knownIds = @($services.Id + $legacyServiceIds)
    $knownMap = @{}
    foreach ($id in $knownIds) {
        $knownMap[$id.ToLowerInvariant()] = $true
    }

    foreach ($key in Get-ChildItem -Path $rootPath -ErrorAction SilentlyContinue) {
        $name = [string]$key.PSChildName
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        if ($knownMap.ContainsKey($name.ToLowerInvariant())) {
            continue
        }

        $elementName = ''
        try {
            $elementName = [string](Get-ItemPropertyValue -Path $key.PSPath -Name 'ElementName' -ErrorAction Stop)
        }
        catch {
            $elementName = ''
        }

        if ($elementName -like 'HyperTool*') {
            Remove-Item -LiteralPath $key.PSPath -Recurse -Force
            Write-Host "[OK] Unbekannter HyperTool-Eintrag entfernt: $name ($elementName)"
        }
    }
}

function Remove-ServiceKeys {
    foreach ($svc in $services) {
        $keyPath = Join-Path $rootPath $svc.Id
        if (Test-Path -LiteralPath $keyPath) {
            Remove-Item -LiteralPath $keyPath -Recurse -Force
            Write-Host "[OK] Entfernt: $($svc.Id)"
        }
        else {
            Write-Host "[INFO] Nicht vorhanden: $($svc.Id)"
        }
    }
}

function Install-ServiceKeys {
    if (-not (Test-Path -LiteralPath $rootPath)) {
        New-Item -Path $rootPath -Force | Out-Null
    }

    foreach ($svc in $services) {
        $keyPath = Join-Path $rootPath $svc.Id
        $key = New-Item -Path $keyPath -Force
        New-ItemProperty -Path $key.PSPath -Name 'ElementName' -Value $svc.Name -PropertyType String -Force | Out-Null
        Write-Host "[OK] Installiert: $($svc.Id) -> $($svc.Name)"
    }
}

function Show-ServiceState {
    Write-Host ''
    Write-Host '=== HyperTool Hyper-V Socket Service-Status ==='

    foreach ($svc in $services) {
        $keyPath = Join-Path $rootPath $svc.Id
        $exists = Test-Path -LiteralPath $keyPath
        $elementName = ''

        if ($exists) {
            try {
                $elementName = [string](Get-ItemPropertyValue -Path $keyPath -Name 'ElementName' -ErrorAction Stop)
            }
            catch {
                $elementName = '<kein ElementName>'
            }
        }

        Write-Host ("- {0} | Exists={1} | ElementName={2}" -f $svc.Id, $exists, $elementName)
    }
}

try {
    Assert-Administrator
    $selectedMode = Resolve-Mode

    if ($selectedMode -eq 'check') {
        Show-ServiceState
        Write-Host ''
        Write-Host '[DONE] Pruefung abgeschlossen (keine Aenderungen vorgenommen).'
        exit 0
    }

    if ($selectedMode -in @('remove', 'reset')) {
        Write-Host '[INFO] Entferne vorhandene HyperTool Hyper-V Socket Registry-Eintraege...'
        Remove-ServiceKeys
        Remove-LegacyServiceKeys
    }

    if ($selectedMode -in @('install', 'reset')) {
        Write-Host '[INFO] Pruefe auf unbekannte/alte HyperTool Registry-Eintraege...'
        Remove-UnknownHyperToolKeys
        Write-Host '[INFO] Installiere HyperTool Hyper-V Socket Registry-Eintraege neu...'
        Install-ServiceKeys
    }

    Show-ServiceState
    Write-Host ''
    Write-Host '[DONE] Fertig.'
    exit 0
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
finally {
    if (-not $NoPause) {
        Write-Host ''
        Write-Host 'Fenster schliessen mit Enter...'
        [void](Read-Host)
    }
}
