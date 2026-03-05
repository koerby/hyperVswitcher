namespace HyperTool.Models;

public sealed class HyperToolConfig
{
    public string DefaultVmName { get; set; } = string.Empty;

    public string LastSelectedVmName { get; set; } = string.Empty;

    public List<VmDefinition> Vms { get; set; } = [];

    public string DefaultSwitchName { get; set; } = "Default Switch";

    public string VmConnectComputerName { get; set; } = Environment.MachineName;

    public HnsSettings Hns { get; set; } = new();

    public UiSettings Ui { get; set; } = new();

    public UpdateSettings Update { get; set; } = new();

    public UsbSettings Usb { get; set; } = new();

    public SharedFolderSettings SharedFolders { get; set; } = new();

    public static HyperToolConfig CreateDefault() => new();
}

public sealed class HnsSettings
{
    public bool Enabled { get; set; }

    public bool AutoRestartAfterDefaultSwitch { get; set; }

    public bool AutoRestartAfterAnyConnect { get; set; }
}

public sealed class UiSettings
{
    public string WindowTitle { get; set; } = "HyperTool";

    public string Theme { get; set; } = "Dark";

    public bool StartMinimized { get; set; }

    public bool MinimizeToTray { get; set; } = true;

    public bool EnableTrayIcon { get; set; } = true;

    public bool EnableTrayMenu { get; set; } = true;

    public bool StartWithWindows { get; set; } = true;

    public bool OpenConsoleAfterVmStart { get; set; } = true;

    public bool OpenVmConnectWithSessionEdit { get; set; }

    public List<string> TrayVmNames { get; set; } = [];
}

public sealed class UpdateSettings
{
    public bool CheckOnStartup { get; set; } = true;

    public string GitHubOwner { get; set; } = "koerby";

    public string GitHubRepo { get; set; } = "HyperTool";
}

public sealed class UsbSettings
{
    public bool Enabled { get; set; } = true;

    public bool AutoDetachOnClientDisconnect { get; set; } = true;

    public bool UnshareOnExit { get; set; } = true;

    public List<string> AutoShareDeviceKeys { get; set; } = [];
}

public sealed class SharedFolderSettings
{
    public bool Enabled { get; set; } = true;

    public List<HostSharedFolderDefinition> HostDefinitions { get; set; } = [];
}