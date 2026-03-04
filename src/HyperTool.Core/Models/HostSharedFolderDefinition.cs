namespace HyperTool.Models;

public sealed class HostSharedFolderDefinition : System.ComponentModel.INotifyPropertyChanged
{
    private string _id = Guid.NewGuid().ToString("N");
    private string _label = string.Empty;
    private string _localPath = string.Empty;
    private string _shareName = string.Empty;
    private bool _enabled = true;
    private bool _readOnly;

    public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;

    public string Id
    {
        get => _id;
        set
        {
            if (string.Equals(_id, value, StringComparison.Ordinal))
            {
                return;
            }

            _id = value;
            PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(Id)));
        }
    }

    public string Label
    {
        get => _label;
        set
        {
            if (string.Equals(_label, value, StringComparison.Ordinal))
            {
                return;
            }

            _label = value;
            PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(Label)));
        }
    }

    public string LocalPath
    {
        get => _localPath;
        set
        {
            if (string.Equals(_localPath, value, StringComparison.Ordinal))
            {
                return;
            }

            _localPath = value;
            PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(LocalPath)));
        }
    }

    public string ShareName
    {
        get => _shareName;
        set
        {
            if (string.Equals(_shareName, value, StringComparison.Ordinal))
            {
                return;
            }

            _shareName = value;
            PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(ShareName)));
        }
    }

    public bool Enabled
    {
        get => _enabled;
        set
        {
            if (_enabled == value)
            {
                return;
            }

            _enabled = value;
            PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(Enabled)));
        }
    }

    public bool ReadOnly
    {
        get => _readOnly;
        set
        {
            if (_readOnly == value)
            {
                return;
            }

            _readOnly = value;
            PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(ReadOnly)));
            PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(AccessModeLabel)));
        }
    }

    public string AccessModeLabel => ReadOnly ? "Lesezugriff" : "Lese-/Schreibzugriff";
}
