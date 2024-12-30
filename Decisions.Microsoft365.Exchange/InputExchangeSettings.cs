using System.ComponentModel;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.Properties.Attributes;

namespace Decisions.Microsoft365.Exchange;

public class InputExchangeSettings: IEntityPickerLocation
{
    private string graphUrl = "https://graph.microsoft.com/v1.0";

    [PropertyClassification(0, "Graph URL", "Exchange Settings")]
    [WritableValue]
    public string GraphUrl
    {
        get => graphUrl;
        set
        {
            graphUrl = value.TrimEnd('/', '\\');
            OnPropertyChanged(nameof(GraphUrl));
        }
    }
    
    private string? tokenId;

    [WritableValue]
    [PropertyClassification(new string[] { "Credentials" }, "OAuth Token", 1)]
    [TokenPicker]
    public string? TokenId
    {
        get => tokenId;
        set
        {
            tokenId = value;
            OnPropertyChanged(nameof(TokenId));
        }
    }
    
    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private string createChildEntityInFolderId;

    [PropertyHidden(hiddenInMapping: true)]
    public string CreateChildEntityInFolderId
    {
        get => createChildEntityInFolderId;
        set
        {
            createChildEntityInFolderId = value;
        }
    }
}