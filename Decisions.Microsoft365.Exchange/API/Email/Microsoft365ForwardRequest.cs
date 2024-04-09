using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.Email;

[Writable]
public class Microsoft365ForwardRequest
{
    [WritableValue]
    [JsonProperty("comment")]
    public string? Comment { get; set; }

    [WritableValue]
    [JsonProperty("toRecipients")]
    public Microsoft365Recipient[]? ToRecipients { get; set; }
}