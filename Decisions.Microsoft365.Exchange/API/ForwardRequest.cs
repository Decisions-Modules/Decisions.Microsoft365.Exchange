using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API;

[Writable]
public class ForwardRequest
{
    [WritableValue]
    [JsonProperty("comment")]
    public string Comment { get; set; }

    [WritableValue]
    [JsonProperty("toRecipients")]
    public Recipient[] ToRecipients { get; set; }
}