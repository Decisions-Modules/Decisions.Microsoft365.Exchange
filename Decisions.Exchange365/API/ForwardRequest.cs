using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Exchange365.API;

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