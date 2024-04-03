using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API;

[Writable]
public class ExchangeForwardRequest
{
    [WritableValue]
    [JsonProperty("comment")]
    public string Comment { get; set; }

    [WritableValue]
    [JsonProperty("toRecipients")]
    public ExchangeRecipient[] ToRecipients { get; set; }
}