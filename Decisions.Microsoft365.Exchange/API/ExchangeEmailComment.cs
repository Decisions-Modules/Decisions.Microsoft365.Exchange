using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API;

[Writable]
public class ExchangeEmailComment
{
    [WritableValue]
    [JsonProperty("comment")]
    public string Comment { get; set; }
}