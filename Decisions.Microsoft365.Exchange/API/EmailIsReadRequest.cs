using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API;

[Writable]
public class EmailIsReadRequest
{
    [WritableValue]
    [JsonProperty("isRead")]
    public bool IsRead { get; set; }
}