using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Exchange365.API;

[Writable]
public class EmailIsReadRequest
{
    [WritableValue]
    [JsonProperty("isRead")]
    public bool IsRead { get; set; }
}