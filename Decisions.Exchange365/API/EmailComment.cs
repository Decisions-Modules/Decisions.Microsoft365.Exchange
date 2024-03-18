using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Exchange365.API;

[Writable]
public class EmailComment
{
    [WritableValue]
    [JsonProperty("comment")]
    public string Comment { get; set; }
}