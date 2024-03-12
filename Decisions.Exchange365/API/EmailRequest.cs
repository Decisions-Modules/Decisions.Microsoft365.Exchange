using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Exchange365.API
{
    [Writable]
    public class EmailRequest
    {
        [WritableValue]
        [JsonProperty("Message")]
        public Message Message { get; set; }

        [WritableValue]
        [JsonProperty("SaveToSentItems")]
        public bool SaveToSentItems { get; set; }
    }
}