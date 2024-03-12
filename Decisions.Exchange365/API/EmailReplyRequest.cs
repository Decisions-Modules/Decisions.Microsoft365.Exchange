using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Exchange365.API
{
    [Writable]
    public class EmailReplyRequest
    {
        [WritableValue]
        [JsonProperty("Message")]
        public Message Message { get; set; }

        [WritableValue]
        [JsonProperty("Comment")]
        public string Comment { get; set; }
    }
}