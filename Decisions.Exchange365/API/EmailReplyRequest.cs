using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Exchange365.API
{
    [Writable]
    public class EmailReplyRequest
    {
        [WritableValue]
        [JsonProperty("message")]
        public SendEmailRequest Message { get; set; }

        [WritableValue]
        [JsonProperty("comment")]
        public string Comment { get; set; }
    }
}