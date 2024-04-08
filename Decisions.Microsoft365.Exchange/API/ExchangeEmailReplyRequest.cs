using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API
{
    [Writable]
    public class ExchangeEmailReplyRequest
    {
        [WritableValue]
        [JsonProperty("message")]
        public ExchangeSendEmailRequest? Message { get; set; }

        [WritableValue]
        [JsonProperty("comment")]
        public string? Comment { get; set; }
    }
}