using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.Email
{
    [Writable]
    public class Microsoft365EmailReplyRequest
    {
        [WritableValue]
        [JsonProperty("message")]
        public Microsoft365SendEmailRequest? Message { get; set; }

        [WritableValue]
        [JsonProperty("comment")]
        public string? Comment { get; set; }
    }
}