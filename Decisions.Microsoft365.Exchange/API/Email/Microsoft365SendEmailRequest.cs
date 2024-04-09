using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.Email
{
    [Writable]
    public class Microsoft365SendEmailRequest
    {
        [WritableValue]
        [JsonProperty("message")]
        public Microsoft365EmailMessage? Message { get; set; }

        [WritableValue]
        [JsonProperty("saveToSentItems")]
        public bool? SaveToSentItems { get; set; }
    }

    [Writable]
    public class Microsoft365EmailMessage
    {
        [WritableValue]
        [JsonProperty("subject")]
        public string? Subject { get; set; }

        [WritableValue]
        [JsonProperty("body")]
        public Microsoft365EmailBody? Body { get; set; }

        [WritableValue]
        [JsonProperty("toRecipients")]
        public Microsoft365Recipient[]? ToRecipients { get; set; }

        [WritableValue]
        [JsonProperty("ccRecipients")]
        public Microsoft365Recipient[]? CcRecipients { get; set; }
    }

    [Writable]
    public class Microsoft365Recipient
    {
        [WritableValue]
        [JsonProperty("emailAddress")]
        public Microsoft365Address? EmailAddress { get; set; }
    }

    [Writable]
    public class Microsoft365Address
    {
        [WritableValue]
        [JsonProperty("address")]
        public string? Address { get; set; }
    }
}
