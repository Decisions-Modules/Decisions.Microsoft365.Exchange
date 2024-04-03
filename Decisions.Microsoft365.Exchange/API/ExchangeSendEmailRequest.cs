using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API
{
    [Writable]
    public class ExchangeSendEmailRequest
    {
        [WritableValue]
        [JsonProperty("message")]
        public EmailMessage Message { get; set; }

        [WritableValue]
        [JsonProperty("saveToSentItems")]
        public bool SaveToSentItems { get; set; }
    }

    [Writable]
    public class EmailMessage
    {
        [WritableValue]
        [JsonProperty("subject")]
        public string Subject { get; set; }

        [WritableValue]
        [JsonProperty("body")]
        public Body Body { get; set; }

        [WritableValue]
        [JsonProperty("toRecipients")]
        public ExchangeRecipient[]? ToRecipients { get; set; }

        [WritableValue]
        [JsonProperty("ccRecipients")]
        public ExchangeRecipient[]? CcRecipients { get; set; }
    }

    [Writable]
    public class Body
    {
        [WritableValue]
        [JsonProperty("contentType")]
        public string ContentType { get; set; }

        [WritableValue]
        [JsonProperty("content")]
        public string? Content { get; set; }
    }

    [Writable]
    public class ExchangeRecipient
    {
        [WritableValue]
        [JsonProperty("emailAddress")]
        public ExchangeEmailAddress? EmailAddress { get; set; }
    }

    [Writable]
    public class ExchangeEmailAddress
    {
        [WritableValue]
        [JsonProperty("address")]
        public string? Address { get; set; }
    }
}
