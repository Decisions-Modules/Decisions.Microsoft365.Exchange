using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Exchange365.API
{
    [Writable]
    public class SendEmailRequest
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
        public Recipient[] ToRecipients { get; set; }

        [WritableValue]
        [JsonProperty("ccRecipients")]
        public Recipient[] CcRecipients { get; set; }
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
    public class Recipient
    {
        [WritableValue]
        [JsonProperty("emailAddress")]
        public EmailAddress EmailAddress { get; set; }
    }

    [Writable]
    public class EmailAddress
    {
        [WritableValue]
        [JsonProperty("address")]
        public string Address { get; set; }
    }
}
