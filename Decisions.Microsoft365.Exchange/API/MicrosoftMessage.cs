using System.Runtime.Serialization;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API
{
    [Writable]
    public class MicrosoftMessage
    {
        [WritableValue]
        [JsonProperty("@odata.context")]
        public string? OdataContext { get; set; }

        [WritableValue]
        [JsonProperty("@odata.etag")]
        public string? OdataEtag { get; set; }

        [WritableValue]
        [JsonProperty("id")]
        public string? Id { get; set; }

        [WritableValue]
        [JsonProperty("createdDateTime")]
        public DateTimeOffset? CreatedDateTime { get; set; }
        
        [WritableValue]
        [JsonProperty("lastModifiedDateTime")]
        public DateTimeOffset? LastModifiedDateTime { get; set; }

        [WritableValue]
        [JsonProperty("changeKey")]
        public string? ChangeKey { get; set; }

        [WritableValue]
        [JsonProperty("categories")]
        public string[]? Categories { get; set; }

        [WritableValue]
        [JsonProperty("receivedDateTime")]
        public DateTimeOffset? ReceivedDateTime { get; set; }

        [WritableValue]
        [JsonProperty("sentDateTime")]
        public DateTimeOffset? SentDateTime { get; set; }

        [WritableValue]
        [JsonProperty("hasAttachments")]
        public bool? HasAttachments { get; set; }

        [WritableValue]
        [JsonProperty("internetMessageId")]
        public string? InternetMessageId { get; set; }

        [WritableValue]
        [JsonProperty("subject")]
        public string? Subject { get; set; }

        [WritableValue]
        [JsonProperty("bodyPreview")]
        public string? BodyPreview { get; set; }

        [WritableValue]
        [JsonProperty("importance")]
        public string? Importance { get; set; }

        [WritableValue]
        [JsonProperty("parentFolderId")]
        public string? ParentFolderId { get; set; }

        [WritableValue]
        [JsonProperty("conversationId")]
        public string? ConversationId { get; set; }

        [WritableValue]
        [JsonProperty("conversationIndex")]
        public string? ConversationIndex { get; set; }

        [WritableValue]
        [JsonProperty("isDeliveryReceiptRequested")]
        public bool? IsDeliveryReceiptRequested { get; set; }

        [WritableValue]
        [JsonProperty("isReadReceiptRequested")]
        public bool? IsReadReceiptRequested { get; set; }

        [WritableValue]
        [JsonProperty("isRead")]
        public bool? IsRead { get; set; }

        [WritableValue]
        [JsonProperty("isDraft")]
        public bool? IsDraft { get; set; }

        [WritableValue]
        [JsonProperty("webLink")]
        public string? WebLink { get; set; }

        [WritableValue]
        [JsonProperty("inferenceClassification")]
        public string? InferenceClassification { get; set; }

        [WritableValue]
        [JsonProperty("body")]
        public MicrosoftEmailBody? Body { get; set; }

        [WritableValue]
        [JsonProperty("sender")]
        public MicrosoftEmailFrom? Sender { get; set; }

        [WritableValue]
        [JsonProperty("from")]
        public MicrosoftEmailFrom? From { get; set; }

        [WritableValue]
        [JsonProperty("toRecipients")]
        public MicrosoftEmailFrom[]? ToRecipients { get; set; }

        [WritableValue]
        [JsonProperty("ccRecipients")]
        public string[]? CcRecipients { get; set; }

        [WritableValue]
        [JsonProperty("bccRecipients")]
        public string[]? BccRecipients { get; set; }

        [WritableValue]
        [JsonProperty("replyTo")]
        public string[]? ReplyTo { get; set; }

        [WritableValue]
        [JsonProperty("flag")]
        public MicrosoftEmailFlag? Flag { get; set; }
    }

    [Writable]
    public class MicrosoftEmailBody
    {
        [WritableValue]
        [JsonProperty("contentType")]
        public string? ContentType { get; set; }

        [WritableValue]
        [JsonProperty("content")]
        public string? Content { get; set; }
    }

    [Writable]
    public class MicrosoftEmailFlag
    {
        [WritableValue]
        [JsonProperty("flagStatus")]
        public string? FlagStatus { get; set; }
    }

    [Writable]
    public class MicrosoftEmailFrom
    {
        [WritableValue]
        [JsonProperty("emailAddress")]
        public MicrosoftEmailAddress? EmailAddress { get; set; }
    }

    [Writable]
    public class MicrosoftEmailAddress
    {
        [WritableValue]
        [JsonProperty("name")]
        public string? Name { get; set; }

        [WritableValue]
        [JsonProperty("address")]
        public string? Address { get; set; }
    }
    
    public enum MicrosoftBodyType {
        [EnumMember(Value = "text")]
        Text,
        [EnumMember(Value = "html")]
        Html,
    }
}
