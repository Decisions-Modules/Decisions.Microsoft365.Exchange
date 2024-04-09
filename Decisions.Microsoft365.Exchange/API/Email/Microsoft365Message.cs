using System;
using System.Runtime.Serialization;
using DecisionsFramework;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.Email
{
    [Writable]
    public class Microsoft365Message
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
        public Microsoft365EmailBody? Body { get; set; }

        [WritableValue]
        [JsonProperty("sender")]
        public Microsoft365EmailFrom? Sender { get; set; }

        [WritableValue]
        [JsonProperty("from")]
        public Microsoft365EmailFrom? From { get; set; }

        [WritableValue]
        [JsonProperty("toRecipients")]
        public Microsoft365EmailFrom[]? ToRecipients { get; set; }

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
        public Microsoft365EmailFlag? Flag { get; set; }
        
        public static Microsoft365Message? JsonDeserialize(string content)
        {
            try
            {
                return JsonConvert.DeserializeObject<Microsoft365Message>(content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("Could not deserialize result.", ex);
            }
        }
    }

    [Writable]
    public class Microsoft365EmailBody
    {
        [WritableValue]
        [JsonProperty("contentType")]
        public string? ContentType { get; set; }

        [WritableValue]
        [JsonProperty("content")]
        public string? Content { get; set; }
    }

    [Writable]
    public class Microsoft365EmailFlag
    {
        [WritableValue]
        [JsonProperty("flagStatus")]
        public string? FlagStatus { get; set; }
    }

    [Writable]
    public class Microsoft365EmailFrom
    {
        [WritableValue]
        [JsonProperty("emailAddress")]
        public Microsoft365EmailAddress? EmailAddress { get; set; }
    }

    [Writable]
    public class Microsoft365EmailAddress
    {
        [WritableValue]
        [JsonProperty("name")]
        public string? Name { get; set; }

        [WritableValue]
        [JsonProperty("address")]
        public string? Address { get; set; }
    }
    
    public enum Microsoft365BodyType {
        [EnumMember(Value = "text")]
        Text,
        [EnumMember(Value = "html")]
        Html,
    }
}
