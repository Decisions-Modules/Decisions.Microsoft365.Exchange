using System;
using Decisions.Microsoft365.Exchange.API.Email;
using DecisionsFramework;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.Calendar
{
    [Writable]
    public class Microsoft365Event
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
        [JsonProperty("transactionId")]
        public string? TransactionId { get; set; }

        [WritableValue]
        [JsonProperty("originalStartTimeZone")]
        public string? OriginalStartTimeZone { get; set; }

        [WritableValue]
        [JsonProperty("originalEndTimeZone")]
        public string? OriginalEndTimeZone { get; set; }

        [WritableValue]
        [JsonProperty("iCalUId")]
        public string? ICalUId { get; set; }

        [WritableValue]
        [JsonProperty("reminderMinutesBeforeStart")]
        public int? ReminderMinutesBeforeStart { get; set; }

        [WritableValue]
        [JsonProperty("isReminderOn")]
        public bool? IsReminderOn { get; set; }

        [WritableValue]
        [JsonProperty("hasAttachments")]
        public bool? HasAttachments { get; set; }

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
        [JsonProperty("sensitivity")]
        public string? Sensitivity { get; set; }

        [WritableValue]
        [JsonProperty("isAllDay")]
        public bool? IsAllDay { get; set; }

        [WritableValue]
        [JsonProperty("isCancelled")]
        public bool? IsCancelled { get; set; }

        [WritableValue]
        [JsonProperty("isOrganizer")]
        public bool? IsOrganizer { get; set; }

        [WritableValue]
        [JsonProperty("responseRequested")]
        public bool? ResponseRequested { get; set; }

        [WritableValue]
        [JsonProperty("seriesMasterId")]
        public string? SeriesMasterId { get; set; }

        [WritableValue]
        [JsonProperty("showAs")]
        public string? ShowAs { get; set; }

        [WritableValue]
        [JsonProperty("type")]
        public string? Type { get; set; }

        [WritableValue]
        [JsonProperty("webLink")]
        public string? WebLink { get; set; }

        [WritableValue]
        [JsonProperty("onlineMeetingUrl")]
        public string? OnlineMeetingUrl { get; set; }

        [WritableValue]
        [JsonProperty("isOnlineMeeting")]
        public bool? IsOnlineMeeting { get; set; }

        [WritableValue]
        [JsonProperty("onlineMeetingProvider")]
        public string? OnlineMeetingProvider { get; set; }

        [WritableValue]
        [JsonProperty("allowNewTimeProposals")]
        public bool? AllowNewTimeProposals { get; set; }

        [WritableValue]
        [JsonProperty("occurrenceId")]
        public string? OccurrenceId { get; set; }

        [WritableValue]
        [JsonProperty("isDraft")]
        public bool? IsDraft { get; set; }

        [WritableValue]
        [JsonProperty("hideAttendees")]
        public bool? HideAttendees { get; set; }

        [WritableValue]
        [JsonProperty("responseStatus")]
        public Microsoft365EventStatus? ResponseStatus { get; set; }

        [WritableValue]
        [JsonProperty("body")]
        public Microsoft365EventBody? Body { get; set; }

        [WritableValue]
        [JsonProperty("start")]
        public Microsoft365EventTime? Start { get; set; }

        [WritableValue]
        [JsonProperty("end")]
        public Microsoft365EventTime? End { get; set; }

        [WritableValue]
        [JsonProperty("location")]
        public Microsoft365LocationName? Location { get; set; }

        [WritableValue]
        [JsonProperty("locations")]
        public Microsoft365LocationName[]? Locations { get; set; }

        [WritableValue]
        [JsonProperty("recurrence")]
        public string? Recurrence { get; set; }

        [WritableValue]
        [JsonProperty("attendees")]
        public Microsoft365EventAttendee[]? Attendees { get; set; }

        [WritableValue]
        [JsonProperty("organizer")]
        public Microsoft365EventOrganizer? Organizer { get; set; }

        [WritableValue]
        [JsonProperty("onlineMeeting")]
        public string? OnlineMeeting { get; set; }
        
        public static Microsoft365Event? JsonDeserialize(string content)
        {
            try
            {
                return JsonConvert.DeserializeObject<Microsoft365Event>(content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("Could not deserialize result.", ex);
            }
        }
    }

    [Writable]
    public class Microsoft365EventAttendee
    {
        [WritableValue]
        [JsonProperty("type")]
        public string? Type { get; set; }

        [WritableValue]
        [JsonProperty("status")]
        public Microsoft365EventStatus? Status { get; set; }

        [WritableValue]
        [JsonProperty("emailAddress")]
        public Microsoft365EmailAddress? EmailAddress { get; set; }
    }

    [Writable]
    public class Microsoft365EventStatus
    {
        [WritableValue]
        [JsonProperty("response")]
        public string? Response { get; set; }

        [WritableValue]
        [JsonProperty("time")]
        public DateTimeOffset? Time { get; set; }
    }

    [Writable]
    public class Microsoft365EventBody
    {
        [WritableValue]
        [JsonProperty("contentType")]
        public string? ContentType { get; set; }

        [WritableValue]
        [JsonProperty("content")]
        public string? Content { get; set; }
    }

    [Writable]
    public class Microsoft365EventTime
    {
        [WritableValue]
        [JsonProperty("dateTime")]
        public DateTimeOffset? DateTime { get; set; }

        [WritableValue]
        [JsonProperty("timeZone")]
        public string? TimeZone { get; set; }
    }

    [Writable]
    public class Microsoft365Location
    {
        [WritableValue]
        [JsonProperty("displayName")]
        public string? DisplayName { get; set; }

        [WritableValue]
        [JsonProperty("locationType")]
        public string? LocationType { get; set; }

        [WritableValue]
        [JsonProperty("uniqueId")]
        public string? UniqueId { get; set; }

        [WritableValue]
        [JsonProperty("uniqueIdType")]
        public string? UniqueIdType { get; set; }
    }

    [Writable]
    public class Microsoft365EventOrganizer
    {
        [WritableValue]
        [JsonProperty("emailAddress")]
        public Microsoft365EmailAddress? EmailAddress { get; set; }
    }
}