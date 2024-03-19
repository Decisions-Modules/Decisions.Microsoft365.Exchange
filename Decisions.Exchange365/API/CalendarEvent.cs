using Newtonsoft.Json;

namespace Decisions.Exchange365.API
{
    public partial class CalendarEvent
    {
        [JsonProperty("subject")]
        public string Subject { get; set; }

        [JsonProperty("body")]
        public EventBody Body { get; set; }

        [JsonProperty("start")]
        public DateTimeZone Start { get; set; }

        [JsonProperty("end")]
        public DateTimeZone End { get; set; }

        [JsonProperty("location")]
        public Location Location { get; set; }

        [JsonProperty("attendees")]
        public Attendee[] Attendees { get; set; }

        [JsonProperty("allowNewTimeProposals")]
        public bool AllowNewTimeProposals { get; set; }

        [JsonProperty("transactionId")]
        public string TransactionId { get; set; }
    }

    public class Attendee
    {
        [JsonProperty("emailAddress")]
        public EmailAddress EmailAddress { get; set; }

        [JsonProperty("type")]
        public string Type { get; set; }
    }

    public class EventBody
    {
        [JsonProperty("contentType")]
        public string ContentType { get; set; }

        [JsonProperty("content")]
        public string Content { get; set; }
    }

    public class DateTimeZone
    {
        [JsonProperty("dateTime")]
        public DateTime DateTime { get; set; }

        [JsonProperty("timeZone")]
        public string TimeZone { get; set; }
    }

    public partial class Location
    {
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }
    }
}
