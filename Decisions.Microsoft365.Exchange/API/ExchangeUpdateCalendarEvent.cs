using DecisionsFramework;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API
{
    [Writable]
    public class ExchangeUpdateCalendarEvent
    {
        [WritableValue]
        [JsonProperty("subject")]
        public string? Subject { get; set; }

        [WritableValue]
        [JsonProperty("body")]
        public EventBody? Body { get; set; }

        [WritableValue]
        [JsonProperty("start")]
        public DateTimeZone? Start { get; set; }

        [WritableValue]
        [JsonProperty("end")]
        public DateTimeZone? End { get; set; }
        
        [WritableValue]
        [JsonProperty("location")]
        public Location? Location { get; set; }
        
        [WritableValue]
        [JsonProperty("locations")]
        public Location[]? Locations { get; set; }

        [WritableValue]
        [JsonProperty("attendees")]
        public Attendee[]? Attendees { get; set; }
        
        [WritableValue]
        [JsonProperty("allowNewTimeProposals")]
        public bool? AllowNewTimeProposals { get; set; }
        
        [WritableValue]
        [JsonProperty("recurrence")]
        public Recurrence? Recurrence { get; set; }
        
        [WritableValue]
        [JsonProperty("reminderMinutesBeforeStart")]
        public int? ReminderMinutesBeforeStart { get; set; }
        
        [WritableValue]
        [JsonProperty("isOnlineMeeting")]
        public bool? IsOnlineMeeting { get; set; }
        
        [WritableValue]
        [JsonProperty("onlineMeetingProvider")]
        public string? OnlineMeetingProvider { get; set; }
        
        [WritableValue]
        [JsonProperty("isAllDay")]
        public bool? IsAllDay { get; set; }
        
        [WritableValue]
        [JsonProperty("isReminderOn")]
        public bool? IsReminderOn { get; set; }
        
        [WritableValue]
        [JsonProperty("hideAttendees")]
        public bool? HideAttendees { get; set; }
        
        [WritableValue]
        [JsonProperty("categories")]
        public string[]? Categories { get; set; }
        
        [WritableValue]
        [JsonProperty("sensitivity")]
        public string? Sensitivity { get; set; }
        
        [WritableValue]
        [JsonProperty("importance")]
        public string? Importance { get; set; }
        
        [WritableValue]
        [JsonProperty("showAs")]
        public string? ShowAs { get; set; }
        
        [WritableValue]
        [JsonProperty("responseRequested")]
        public bool? ResponseRequested { get; set; }
    }
}