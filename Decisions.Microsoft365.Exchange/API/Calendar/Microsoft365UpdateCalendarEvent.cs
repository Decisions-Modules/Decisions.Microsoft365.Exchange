using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.Calendar
{
    [Writable]
    public class Microsoft365UpdateCalendarEvent
    {
        [WritableValue]
        [JsonProperty("subject")]
        public string? Subject { get; set; }

        [WritableValue]
        [JsonProperty("body")]
        public Microsoft365EventBody? Body { get; set; }

        [WritableValue]
        [JsonProperty("start")]
        public Microsoft365DateTimeZone? Start { get; set; }

        [WritableValue]
        [JsonProperty("end")]
        public Microsoft365DateTimeZone? End { get; set; }
        
        [WritableValue]
        [JsonProperty("location")]
        public Microsoft365LocationName? Location { get; set; }
        
        [WritableValue]
        [JsonProperty("locations")]
        public Microsoft365LocationName[]? Locations { get; set; }

        [WritableValue]
        [JsonProperty("attendees")]
        public Microsoft365EventAttendee[]? Attendees { get; set; }
        
        [WritableValue]
        [JsonProperty("allowNewTimeProposals")]
        public bool? AllowNewTimeProposals { get; set; }
        
        [WritableValue]
        [JsonProperty("recurrence")]
        public Microsoft365Recurrence? Recurrence { get; set; }
        
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