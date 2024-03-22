using DecisionsFramework.Data.DataTypes;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Exchange365.API
{
    public class CalendarEvent
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
        
        [JsonProperty("locations")]
        public Location[] Locations { get; set; }

        [JsonProperty("attendees")]
        public Attendee[] Attendees { get; set; }
        
        [JsonProperty("responseStatus")]
        public Microsoft.Graph.Models.ResponseStatus ResponseStatus { get; set; }
        
        [JsonProperty("recurrence")]
        public Recurrence Recurrence { get; set; }
        
        [JsonProperty("reminderMinutesBeforeStart")]
        public int ReminderMinutesBeforeStart { get; set; }
        
        [JsonProperty("isOnlineMeeting")]
        public bool IsOnlineMeeting { get; set; }
        
        [JsonProperty("onlineMeetingProvider")]
        public string OnlineMeetingProvider { get; set; }
        
        [JsonProperty("isAllDay")]
        public bool IsAllDay { get; set; }
        
        [JsonProperty("isReminderOn")]
        public bool IsReminderOn { get; set; }
        
        [JsonProperty("hideAttendees")]
        public bool HideAttendees { get; set; }
        
        [JsonProperty("categories")]
        public string[] Categories { get; set; }
        
        [JsonProperty("sensitivity")]
        public string Sensitivity { get; set; }
        
        [JsonProperty("importance")]
        public string Importance { get; set; }
        
        [JsonProperty("showAs")]
        public string ShowAs { get; set; }
        
        [JsonProperty("responseRequested")]
        public bool ResponseRequested { get; set; }
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
    
    public class ResponseStatus
    {
        [JsonProperty("response")]
        public string Response { get; set; }

        [JsonProperty("time")]
        public DateTime Time { get; set; }
    }

    public class Recurrence
    {
        [JsonProperty("pattern")]
        public Pattern Pattern { get; set; }

        [JsonProperty("range")]
        public Range Range { get; set; }
    }
    
    public class Pattern
    {
        [JsonProperty("dayOfMonth")]
        public int DayOfMonth { get; set; }

        [JsonProperty("daysOfWeek")]
        public DayOfWeek[] DaysOfWeek { get; set; }
        
        [JsonProperty("firstDayOfWeek")]
        public DayOfWeek FirstDayOfWeek { get; set; }

        [JsonProperty("index")]
        public WeekIndex Index { get; set; }
        
        [JsonProperty("interval")]
        public int Interval { get; set; }

        [JsonProperty("month")]
        public int Month { get; set; }
        
        [JsonProperty("type")]
        public RecurrencePatternType Type { get; set; }
    }
    
    public class Range
    {
        [JsonProperty("endDate")]
        public Date EndDate { get; set; }

        [JsonProperty("numberOfOccurrences")]
        public int NumberOfOccurrences { get; set; }
        
        [JsonProperty("recurrenceTimeZone")]
        public string RecurrenceTimeZone { get; set; }
        
        [JsonProperty("startDate")]
        public Date StartDate { get; set; }
        
        [JsonProperty("type")]
        public RecurrenceRangeType Type { get; set; }
    }
}
