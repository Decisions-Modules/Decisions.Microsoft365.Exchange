using DecisionsFramework.Data.DataTypes;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API
{
    [Writable]
    public class CalendarEvent
    {
        [WritableValue]
        [JsonProperty("subject")]
        public string Subject { get; set; }

        [WritableValue]
        [JsonProperty("body")]
        public EventBody? Body { get; set; }

        [WritableValue]
        [JsonProperty("start")]
        public DateTimeZone Start { get; set; }

        [WritableValue]
        [JsonProperty("end")]
        public DateTimeZone End { get; set; }
        
        [WritableValue]
        [JsonProperty("location")]
        public Location? Location { get; set; }
        [WritableValue]
        [JsonProperty("attendees")]
        public Attendee[]? Attendees { get; set; }
        
        [WritableValue]
        [JsonProperty("allowNewTimeProposals")]
        public bool? AllowNewTimeProposals { get; set; }
        [WritableValue]
        [JsonProperty("transactionId")]
        public string? TransactionId { get; set; }
    }

    [Writable]
    public class Attendee
    {
        [WritableValue]
        [JsonProperty("emailAddress")]
        public EmailAddressName EmailAddress { get; set; }

        [WritableValue]
        [JsonProperty("type")]
        public string Type { get; set; }
    }

    public class EventBody
    {
        [WritableValue]
        [JsonProperty("contentType")]
        public string ContentType { get; set; }

        [WritableValue]
        [JsonProperty("content")]
        public string Content { get; set; }
    }

    [Writable]
    public class DateTimeZone
    {
        [WritableValue]
        [JsonProperty("dateTime")]
        public DateTime DateTime { get; set; }

        [WritableValue]
        [JsonProperty("timeZone")]
        public string TimeZone { get; set; }
    }

    [Writable]
    public class Location
    {
        [WritableValue]
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }
    }
    
    [Writable]
    public class ResponseStatus
    {
        [WritableValue]
        [JsonProperty("response")]
        public string Response { get; set; }

        [WritableValue]
        [JsonProperty("time")]
        public DateTime Time { get; set; }
    }

    [Writable]
    public class Recurrence
    {
        [WritableValue]
        [JsonProperty("pattern")]
        public Pattern Pattern { get; set; }

        [WritableValue]
        [JsonProperty("range")]
        public Range Range { get; set; }
    }
    
    [Writable]
    public class Pattern
    {
        [WritableValue]
        [JsonProperty("dayOfMonth")]
        public int DayOfMonth { get; set; }

        [WritableValue]
        [JsonProperty("daysOfWeek")]
        public DayOfWeek[] DaysOfWeek { get; set; }
        
        [WritableValue]
        [JsonProperty("firstDayOfWeek")]
        public DayOfWeek FirstDayOfWeek { get; set; }

        [WritableValue]
        [JsonProperty("index")]
        public WeekIndex Index { get; set; }
        
        [WritableValue]
        [JsonProperty("interval")]
        public int Interval { get; set; }

        [WritableValue]
        [JsonProperty("month")]
        public int Month { get; set; }
        
        [WritableValue]
        [JsonProperty("type")]
        public RecurrencePatternType Type { get; set; }
    }
    
    [Writable]
    public class Range
    {
        [WritableValue]
        [JsonProperty("endDate")]
        public Date EndDate { get; set; }

        [WritableValue]
        [JsonProperty("numberOfOccurrences")]
        public int NumberOfOccurrences { get; set; }
        
        [WritableValue]
        [JsonProperty("recurrenceTimeZone")]
        public string RecurrenceTimeZone { get; set; }
        
        [WritableValue]
        [JsonProperty("startDate")]
        public Date StartDate { get; set; }
        
        [WritableValue]
        [JsonProperty("type")]
        public RecurrenceRangeType Type { get; set; }
    }
}
