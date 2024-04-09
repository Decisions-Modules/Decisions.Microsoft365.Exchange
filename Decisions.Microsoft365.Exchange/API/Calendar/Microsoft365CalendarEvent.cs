using System;
using System.Runtime.Serialization;
using DecisionsFramework.Data.DataTypes;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.Calendar
{
    [Writable]
    public class Microsoft365CalendarEvent
    {
        [WritableValue]
        [JsonProperty("subject")]
        public string? Subject { get; set; }

        [WritableValue]
        [JsonProperty("body")]
        public Microsoft365EventBody? Body { get; set; }

        [WritableValue]
        [JsonProperty("start")]
        public Microsoft365DateTimeZone Start { get; set; }

        [WritableValue]
        [JsonProperty("end")]
        public Microsoft365DateTimeZone End { get; set; }
        
        [WritableValue]
        [JsonProperty("location")]
        public Microsoft365LocationName? Location { get; set; }
        
        [WritableValue]
        [JsonProperty("attendees")]
        public Microsoft365EventAttendee[]? Attendees { get; set; }
        
        [WritableValue]
        [JsonProperty("allowNewTimeProposals")]
        public bool? AllowNewTimeProposals { get; set; }
        [WritableValue]
        [JsonProperty("transactionId")]
        public string? TransactionId { get; set; }
    }
    
    [Writable]
    public class Microsoft365DateTimeZone
    {
        [WritableValue]
        [JsonProperty("dateTime")]
        public DateTime? DateTime { get; set; }

        [WritableValue]
        [JsonProperty("timeZone")]
        public string? TimeZone { get; set; }
    }

    [Writable]
    public class Microsoft365LocationName
    {
        [WritableValue]
        [JsonProperty("displayName")]
        public string? DisplayName { get; set; }
    }
    
    [Writable]
    public class Microsoft365ResponseStatus
    {
        [WritableValue]
        [JsonProperty("response")]
        public string? Response { get; set; }

        [WritableValue]
        [JsonProperty("time")]
        public DateTime? Time { get; set; }
    }

    [Writable]
    public class Microsoft365Recurrence
    {
        [WritableValue]
        [JsonProperty("pattern")]
        public Microsoft365Pattern? Pattern { get; set; }

        [WritableValue]
        [JsonProperty("range")]
        public Microsoft365Range? Range { get; set; }
    }
    
    [Writable]
    public class Microsoft365Pattern
    {
        [WritableValue]
        [JsonProperty("dayOfMonth")]
        public int? DayOfMonth { get; set; }

        [WritableValue]
        [JsonProperty("daysOfWeek")]
        public DayOfWeek[]? DaysOfWeek { get; set; }
        
        [WritableValue]
        [JsonProperty("firstDayOfWeek")]
        public DayOfWeek? FirstDayOfWeek { get; set; }

        [WritableValue]
        [JsonProperty("index")]
        public Microsoft365WeekIndex? Index { get; set; }
        
        [WritableValue]
        [JsonProperty("interval")]
        public int? Interval { get; set; }

        [WritableValue]
        [JsonProperty("month")]
        public int? Month { get; set; }
        
        [WritableValue]
        [JsonProperty("type")]
        public Microsoft365RecurrencePatternType? Type { get; set; }
    }
    
    [Writable]
    public class Microsoft365Range
    {
        [WritableValue]
        [JsonProperty("endDate")]
        public Date? EndDate { get; set; }

        [WritableValue]
        [JsonProperty("numberOfOccurrences")]
        public int? NumberOfOccurrences { get; set; }
        
        [WritableValue]
        [JsonProperty("recurrenceTimeZone")]
        public string? RecurrenceTimeZone { get; set; }
        
        [WritableValue]
        [JsonProperty("startDate")]
        public Date? StartDate { get; set; }
        
        [WritableValue]
        [JsonProperty("type")]
        public Microsoft365RecurrenceRangeType? Type { get; set; }
    }
    
    public enum Microsoft365WeekIndex
    {
        [EnumMember(Value = "first")]
        First,
        [EnumMember(Value = "second")]
        Second,
        [EnumMember(Value = "third")]
        Third,
        [EnumMember(Value = "fourth")]
        Fourth,
        [EnumMember(Value = "last")]
        Last
    }
    
    public enum Microsoft365RecurrencePatternType
    {
        [EnumMember(Value = "daily")]
        Daily,
        [EnumMember(Value = "weekly")]
        Weekly,
        [EnumMember(Value = "absoluteMonthly")]
        AbsoluteMonthly,
        [EnumMember(Value = "relativeMonthly")]
        RelativeMonthly,
        [EnumMember(Value = "absoluteYearly")]
        AbsoluteYearly,
        [EnumMember(Value = "relativeYearly")]
        RelativeYearly
    }

    public enum Microsoft365RecurrenceRangeType
    {
        [EnumMember(Value = "endDate")]
        EndDate,
        [EnumMember(Value = "noEnd")]
        NoEnd,
        [EnumMember(Value = "numbered")]
        Numbered
    }
}
