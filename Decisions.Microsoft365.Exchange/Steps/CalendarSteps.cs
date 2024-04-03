using System.Net.Http.Json;
using System.Text;
using Decisions.Microsoft365.Exchange.API;
using DecisionsFramework.Design.Flow;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using Attendee = Decisions.Microsoft365.Exchange.API.Attendee;
using Location = Decisions.Microsoft365.Exchange.API.Location;

namespace Decisions.Microsoft365.Exchange.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Microsoft365/Exchange/Calendar")]
    public class CalendarSteps
    {
        public Event? CreateCalendarEvent(CalendarEvent calendarEvent, string userIdentifier, string? calendarId)
        {
            string urlExtension = $"/users/{userIdentifier}";
            urlExtension = (!string.IsNullOrEmpty(calendarId)) ? $"{urlExtension}/calendars/{calendarId}/events"
                    : $"{urlExtension}/calendar/events";
            
            JsonContent content = JsonContent.Create(calendarEvent);
            
            return JsonConvert.DeserializeObject<Event>(GraphRest.Post(urlExtension, content));
        }

        public string DeleteCalendarEvent(string userIdentifier, string eventId, string? calendarId, string? calendarGroupId)
        {
            string urlExtension = $"/users/{userIdentifier}";
            urlExtension = (!string.IsNullOrEmpty(calendarId) && !string.IsNullOrEmpty(calendarGroupId))
                ? $"{urlExtension}/calendarGroups/{calendarGroupId}/calendars/{calendarId}/events/{eventId}"
                : (!string.IsNullOrEmpty(calendarId)) ? $"{urlExtension}/calendars/{calendarId}/events/{eventId}"
                    : $"{urlExtension}/events/{eventId}";

            return GraphRest.Delete(urlExtension).StatusCode.ToString();
        }
        
        public EventList ListCalendarEvents(string userIdentifier, string? calendarId, string? calendarGroupId)
        {
            string urlExtension = $"/users/{userIdentifier}";
            if (!string.IsNullOrEmpty(calendarId))
            {
                urlExtension = (!string.IsNullOrEmpty(calendarGroupId)) ? $"{urlExtension}/calendarGroups/{calendarGroupId}/calendars/{calendarId}"
                    : $"{urlExtension}/calendars/{calendarId}";
            }
            urlExtension += "/events";

            string result = GraphRest.Get(urlExtension);
            return JsonConvert.DeserializeObject<EventList>(result) ?? new EventList();
        }

        public Event? UpdateCalendarEvent(string userIdentifier, string eventId, string? subject, EventBody? body, DateTimeZone? startTime,
            DateTimeZone? endTime, Location? location, Attendee[]? attendees, bool? allowNewTimeProposals, int? reminderMinutesBeforeStart,
            bool? isOnlineMeeting, string onlineMeetingProvider, bool? isAllDay, bool? isReminderOn, bool? hideAttendees,
            string[]? categories, Sensitivity? sensitivity, Importance? importance, AttendeeStatus? showAs, bool? responseRequested)
        {
            string urlExtension = $"/users/{userIdentifier}/calendar/events/{eventId}";

            string? sensitivityString = (sensitivity != null) ? sensitivity.ToString() : null;
            string? importanceString = (importance != null) ? importance.ToString() : null;
            string? statusString = (showAs != null) ? showAs.ToString() : null;

            UpdateCalendarEvent calendarEventUpdate = new()
            {
                Subject = subject,
                Body = body,
                Start = startTime,
                End = endTime,
                Location = location,
                Attendees = attendees,
                AllowNewTimeProposals = allowNewTimeProposals,
                ReminderMinutesBeforeStart = reminderMinutesBeforeStart,
                IsOnlineMeeting = isOnlineMeeting,
                OnlineMeetingProvider = onlineMeetingProvider,
                IsAllDay = isAllDay,
                IsReminderOn = isReminderOn,
                HideAttendees = hideAttendees,
                Categories = categories,
                Sensitivity = sensitivityString,
                Importance = importanceString,
                ShowAs = statusString,
                ResponseRequested = responseRequested
            };
            
            HttpContent content = new StringContent(JsonConvert.SerializeObject(calendarEventUpdate, new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore
            }), Encoding.UTF8, "application/json");

            return JsonConvert.DeserializeObject<Event>(GraphRest.Patch(urlExtension, content));
        }
        
        public CalendarList ListCalendars(string userIdentifier)
        {
            string urlExtension = $"/users/{userIdentifier}/calendars";

            string result = GraphRest.Get(urlExtension);
            return JsonConvert.DeserializeObject<CalendarList>(result) ?? new CalendarList();
        }
    }
}