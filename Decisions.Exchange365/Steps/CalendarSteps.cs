using System.Net.Http.Json;
using System.Text;
using Decisions.Exchange365.API;
using Decisions.Exchange365.Data;
using DecisionsFramework.Design.Flow;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using Attendee = Decisions.Exchange365.API.Attendee;
using Location = Decisions.Exchange365.API.Location;

namespace Decisions.Exchange365.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Exchange365/Calendar")]
    public class CalendarSteps
    {
        public Event? CreateCalendarEvent(CalendarEvent calendarEvent, string userIdentifier, string? calendarId)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}";
            url = (!string.IsNullOrEmpty(calendarId)) ? $"{url}/calendars/{calendarId}/events"
                    : $"{url}/calendar/events";
            
            JsonContent content = JsonContent.Create(calendarEvent);
            
            return JsonConvert.DeserializeObject<Event>(GraphRest.Post(url, content));
        }

        public string DeleteCalendarEvent(string userIdentifier, string eventId, string? calendarId, string? calendarGroupId)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}";
            url = (!string.IsNullOrEmpty(calendarId) && !string.IsNullOrEmpty(calendarGroupId))
                ? $"{url}/calendarGroups/{calendarGroupId}/calendars/{calendarId}/events/{eventId}"
                : (!string.IsNullOrEmpty(calendarId)) ? $"{url}/calendars/{calendarId}/events/{eventId}"
                    : $"{url}/events/{eventId}";

            return GraphRest.Delete(url).StatusCode.ToString();
        }
        
        public EventList ListCalendarEvents(string userIdentifier, string? calendarId, string? calendarGroupId)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}";
            if (!string.IsNullOrEmpty(calendarId))
            {
                url = (!string.IsNullOrEmpty(calendarGroupId)) ? $"{url}/calendarGroups/{calendarGroupId}/calendars/{calendarId}"
                    : $"{url}/calendars/{calendarId}";
            }
            url += "/events";

            string result = GraphRest.Get(url);
            return JsonConvert.DeserializeObject<EventList>(result) ?? new EventList();
        }

        public Event? UpdateCalendarEvent(string userIdentifier, string eventId, string? subject, EventBody? body, DateTimeZone? startTime,
            DateTimeZone? endTime, Location? location, Attendee[]? attendees, bool? allowNewTimeProposals, int? reminderMinutesBeforeStart,
            bool? isOnlineMeeting, string onlineMeetingProvider, bool? isAllDay, bool? isReminderOn, bool? hideAttendees,
            string[]? categories, Sensitivity? sensitivity, Importance? importance, AttendeeStatus? showAs, bool? responseRequested)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/calendar/events/{eventId}";

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

            return JsonConvert.DeserializeObject<Event>(GraphRest.Patch(url, content));
        }
        
        public CalendarList ListCalendars(string userIdentifier)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/calendars";

            string result = GraphRest.Get(url);
            return JsonConvert.DeserializeObject<CalendarList>(result) ?? new CalendarList();
        }
    }
}