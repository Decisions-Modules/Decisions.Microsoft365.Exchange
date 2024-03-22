using System.Net.Http.Json;
using Decisions.Exchange365.API;
using Decisions.Exchange365.Data;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Exchange365.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Exchange365/Calendar")]
    public class CalendarSteps
    {
        public string[] EventClassFields
        {
            get
            {
                return new Event().GetType().GetFields().Select(field => field.Name).ToArray();
            }
        }
        
        public string CreateCalendarEvent(CalendarEvent calendarEvent, string userIdentifier, string? calendarId)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}";
            url = (!string.IsNullOrEmpty(calendarId)) ? $"{url}/calendars/{calendarId}/events"
                    : $"{url}/events";
            
            JsonContent content = JsonContent.Create(calendarEvent);
            
            return GraphRest.HttpResponsePost(url, content).StatusCode.ToString();
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

        public EventList SearchCalendarEvents(string userIdentifier, string? calendarId, string? calendarGroupId, string searchString)
        {
            if (string.IsNullOrEmpty(searchString))
            {
                throw new BusinessRuleException("Search String cannot be empty.");
            }
            
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}?$search={searchString}";
            if (!string.IsNullOrEmpty(calendarId))
            {
                url = (!string.IsNullOrEmpty(calendarGroupId)) ? $"{url}/calendarGroups/{calendarGroupId}/calendars/{calendarId}"
                    : $"{url}/calendars/{calendarId}";
            }
            url += "/events";

            string result = GraphRest.Get(url);
            return JsonConvert.DeserializeObject<EventList>(result) ?? new EventList();
        }

        /* TODO: test */
        public string UpdateCalendarEvent(string userIdentifier, string eventId,
            [CheckboxListEditor(nameof(EventClassFields))] CalendarEvent calendarEventUpdate)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/events/{eventId}";
            
            string content = JsonConvert.SerializeObject(calendarEventUpdate, Formatting.Indented, new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore
            });
            
            return GraphRest.Patch(url, content).StatusCode.ToString();
        }
        
        public CalendarList ListCalendars(string userIdentifier)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/calendars";

            string result = GraphRest.Get(url);
            return JsonConvert.DeserializeObject<CalendarList>(result) ?? new CalendarList();
        }
    }
}