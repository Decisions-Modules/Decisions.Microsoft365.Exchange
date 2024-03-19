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
        
        public string CreateCalendarEvent(CalendarEvent calendarEvent)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
            
            JsonContent content = JsonContent.Create(calendarEvent);
            
            return GraphRest.HttpResponsePost(url, content).StatusCode.ToString();
        }

        public string DeleteCalendarEvent(string userIdentifier, string eventId)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/events/{eventId}";

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

        /* TODO: Rework input data. Check next comment for details. */
        public void UpdateCalendarEvent(string userIdentifier, string eventId,
            [CheckboxListEditor(nameof(EventClassFields))] string[] eventDetails)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/events/{eventId}";
            
            /* TODO: Utilize UpdateODataEntityStep features to dynamically build request data */
        }
        
        public CalendarList ListCalendars(string userIdentifier)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/calendars";

            string result = GraphRest.Get(url);
            return JsonConvert.DeserializeObject<CalendarList>(result) ?? new CalendarList();
        }
    }
}