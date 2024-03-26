using System.Net.Http.Json;
using Decisions.Exchange365.API;
using Decisions.Exchange365.Data;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using Request = Decisions.Exchange365.API.Request;

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

        // TODO: test
        public EventList SearchCalendarEvents(string userIdentifier, string query, int? numberOfResults)
        {
            if (string.IsNullOrEmpty(query))
            {
                throw new BusinessRuleException("query cannot be empty.");
            }
            
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/search/query";

            SearchRequests request = new SearchRequests
            {
                Requests = new []
                {
                    new Request
                    {
                        EntityTypes = new []{"event"},
                        Query = new Query
                        {
                            QueryString = $"{query}"
                        },
                        From = 0,
                        Size = numberOfResults ?? 100
                    }
                }
            };

            JsonContent content = JsonContent.Create(request);
{}
            string result = GraphRest.Post(url, content);
            
            return JsonConvert.DeserializeObject<EventList>(result) ?? new EventList();
        }

        /* TODO: test */
        public Event? UpdateCalendarEvent(string userIdentifier, string eventId, UpdateCalendarEvent calendarEventUpdate)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/calendar/events/{eventId}";
            
            string contentString = JsonConvert.SerializeObject(calendarEventUpdate, Formatting.None, new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore
            });
            
            JsonContent content = JsonContent.Create(contentString);

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