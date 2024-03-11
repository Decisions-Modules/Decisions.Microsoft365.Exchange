using System.Net;
using System.Net.Http.Json;
using System.Text;
using Decisions.Exchange365.Data;
using Decisions.Utilities.OpenXmlPowerTools;
using DecisionsFramework.Design.DataStructure;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using FieldInfo = System.Reflection.FieldInfo;

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
        
        public HttpStatusCode CreateCalendarEvent(Event calendarEvent)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
            
            JsonContent content = JsonContent.Create(calendarEvent);
            
            return GraphRest.HttpResponsePost(url, content).StatusCode;
        }

        public HttpStatusCode DeleteCalendarEvent(string userIdentifier, string eventId)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/events/{eventId}";

            return GraphRest.Delete(url).StatusCode;
        }

        public Event[] SearchForCalendarEvent(string userIdentifier, string? calendarId)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}";
            url = (!string.IsNullOrEmpty(calendarId)) ? $"{url}/calendars/{calendarId}/events" : $"{url}/events";
            
            string result = GraphRest.Get(url);
            Event[] response = JsonConvert.DeserializeObject<Event[]>(result) ?? Array.Empty<Event>();
            return response;
        }

        /* TODO: Rework input data. Check next comment for details. */
        public void UpdateCalendarEvent(
            [PropertyClassification(0, "User Identifier", "Required")] string userIdentifier,
            [PropertyClassification(1, "Event ID", "Required")] string eventId,
            [CheckboxListEditor(nameof(EventClassFields))] string[] eventDetails)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/events/{eventId}";
            
            /* TODO: Utilize UpdateODataEntityStep features to dynamically build request data */
        }
        
        public Calendar[] ListCalendars(string userIdentifier)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/calendars";

            string result = GraphRest.Get(url);
            Calendar[] response = JsonConvert.DeserializeObject<Calendar[]>(result) ?? Array.Empty<Calendar>();
            return response;
        }
    }
}