using System.Net.Http.Json;
using System.Text;
using Decisions.Microsoft365.Exchange.API;
using DecisionsFramework.Design.Flow;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Microsoft365/Exchange/Calendar")]
    public class CalendarSteps
    {
        private static JsonSerializerSettings IgnoreNullValues = new()
        {
            NullValueHandling = NullValueHandling.Ignore
        };
        
        public MicrosoftEvent? CreateCalendarEvent(ExchangeCalendarEvent exchangeCalendarEvent, string userIdentifier, string? calendarId)
        {
            string urlExtension = $"/users/{userIdentifier}";
            urlExtension = (!string.IsNullOrEmpty(calendarId)) ? $"{urlExtension}/calendars/{calendarId}/events"
                    : $"{urlExtension}/calendar/events";
            
            JsonContent content = JsonContent.Create(exchangeCalendarEvent);
            
            return JsonConvert.DeserializeObject<MicrosoftEvent>(GraphRest.Post(urlExtension, content));
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
        
        public ExchangeEventList? ListCalendarEvents(string userIdentifier, string? calendarId, string? calendarGroupId)
        {
            string urlExtension = $"/users/{userIdentifier}";
            if (!string.IsNullOrEmpty(calendarId))
            {
                urlExtension = (!string.IsNullOrEmpty(calendarGroupId)) ? $"{urlExtension}/calendarGroups/{calendarGroupId}/calendars/{calendarId}"
                    : $"{urlExtension}/calendars/{calendarId}";
            }
            urlExtension += "/events";

            string result = GraphRest.Get(urlExtension);
            return ExchangeEventList.JsonDeserialize(result);
        }

        public MicrosoftEvent? UpdateCalendarEvent(string userIdentifier, string eventId, ExchangeUpdateCalendarEvent calendarEventExchangeUpdate)
        {
            string urlExtension = $"/users/{userIdentifier}/calendar/events/{eventId}";
            
            HttpContent content = new StringContent(JsonConvert.SerializeObject(calendarEventExchangeUpdate, IgnoreNullValues),
                Encoding.UTF8, "application/json");
            
            return JsonConvert.DeserializeObject<MicrosoftEvent>(GraphRest.Patch(urlExtension, content));
        }
        
        public ExchangeCalendarList? ListCalendars(string userIdentifier)
        {
            string urlExtension = $"/users/{userIdentifier}/calendars";

            string result = GraphRest.Get(urlExtension);
            return ExchangeCalendarList.JsonDeserialize(result);
        }
    }
}