using System.Net.Http.Json;
using System.Text;
using Decisions.Microsoft365.Common;
using Decisions.Microsoft365.Common.API.Calendar;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Microsoft365/Exchange/Calendar")]
    public class CalendarSteps
    {
        private static readonly JsonSerializerSettings IgnoreNullValues = new()
        {
            NullValueHandling = NullValueHandling.Ignore
        };
        
        public Microsoft365Event? CreateCalendarEvent(Microsoft365CalendarEvent microsoft365CalendarEvent, string userIdentifier, string? calendarId,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = Microsoft365UrlHelper.GetCalendarEventUrl(userIdentifier, null, calendarId, null);
            
            JsonContent content = JsonContent.Create(microsoft365CalendarEvent);
            string result = GraphRest.Post(settingsOverride, urlExtension, content);
            
            return JsonHelper<Microsoft365Event?>.JsonDeserialize(result);
        }

        public string DeleteCalendarEvent(string userIdentifier, string eventId, string? calendarId, string? calendarGroupId,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = Microsoft365UrlHelper.GetCalendarEventUrl(userIdentifier, eventId, calendarId, calendarGroupId);

            HttpResponseMessage response = GraphRest.Delete(settingsOverride, urlExtension);

            return response.StatusCode.ToString();
        }
        
        public Microsoft365EventList? ListCalendarEvents(string userIdentifier, string? calendarId, string? calendarGroupId,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = Microsoft365UrlHelper.GetCalendarEventUrl(userIdentifier, null, calendarId, calendarGroupId);
            string result = GraphRest.Get(settingsOverride, urlExtension);
            
            return JsonHelper<Microsoft365EventList?>.JsonDeserialize(result);
        }

        public Microsoft365Event? UpdateCalendarEvent(string userIdentifier, string eventId, string? calendarId, string?
            calendarGroupId, Microsoft365UpdateCalendarEvent calendarEventMicrosoft365Update,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = Microsoft365UrlHelper.GetCalendarEventUrl(userIdentifier, eventId, calendarId, calendarGroupId);
            
            HttpContent content = new StringContent(JsonConvert.SerializeObject(calendarEventMicrosoft365Update, IgnoreNullValues),
                Encoding.UTF8, "application/json");

            string result = GraphRest.Patch(settingsOverride, urlExtension, content);
            
            return JsonHelper<Microsoft365Event?>.JsonDeserialize(result);
        }
        
        public Microsoft365CalendarList? ListCalendars(string userIdentifier,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetUserUrl(userIdentifier)}/calendars";
            string result = GraphRest.Get(settingsOverride, urlExtension);
            
            return JsonHelper<Microsoft365CalendarList?>.JsonDeserialize(result);
        }
    }
}