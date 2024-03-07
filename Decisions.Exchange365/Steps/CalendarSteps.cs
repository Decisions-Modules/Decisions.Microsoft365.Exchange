using Decisions.Exchange365.Data;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Exchange365.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Exchange365/Calendar")]
    public class CalendarSteps
    {
        public Event CreateCalendarEvent(Event calendarEvent)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
            
            try
            {
                return null;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }

        public void DeleteCalendarEvent()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
        }

        public Event[] SearchForCalendarEvent()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
            
            try
            {
                Task<string> result = GraphRest.Get(url);
                Event[] response = JsonConvert.DeserializeObject<Event[]>(result.Result) ?? Array.Empty<Event>();
                return response;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }

        public void UpdateCalendarEvent()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
        }
        
        public Calendar[] ListCalendars(string userEmail)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userEmail}/calendars";

            try
            {
                Task<string> result = GraphRest.Get(url);
                Calendar[] response = JsonConvert.DeserializeObject<Calendar[]>(result.Result) ?? Array.Empty<Calendar>();
                return response;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
    }
}