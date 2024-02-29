using System;
using System.Threading.Tasks;
using Decisions.Exchange365.Data;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using Microsoft.Graph.Models;

namespace Decisions.Exchange365.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Exchange365/Calendar")]
    public class CalendarSteps
    {
        public Event? CreateCalendarEvent(Event calendarEvent)
        {
            try
            {
                return Exchange365Auth.GraphClient.Me.Events.PostAsync(calendarEvent, (requestConfiguration) =>
                {
                    if (!string.IsNullOrEmpty(calendarEvent.OriginalStartTimeZone))
                    {
                        requestConfiguration.Headers.Add("Prefer", $"outlook.timezone=\"{calendarEvent.OriginalStartTimeZone}\"");
                    }
                }).Result;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
    }
}