using Decisions.Exchange365.Data;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Flow.StepImplementations;
using Microsoft.Graph.Models;

namespace Decisions.Exchange365.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Exchange365/Calendar")]
    [ShapeImageAndColorProvider(null, Exchange365Constants.EXCHANGE365_IMAGES_PATH)]
    public class CalendarSteps
    {
        public Task<Event?> CreateCalendarEvent(Event eventDetails, string? preferredTimeZone)
        {
            try
            {
                return Exchange365Auth.GraphClient.Me.Events.PostAsync(eventDetails, (requestConfiguration) =>
                {
                    if (!string.IsNullOrEmpty(preferredTimeZone))
                    {
                        requestConfiguration.Headers.Add("Prefer", $"outlook.timezone=\"{preferredTimeZone}\"");
                    }
                });
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
        
        public Task<Event?> DeleteCalendarEvent(Event eventDetails, string? preferredTimeZone)
        {
            try
            {
                return Exchange365Auth.GraphClient.Me.Events.PostAsync(eventDetails, (requestConfiguration) =>
                {
                    if (!string.IsNullOrEmpty(preferredTimeZone))
                    {
                        requestConfiguration.Headers.Add("Prefer", $"outlook.timezone=\"{preferredTimeZone}\"");
                    }
                });
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
    }
}