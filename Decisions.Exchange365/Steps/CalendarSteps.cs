using Decisions.Exchange365.Data;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.Utilities;
using Microsoft.Graph.Models;

namespace Decisions.Exchange365.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Exchange365/Calendar")]
    public class CalendarSteps
    {
        public void CreateCalendarEvent(string subject, string body, DateTime startTime, DateTime endTime, string? timeZone, string location, DecisionsFramework.Design.Flow.CoreSteps.EMail.EmailAddress[] attendees, bool allowNewTimeProposals)
        {
            List<Attendee> eventAttendees = new List<Attendee>();
            foreach (var attendee in attendees)
            {
                eventAttendees.Add(new Attendee
                {
                    EmailAddress = new EmailAddress()
                    {
                        Address = attendee.Address,
                        Name = attendee.DisplayName
                    }
                });
            }
            
            Event requestBody = new Event
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = body,
                },
                Start = new DateTimeTimeZone
                {
                    DateTime = startTime.ToString(),
                    TimeZone = timeZone,
                },
                End = new DateTimeTimeZone
                {
                    DateTime = endTime.ToString(),
                    TimeZone = timeZone,
                },
                Location = new Location
                {
                    DisplayName = location,
                },
                Attendees = eventAttendees,
                AllowNewTimeProposals = allowNewTimeProposals,
                TransactionId = IDUtility.GetNewIdString()
            };
            
            try
            {
                Task<Event?> response = Exchange365Auth.GraphClient.Me.Events.PostAsync(requestBody, (requestConfiguration) =>
                {
                    if (!string.IsNullOrEmpty(timeZone))
                    {
                        requestConfiguration.Headers.Add("Prefer", $"outlook.timezone=\"{timeZone}\"");
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