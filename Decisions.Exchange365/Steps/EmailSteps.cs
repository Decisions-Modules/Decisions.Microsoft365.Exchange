using System.Net.Http.Json;
using Decisions.Exchange365.Data;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using Microsoft.Graph.Me.SendMail;

namespace Decisions.Exchange365.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Exchange365/Email")]
    public class EmailSteps
    {
        public void SendEmail(string userEmail, SendMailPostRequestBody messageBody)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userEmail}/messages";
            
            try
            {
                JsonContent content = JsonContent.Create(messageBody);

                GraphREST.Post(url, content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
    }
}