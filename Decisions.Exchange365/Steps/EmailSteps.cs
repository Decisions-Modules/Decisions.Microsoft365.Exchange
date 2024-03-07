using System.Net.Http.Json;
using Decisions.Exchange365.Data;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;

namespace Decisions.Exchange365.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Exchange365/Email")]
    public class EmailSteps
    {
        public void SearchForEmail(string userEmail)
        {
            try
            {
                GraphRest.Get(GetEmailUrl(userEmail));
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
        
        public void ListEmails(string userEmail)
        {
            try
            {
                GraphRest.Get(GetEmailUrl(userEmail));
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
        
        public void ForwardEmail(string userEmail)
        {
            try
            {
                JsonContent content = JsonContent.Create("messageBody");

                GraphRest.Post(GetEmailUrl(userEmail), content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
        
        public void ListUnreadEmails(string userEmail)
        {
            try
            {
                JsonContent content = JsonContent.Create("messageBody");

                GraphRest.Post(GetEmailUrl(userEmail), content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
        
        public void MarkEmailAsRead(string userEmail)
        {
            try
            {
                JsonContent content = JsonContent.Create("messageBody");

                GraphRest.Post(GetEmailUrl(userEmail), content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
        
        public void SendEmail(string userEmail)
        {
            try
            {
                JsonContent content = JsonContent.Create("messageBody");

                GraphRest.Post(GetEmailUrl(userEmail), content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
        
        public void SendReply(string userEmail)
        {
            try
            {
                JsonContent content = JsonContent.Create("messageBody");

                GraphRest.Post(GetEmailUrl(userEmail), content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }

        private string GetEmailUrl(string userEmail)
        {
            return $"{Exchange365Constants.GRAPH_URL}/users/{userEmail}/messages";
        }
    }
}