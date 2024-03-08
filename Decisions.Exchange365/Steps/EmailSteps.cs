using System;
using System.Net;
using System.Net.Http.Json;
using Decisions.Exchange365.Data;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using Newtonsoft.Json;

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
        
        public EmailDataList ListEmails(string userEmail)
        {
            try
            {
                string result = GraphRest.Get(GetEmailUrl(userEmail));
                EmailDataList? response = JsonConvert.DeserializeObject<EmailDataList>(result);
                return response;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
        
        public HttpStatusCode ForwardEmail(string userEmail, string emailContext)
        {
            try
            {
                JsonContent content = JsonContent.Create(emailContext);

                return GraphRest.HttpResponsePost(GetEmailUrl(userEmail), content).StatusCode;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
        
        public EmailDataList ListUnreadEmails(string userEmail)
        {
            try
            {
                string result = GraphRest.Get(GetEmailUrl(userEmail));
                EmailDataList? response = JsonConvert.DeserializeObject<EmailDataList>(result);
                return response;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
        
        public HttpStatusCode MarkEmailAsRead(string userEmail, string emailContext)
        {
            try
            {
                JsonContent content = JsonContent.Create(emailContext);

                return GraphRest.HttpResponsePost(GetEmailUrl(userEmail), content).StatusCode;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
        
        public HttpStatusCode SendEmail(string userEmail)
        {
            try
            {
                JsonContent content = JsonContent.Create("messageBody");

                return GraphRest.HttpResponsePost(GetEmailUrl(userEmail), content).StatusCode;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
        
        public HttpStatusCode SendReply(string userEmail, string emailContext)
        {
            try
            {
                JsonContent content = JsonContent.Create(emailContext);

                return GraphRest.HttpResponsePost(GetEmailUrl(userEmail), content).StatusCode;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
        
        public HttpStatusCode SendReplyToAll(string userEmail, string emailContext)
        {
            try
            {
                JsonContent content = JsonContent.Create(emailContext);

                return GraphRest.HttpResponsePost(GetEmailUrl(userEmail), content).StatusCode;
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