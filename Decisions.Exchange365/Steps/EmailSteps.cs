using System.Net;
using System.Net.Http.Json;
using Decisions.Exchange365.API;
using Decisions.Exchange365.Data;
using DecisionsFramework.Design.Flow;
using Newtonsoft.Json;

namespace Decisions.Exchange365.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Exchange365/Email")]
    public class EmailSteps
    {
        // TODO: test
        public void GetEmail(string userIdentifier, string messageId)
        {
            string url = $"{GetUrl(userIdentifier)}/messages/{messageId}";
            string result = GraphRest.Get(url);
            GraphRest.Get(url);
        }
        
        // TODO: configure to SEARCH for email
        private const string Url = $"{Exchange365Constants.GRAPH_URL}/users";
        public void SearchForEmail(string userIdentifier, string messageId)
        {
            string url = $"{GetUrl(userIdentifier)}/messages/{messageId}";
            string result = GraphRest.Get(url);
            GraphRest.Get(url);
        }
        
        // TODO: test
        public EmailList ListEmails(string userIdentifier)
        {
            string url = $"{GetUrl(userIdentifier)}/messages";
            string result = GraphRest.Get(url);
            EmailList? response = JsonConvert.DeserializeObject<EmailList>(result);
            
            return response;
        }
        
        // TODO: configure to FORWARD email
        public HttpStatusCode ForwardEmail(string userIdentifier, string emailContext)
        {
            JsonContent content = JsonContent.Create(emailContext);

            return GraphRest.HttpResponsePost(GetUrl(userIdentifier), content).StatusCode;
        }
        
        // TODO: configure to find UNREAD emails
        public EmailList ListUnreadEmails(string userIdentifier)
        {
            string url = $"{GetUrl(userIdentifier)}/messages";
            string result = GraphRest.Get(url);
            EmailList? response = JsonConvert.DeserializeObject<EmailList>(result);
            
            return response;
        }
        
        /* TODO: Find this in reference */
        public HttpStatusCode MarkEmailAsRead(string userIdentifier)
        {
            string url = $"{GetUrl(userIdentifier)}/???";
            JsonContent content = JsonContent.Create("???");
            
            return GraphRest.HttpResponsePost(url, content).StatusCode;
        }
        
        // TODO: test
        public HttpStatusCode SendEmail(string userIdentifier, EmailRequest emailMessage)
        {
            string url = $"{GetUrl(userIdentifier)}/sendMail";
            JsonContent content = JsonContent.Create(emailMessage);

            return GraphRest.HttpResponsePost(url, content).StatusCode;
        }
        
        // TODO: test
        public HttpStatusCode SendReply(string userIdentifier, string? mailFolderId, string messageId, EmailReplyRequest replyMessage)
        {
            string url = (!string.IsNullOrEmpty(mailFolderId))
                ? $"{GetUrl(userIdentifier)}/mailFolders/{mailFolderId}/messages/{messageId}/reply"
                : $"{GetUrl(userIdentifier)}/messages/{messageId}/reply";
            
            JsonContent content = JsonContent.Create(replyMessage);

            return GraphRest.HttpResponsePost(url, content).StatusCode;
        }
        
        // TODO: test
        public HttpStatusCode SendReplyToAll(string userIdentifier, string? mailFolderId, string messageId, EmailReplyRequest replyMessage)
        {
            string url = (!string.IsNullOrEmpty(mailFolderId))
                ? $"{GetUrl(userIdentifier)}/mailFolders/{mailFolderId}/messages/{messageId}/replyAll"
                : $"{GetUrl(userIdentifier)}/messages/{messageId}/replyAll";
            JsonContent content = JsonContent.Create(replyMessage);

            return GraphRest.HttpResponsePost(url, content).StatusCode;
        }

        private string GetUrl(string userIdentifier)
        {
            return $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}";
        }
    }
}