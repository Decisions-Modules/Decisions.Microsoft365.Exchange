using System.Net;
using System.Net.Http.Json;
using Decisions.Exchange365.Data;
using DecisionsFramework.Design.Flow;
using Newtonsoft.Json;

namespace Decisions.Exchange365.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Exchange365/Email")]
    public class EmailSteps
    {
        public void SearchForEmail(string userIdentifier, string messageId)
        {
            GraphRest.Get($"{GetEmailUrl(userIdentifier)}/{messageId}");
        }
        
        public EmailDataList ListEmails(string userIdentifier)
        {
            string result = GraphRest.Get(GetEmailUrl(userIdentifier));
            EmailDataList? response = JsonConvert.DeserializeObject<EmailDataList>(result);
            
            return response;
        }
        
        public HttpStatusCode ForwardEmail(string userIdentifier, string emailContext)
        {
            JsonContent content = JsonContent.Create(emailContext);

            return GraphRest.HttpResponsePost(GetEmailUrl(userIdentifier), content).StatusCode;
        }
        
        public EmailDataList ListUnreadEmails(string userIdentifier)
        {
            string result = GraphRest.Get(GetEmailUrl(userIdentifier));
            EmailDataList? response = JsonConvert.DeserializeObject<EmailDataList>(result);
            
            return response;
        }
        
        public HttpStatusCode MarkEmailAsRead(string userIdentifier, string emailContext)
        {
            JsonContent content = JsonContent.Create(emailContext);
            
            return GraphRest.HttpResponsePost(GetEmailUrl(userIdentifier), content).StatusCode;
        }
        
        public HttpStatusCode SendEmail(string userIdentifier)
        {
            JsonContent content = JsonContent.Create("messageBody");

            return GraphRest.HttpResponsePost(GetEmailUrl(userIdentifier), content).StatusCode;
        }
        
        public HttpStatusCode SendReply(string userIdentifier, string emailContext)
        {
            JsonContent content = JsonContent.Create(emailContext);

            return GraphRest.HttpResponsePost(GetEmailUrl(userIdentifier), content).StatusCode;
        }
        
        public HttpStatusCode SendReplyToAll(string userIdentifier, string emailContext)
        {
            JsonContent content = JsonContent.Create(emailContext);

            return GraphRest.HttpResponsePost(GetEmailUrl(userIdentifier), content).StatusCode;
        }

        private string GetEmailUrl(string userIdentifier)
        {
            return $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/messages";
        }
    }
}