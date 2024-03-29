using System.Net.Http.Json;
using Decisions.Exchange365.API;
using Decisions.Exchange365.Data;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using EmailAddress = Decisions.Exchange365.API.EmailAddress;
using Recipient = Decisions.Exchange365.API.Recipient;

namespace Decisions.Exchange365.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Exchange365/Email")]
    public class EmailSteps
    {
        public Message GetEmail(string userIdentifier, string messageId)
        {
            string url = $"{GetUrl(userIdentifier)}/messages/{messageId}";
            string result = GraphRest.Get(url);

            return JsonConvert.DeserializeObject<Message>(result) ?? new Message();
        }
        
        public EmailList SearchEmails(string userIdentifier, string searchQuery)
        {
            if (string.IsNullOrEmpty(searchQuery))
            {
                throw new BusinessRuleException("searchQuery cannot be empty.");
            }
            
            string url = $"{GetUrl(userIdentifier)}/messages?$search={searchQuery}";
            string result = GraphRest.Get(url);

            return JsonConvert.DeserializeObject<EmailList>(result) ?? new EmailList();
        }
        
        public EmailList ListEmails(string userIdentifier)
        {
            string url = $"{GetUrl(userIdentifier)}/messages";
            string result = GraphRest.Get(url);

            return JsonConvert.DeserializeObject<EmailList>(result) ?? new EmailList();
        }
        
        public EmailList ListUnreadEmails(string userIdentifier)
        {
            string url = $"{GetUrl(userIdentifier)}/messages";
            string result = GraphRest.Get(url);
            EmailList? response = JsonConvert.DeserializeObject<EmailList>(result);

            List<Message>? messages = new List<Message>();
            foreach (Message email in response.Value)
            {
                if (email.IsRead is false or null)
                {
                    messages.Add(email);
                }
            }

            EmailList? unreadEmails = new EmailList
            {
                OdataContext = response.OdataContext,
                Value = messages.ToArray()
            };
            
            return unreadEmails;
        }
        
        public string MarkEmailAsRead(string userIdentifier, string messageId)
        {
            string url = $"{GetUrl(userIdentifier)}/messages/{messageId}";
            JsonContent content = JsonContent.Create(new EmailIsReadRequest{IsRead = true});

            return GraphRest.HttpResponsePatch(url, content).StatusCode.ToString();
        }
        
        public string SendEmail(string userIdentifier, string[] to, string[]? cc, string subject, string? body,
            BodyType? contentType, bool saveToSentItems)
        {
            string url = $"{GetUrl(userIdentifier)}/sendMail";
            
            Recipient[] recipients = GetRecipients(to) ?? Array.Empty<Recipient>();
            Recipient[]? ccRecipients = (cc != null) ? GetRecipients(cc) : Array.Empty<Recipient>();

            SendEmailRequest emailMessage = new()
            {
                Message = new()
                {
                    Body = new Body
                    {
                        ContentType = contentType.ToString() ?? BodyType.Text.ToString(),
                        Content = body
                    },
                    Subject = subject,
                    ToRecipients = recipients,
                    CcRecipients = ccRecipients
                },
                SaveToSentItems = saveToSentItems
            };
            
            JsonContent content = JsonContent.Create(emailMessage);

            return GraphRest.HttpResponsePost(url, content).StatusCode.ToString();
        }
        
        public string SendReply(string userIdentifier, string? mailFolderId, string messageId,
            string[] to, string[]? cc, string subject, string? body,
            BodyType? contentType, bool saveToSentItems)
        {
            string url = (!string.IsNullOrEmpty(mailFolderId))
                ? $"{GetUrl(userIdentifier)}/mailFolders/{mailFolderId}/messages/{messageId}/reply"
                : $"{GetUrl(userIdentifier)}/messages/{messageId}/reply";
            
            Recipient[] recipients = GetRecipients(to) ?? Array.Empty<Recipient>();
            Recipient[]? ccRecipients = (cc != null) ? GetRecipients(cc) : Array.Empty<Recipient>();

            SendEmailRequest emailMessage = new()
            {
                Message = new()
                {
                    Body = new Body
                    {
                        ContentType = contentType.ToString() ?? BodyType.Text.ToString(),
                        Content = body
                    },
                    Subject = subject,
                    ToRecipients = recipients,
                    CcRecipients = ccRecipients
                },
                SaveToSentItems = saveToSentItems
            };
            
            JsonContent content = JsonContent.Create(emailMessage);

            return GraphRest.HttpResponsePost(url, content).StatusCode.ToString();
        }
        
        public string SendReplyToAll(string userIdentifier, string? mailFolderId, string messageId,
            string? comment)
        {
            string url = (!string.IsNullOrEmpty(mailFolderId))
                ? $"{GetUrl(userIdentifier)}/mailFolders/{mailFolderId}/messages/{messageId}/replyAll"
                : $"{GetUrl(userIdentifier)}/messages/{messageId}/replyAll";
            
            JsonContent content = JsonContent.Create(new EmailComment{Comment = comment});

            return GraphRest.HttpResponsePost(url, content).StatusCode.ToString();
        }
        
        public string ForwardEmail(string userIdentifier, string? mailFolderId, string messageId, string[] to, string comment)
        {
            string url = (!string.IsNullOrEmpty(mailFolderId))
                ? $"{GetUrl(userIdentifier)}/mailFolders/{mailFolderId}/messages/{messageId}/forward"
                : $"{GetUrl(userIdentifier)}/messages/{messageId}/forward";

            Recipient[] recipients = GetRecipients(to);

            ForwardRequest forwardRequest = new()
            {
                Comment = comment,
                ToRecipients = recipients
            };
            
            JsonContent content = JsonContent.Create(forwardRequest);

            return GraphRest.HttpResponsePost(url, content).StatusCode.ToString();
        }

        private string GetUrl(string userIdentifier)
        {
            return $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}";
        }
        
        private Recipient[]? GetRecipients(string[] emailAddresses)
        {
            List<Recipient> recipients = new List<Recipient>();
            if (emailAddresses.Length > 0)
            {
                foreach (string emailAddress in emailAddresses)
                {
                    Recipient recipient = new()
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = emailAddress
                        }
                    };
                    recipients.Add(recipient);
                }
                
                return recipients.ToArray();
            }

            recipients.Add(new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = String.Empty
                }
            });

            return recipients.ToArray();
        }
    }
}