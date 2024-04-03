using System.Net.Http.Json;
using Decisions.Microsoft365.Exchange.API;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using EmailAddress = Decisions.Microsoft365.Exchange.API.EmailAddress;
using Recipient = Decisions.Microsoft365.Exchange.API.Recipient;

namespace Decisions.Microsoft365.Exchange.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Microsoft365/Exchange/Email")]
    public class EmailSteps
    {
        public Message GetEmail(string userIdentifier, string messageId)
        {
            string urlExtension = $"{GetUrlExtension(userIdentifier)}/messages/{messageId}";
            string result = GraphRest.Get(urlExtension);

            return JsonConvert.DeserializeObject<Message>(result) ?? new Message();
        }
        
        public EmailList SearchEmails(string userIdentifier, string searchQuery)
        {
            if (string.IsNullOrEmpty(searchQuery))
            {
                throw new BusinessRuleException("searchQuery cannot be empty.");
            }
            
            string urlExtension = $"{GetUrlExtension(userIdentifier)}/messages?$search={searchQuery}";
            string result = GraphRest.Get(urlExtension);

            return JsonConvert.DeserializeObject<EmailList>(result) ?? new EmailList();
        }
        
        public EmailList ListEmails(string userIdentifier)
        {
            string urlExtension = $"{GetUrlExtension(userIdentifier)}/messages";
            string result = GraphRest.Get(urlExtension);

            return JsonConvert.DeserializeObject<EmailList>(result) ?? new EmailList();
        }
        
        public EmailList ListUnreadEmails(string userIdentifier)
        {
            string urlExtension = $"{GetUrlExtension(userIdentifier)}/messages";
            string result = GraphRest.Get(urlExtension);
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
            string urlExtension = $"{GetUrlExtension(userIdentifier)}/messages/{messageId}";
            JsonContent content = JsonContent.Create(new EmailIsReadRequest{IsRead = true});

            return GraphRest.HttpResponsePatch(urlExtension, content).StatusCode.ToString();
        }
        
        public string SendEmail(string userIdentifier, string[] to, string[]? cc, string subject, string? body,
            BodyType? contentType, bool saveToSentItems)
        {
            string urlExtension = $"{GetUrlExtension(userIdentifier)}/sendMail";
            
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

            return GraphRest.HttpResponsePost(urlExtension, content).StatusCode.ToString();
        }
        
        public string SendReply(string userIdentifier, string? mailFolderId, string messageId,
            string[] to, string[]? cc, string subject, string? body,
            BodyType? contentType, bool saveToSentItems)
        {
            string urlExtension = (!string.IsNullOrEmpty(mailFolderId))
                ? $"{GetUrlExtension(userIdentifier)}/mailFolders/{mailFolderId}/messages/{messageId}/reply"
                : $"{GetUrlExtension(userIdentifier)}/messages/{messageId}/reply";
            
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

            return GraphRest.HttpResponsePost(urlExtension, content).StatusCode.ToString();
        }
        
        public string SendReplyToAll(string userIdentifier, string? mailFolderId, string messageId,
            string? comment)
        {
            string urlExtension = (!string.IsNullOrEmpty(mailFolderId))
                ? $"{GetUrlExtension(userIdentifier)}/mailFolders/{mailFolderId}/messages/{messageId}/replyAll"
                : $"{GetUrlExtension(userIdentifier)}/messages/{messageId}/replyAll";
            
            JsonContent content = JsonContent.Create(new EmailComment{Comment = comment});

            return GraphRest.HttpResponsePost(urlExtension, content).StatusCode.ToString();
        }
        
        public string ForwardEmail(string userIdentifier, string? mailFolderId, string messageId, string[] to, string comment)
        {
            string urlExtension = (!string.IsNullOrEmpty(mailFolderId))
                ? $"{GetUrlExtension(userIdentifier)}/mailFolders/{mailFolderId}/messages/{messageId}/forward"
                : $"{GetUrlExtension(userIdentifier)}/messages/{messageId}/forward";

            Recipient[] recipients = GetRecipients(to);

            ForwardRequest forwardRequest = new()
            {
                Comment = comment,
                ToRecipients = recipients
            };
            
            JsonContent content = JsonContent.Create(forwardRequest);

            return GraphRest.HttpResponsePost(urlExtension, content).StatusCode.ToString();
        }

        private string GetUrlExtension(string userIdentifier)
        {
            return $"/users/{userIdentifier}";
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