using System.Net.Http.Json;
using Decisions.Microsoft365.Exchange.API;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Microsoft365/Exchange/Email")]
    public class EmailSteps
    {
        // TODO: create new Message class called "ExchangeMessage"
        public Message? GetEmail(string userIdentifier, string messageId)
        {
            string urlExtension = $"{GetUrlExtension(userIdentifier)}/messages/{messageId}";
            string result = GraphRest.Get(urlExtension);
            
            // TODO: return ExchangeMessageResponse.JsonDeserialize(result);
            return JsonConvert.DeserializeObject<Message>(result);
        }
        
        public ExchangeEmailList? SearchEmails(string userIdentifier, string searchQuery)
        {
            if (string.IsNullOrEmpty(searchQuery))
            {
                throw new BusinessRuleException("searchQuery cannot be empty.");
            }
            
            string urlExtension = $"{GetUrlExtension(userIdentifier)}/messages?$search={searchQuery}";
            string result = GraphRest.Get(urlExtension);

            return ExchangeEmailList.JsonDeserialize(result);
        }
        
        public ExchangeEmailList? ListEmails(string userIdentifier)
        {
            string urlExtension = $"{GetUrlExtension(userIdentifier)}/messages";
            string result = GraphRest.Get(urlExtension);

            return ExchangeEmailList.JsonDeserialize(result);
        }
        
        public ExchangeEmailList ListUnreadEmails(string userIdentifier)
        {
            string urlExtension = $"{GetUrlExtension(userIdentifier)}/messages";
            string result = GraphRest.Get(urlExtension);
            ExchangeEmailList? response = ExchangeEmailList.JsonDeserialize(result);

            // TODO: create new Message class called "ExchangeMessage"
            List<Message>? messages = new List<Message>();
            foreach (Message email in response.Value)
            {
                if (email.IsRead is false or null)
                {
                    messages.Add(email);
                }
            }

            ExchangeEmailList? unreadEmails = new ExchangeEmailList
            {
                OdataContext = response.OdataContext,
                Value = messages.ToArray()
            };
            
            return unreadEmails;
        }
        
        public string MarkEmailAsRead(string userIdentifier, string messageId)
        {
            string urlExtension = $"{GetUrlExtension(userIdentifier)}/messages/{messageId}";
            JsonContent content = JsonContent.Create(new ExchangeEmailIsReadRequest{IsRead = true});

            return GraphRest.HttpResponsePatch(urlExtension, content).StatusCode.ToString();
        }
        
        public string SendEmail(string userIdentifier, string[] to, string[]? cc, string subject, string? body,
            BodyType? contentType, bool saveToSentItems)
        {
            string urlExtension = $"{GetUrlExtension(userIdentifier)}/sendMail";
            
            ExchangeRecipient[] recipients = GetRecipients(to) ?? Array.Empty<ExchangeRecipient>();
            ExchangeRecipient[]? ccRecipients = (cc != null) ? GetRecipients(cc) : Array.Empty<ExchangeRecipient>();

            ExchangeSendEmailRequest emailMessage = new()
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
            
            ExchangeRecipient[] recipients = GetRecipients(to) ?? Array.Empty<ExchangeRecipient>();
            ExchangeRecipient[]? ccRecipients = (cc != null) ? GetRecipients(cc) : Array.Empty<ExchangeRecipient>();

            ExchangeSendEmailRequest emailMessage = new()
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
            
            JsonContent content = JsonContent.Create(new ExchangeEmailComment{Comment = comment});

            return GraphRest.HttpResponsePost(urlExtension, content).StatusCode.ToString();
        }
        
        public string ForwardEmail(string userIdentifier, string? mailFolderId, string messageId, string[] to, string comment)
        {
            string urlExtension = (!string.IsNullOrEmpty(mailFolderId))
                ? $"{GetUrlExtension(userIdentifier)}/mailFolders/{mailFolderId}/messages/{messageId}/forward"
                : $"{GetUrlExtension(userIdentifier)}/messages/{messageId}/forward";

            ExchangeRecipient[] recipients = GetRecipients(to);

            ExchangeForwardRequest exchangeForwardRequest = new()
            {
                Comment = comment,
                ToRecipients = recipients
            };
            
            JsonContent content = JsonContent.Create(exchangeForwardRequest);

            return GraphRest.HttpResponsePost(urlExtension, content).StatusCode.ToString();
        }

        private string GetUrlExtension(string userIdentifier)
        {
            return $"/users/{userIdentifier}";
        }
        
        private ExchangeRecipient[]? GetRecipients(string[] emailAddresses)
        {
            List<ExchangeRecipient> recipients = new List<ExchangeRecipient>();
            if (emailAddresses.Length > 0)
            {
                foreach (string emailAddress in emailAddresses)
                {
                    ExchangeRecipient recipient = new()
                    {
                        EmailAddress = new ExchangeEmailAddress
                        {
                            Address = emailAddress
                        }
                    };
                    recipients.Add(recipient);
                }
                
                return recipients.ToArray();
            }

            recipients.Add(new ExchangeRecipient
            {
                EmailAddress = new ExchangeEmailAddress
                {
                    Address = String.Empty
                }
            });

            return recipients.ToArray();
        }
    }
}