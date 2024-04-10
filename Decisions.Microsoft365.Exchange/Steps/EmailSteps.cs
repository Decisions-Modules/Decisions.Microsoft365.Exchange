using System.Net.Http.Json;
using Decisions.Microsoft365.Common;
using Decisions.Microsoft365.Common.API.Email;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;

namespace Decisions.Microsoft365.Exchange.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Microsoft365/Exchange/Email")]
    public class EmailSteps
    {
        public Microsoft365Message? GetEmail(string userIdentifier, string messageId)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetUserUrl(userIdentifier)}/messages/{messageId}";
            string result = GraphRest.Get(urlExtension);
            
            return JsonHelper<Microsoft365Message?>.JsonDeserialize(result);
        }
        
        public Microsoft365EmailList? SearchEmails(string userIdentifier, string searchQuery)
        {
            if (string.IsNullOrEmpty(searchQuery))
            {
                throw new BusinessRuleException("searchQuery cannot be empty.");
            }
            
            string urlExtension = $"{Microsoft365UrlHelper.GetUserUrl(userIdentifier)}/messages?$search={searchQuery}";
            string result = GraphRest.Get(urlExtension);

            return JsonHelper<Microsoft365EmailList?>.JsonDeserialize(result);
        }
        
        public Microsoft365EmailList? ListEmails(string userIdentifier)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetUserUrl(userIdentifier)}/messages";
            string result = GraphRest.Get(urlExtension);

            return JsonHelper<Microsoft365EmailList?>.JsonDeserialize(result);
        }
        
        public Microsoft365EmailList ListUnreadEmails(string userIdentifier)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetUserUrl(userIdentifier)}/messages";
            string result = GraphRest.Get(urlExtension);
            
            Microsoft365EmailList? response = JsonHelper<Microsoft365EmailList?>.JsonDeserialize(result);

            List<Microsoft365Message>? messages = new List<Microsoft365Message>();
            foreach (Microsoft365Message email in response.Value)
            {
                if (email.IsRead is false or null)
                {
                    messages.Add(email);
                }
            }

            Microsoft365EmailList? unreadEmails = new Microsoft365EmailList
            {
                OdataContext = response.OdataContext,
                Value = messages.ToArray()
            };
            
            return unreadEmails;
        }
        
        public string MarkEmailAsRead(string userIdentifier, string messageId)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetUserUrl(userIdentifier)}/messages/{messageId}";
            
            JsonContent content = JsonContent.Create(new Microsoft365EmailIsReadRequest{IsRead = true});
            HttpResponseMessage response = GraphRest.HttpResponsePatch(urlExtension, content);

            return response.StatusCode.ToString();
        }
        
        public string SendEmail(string userIdentifier, string[] to, string[]? cc, string subject, string? body,
            Microsoft365BodyType? contentType, bool saveToSentItems)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetUserUrl(userIdentifier)}/sendMail";
            
            Microsoft365Recipient[] recipients = GetRecipients(to) ?? Array.Empty<Microsoft365Recipient>();
            Microsoft365Recipient[]? ccRecipients = (cc != null) ? GetRecipients(cc) : Array.Empty<Microsoft365Recipient>();

            Microsoft365SendEmailRequest emailMessage = new()
            {
                Message = new()
                {
                    Body = new Microsoft365EmailBody
                    {
                        ContentType = contentType.ToString() ?? Microsoft365BodyType.Text.ToString(),
                        Content = body
                    },
                    Subject = subject,
                    ToRecipients = recipients,
                    CcRecipients = ccRecipients
                },
                SaveToSentItems = saveToSentItems
            };
            
            JsonContent content = JsonContent.Create(emailMessage);
            HttpResponseMessage response = GraphRest.HttpResponsePost(urlExtension, content);

            return response.StatusCode.ToString();
        }
        
        public string SendReply(string userIdentifier, string? mailFolderId, string messageId,
            string[] to, string[]? cc, string subject, string? body,
            Microsoft365BodyType? contentType, bool saveToSentItems)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetEmailUrl(userIdentifier, messageId, mailFolderId)}/reply";
            
            Microsoft365Recipient[] recipients = GetRecipients(to) ?? Array.Empty<Microsoft365Recipient>();
            Microsoft365Recipient[]? ccRecipients = (cc != null) ? GetRecipients(cc) : Array.Empty<Microsoft365Recipient>();

            Microsoft365SendEmailRequest emailMessage = new()
            {
                Message = new()
                {
                    Body = new Microsoft365EmailBody
                    {
                        ContentType = contentType.ToString() ?? Microsoft365BodyType.Text.ToString(),
                        Content = body
                    },
                    Subject = subject,
                    ToRecipients = recipients,
                    CcRecipients = ccRecipients
                },
                SaveToSentItems = saveToSentItems
            };
            
            JsonContent content = JsonContent.Create(emailMessage);
            HttpResponseMessage response = GraphRest.HttpResponsePost(urlExtension, content);

            return response.StatusCode.ToString();
        }
        
        public string SendReplyToAll(string userIdentifier, string? mailFolderId, string messageId,
            string? comment)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetEmailUrl(userIdentifier, messageId, mailFolderId)}/replyAll";
            
            JsonContent content = JsonContent.Create(new Microsoft365EmailComment{Comment = comment});
            HttpResponseMessage response = GraphRest.HttpResponsePost(urlExtension, content);

            return response.StatusCode.ToString();
        }
        
        public string ForwardEmail(string userIdentifier, string messageId, string? mailFolderId, string[] to, string comment)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetEmailUrl(userIdentifier, messageId, mailFolderId)}/forward";

            Microsoft365Recipient[] recipients = GetRecipients(to);
            Microsoft365ForwardRequest microsoft365ForwardRequest = new()
            {
                Comment = comment,
                ToRecipients = recipients
            };
            
            JsonContent content = JsonContent.Create(microsoft365ForwardRequest);
            HttpResponseMessage response = GraphRest.HttpResponsePost(urlExtension, content);

            return response.StatusCode.ToString();
        }
        
        private Microsoft365Recipient[]? GetRecipients(string[] emailAddresses)
        {
            List<Microsoft365Recipient> recipients = new List<Microsoft365Recipient>();
            if (emailAddresses.Length > 0)
            {
                foreach (string emailAddress in emailAddresses)
                {
                    Microsoft365Recipient recipient = new()
                    {
                        EmailAddress = new Microsoft365Address
                        {
                            Address = emailAddress
                        }
                    };
                    recipients.Add(recipient);
                }
                
                return recipients.ToArray();
            }

            recipients.Add(new Microsoft365Recipient
            {
                EmailAddress = new Microsoft365Address
                {
                    Address = string.Empty
                }
            });

            return recipients.ToArray();
        }
    }
}