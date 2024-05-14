using System.Net.Http.Json;
using Decisions.Microsoft365.Common;
using Decisions.Microsoft365.Common.API.Email;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;

namespace Decisions.Microsoft365.Exchange.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Microsoft365/Exchange/Email")]
    public class EmailSteps
    {
        public Microsoft365Message? GetEmail(string userIdentifier, string messageId,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetUserUrl(userIdentifier)}/messages/{messageId}";
            string result = GraphRest.Get(settingsOverride, urlExtension);
            
            return JsonHelper<Microsoft365Message?>.JsonDeserialize(result);
        }
        
        public Microsoft365Message?[] SearchEmails(string userIdentifier, string searchQuery, int? maxPageCount,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            if (string.IsNullOrEmpty(searchQuery))
            {
                throw new BusinessRuleException("searchQuery cannot be empty.");
            }
            
            int pageCount = (int)((maxPageCount > 0) ? maxPageCount : 1);
            string urlExtension = $"{Microsoft365UrlHelper.GetUserUrl(userIdentifier)}/messages?$search={searchQuery}";
            string result = GraphRest.Get(settingsOverride, urlExtension);

            List<Microsoft365EmailList?> emailLists = new List<Microsoft365EmailList?>();
            emailLists?.Add(JsonHelper<Microsoft365EmailList?>.JsonDeserialize(result));
            
            Microsoft365EmailList? tempEmailList = emailLists.First();
            for (int i = 0; i <= pageCount - 1 && !string.IsNullOrEmpty(tempEmailList.OdataNextLink); i++)
            {
                tempEmailList = ODataHelper<Microsoft365EmailList?>.GetNextPage(settingsOverride, tempEmailList.OdataNextLink);
                emailLists.Add(tempEmailList);
            }

            List<Microsoft365Message>? messages = new List<Microsoft365Message>();
            foreach (Microsoft365EmailList? emailList in emailLists)
            {
                foreach (Microsoft365Message email in emailList?.Value!)
                {
                    messages.Add(email);
                }
            }

            return messages.ToArray();
        }
        
        public Microsoft365Message?[] ListEmails(string userIdentifier, int? maxPageCount,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            int pageCount = (int)((maxPageCount > 0) ? maxPageCount : 1);
            string urlExtension = $"{Microsoft365UrlHelper.GetUserUrl(userIdentifier)}/messages";
            string result = GraphRest.Get(settingsOverride, urlExtension);

            List<Microsoft365EmailList?> emailLists = new List<Microsoft365EmailList?>();
            emailLists?.Add(JsonHelper<Microsoft365EmailList?>.JsonDeserialize(result));
            
            Microsoft365EmailList? tempEmailList = emailLists.First();
            for (int i = 0; i <= pageCount - 1 && !string.IsNullOrEmpty(tempEmailList.OdataNextLink); i++)
            {
                tempEmailList = ODataHelper<Microsoft365EmailList?>.GetNextPage(settingsOverride, tempEmailList.OdataNextLink);
                emailLists.Add(tempEmailList);
            }
            
            List<Microsoft365Message>? messages = new List<Microsoft365Message>();
            foreach (Microsoft365EmailList? emailList in emailLists)
            {
                foreach (Microsoft365Message email in emailList?.Value!)
                {
                    messages.Add(email);
                }
            }

            return messages.ToArray();
        }
        
        public Microsoft365Message?[] ListUnreadEmails(string userIdentifier, int? maxPageCount,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            int pageCount = (int)((maxPageCount > 0) ? maxPageCount : 1);
            string urlExtension = $"{Microsoft365UrlHelper.GetUserUrl(userIdentifier)}/messages";
            string result = GraphRest.Get(settingsOverride, urlExtension);
            
            List<Microsoft365EmailList?> emailLists = new List<Microsoft365EmailList?>();
            emailLists?.Add(JsonHelper<Microsoft365EmailList?>.JsonDeserialize(result));
            
            Microsoft365EmailList? tempEmailList = emailLists.First();
            for (int i = 0; i <= pageCount - 1 && !string.IsNullOrEmpty(tempEmailList.OdataNextLink); i++)
            {
                tempEmailList = ODataHelper<Microsoft365EmailList?>.GetNextPage(settingsOverride, tempEmailList.OdataNextLink);
                emailLists.Add(tempEmailList);
            }

            List<Microsoft365Message>? messages = new List<Microsoft365Message>();
            foreach (Microsoft365EmailList? emailList in emailLists)
            {
                foreach (Microsoft365Message email in emailList?.Value!)
                {
                    if (email.IsRead is false or null)
                    {
                        messages.Add(email);
                    }
                }
            }
            
            return messages.ToArray();
        }
        
        public string MarkEmailAsRead(string userIdentifier, string messageId,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetUserUrl(userIdentifier)}/messages/{messageId}";
            
            JsonContent content = JsonContent.Create(new Microsoft365EmailIsReadRequest{IsRead = true});
            HttpResponseMessage response = GraphRest.HttpResponsePatch(settingsOverride, urlExtension, content);

            return response.StatusCode.ToString();
        }
        
        public string SendEmail(string userIdentifier, string[] to, string[]? cc, string subject, string? body,
            Microsoft365BodyType? contentType, bool saveToSentItems,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
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
            HttpResponseMessage response = GraphRest.HttpResponsePost(settingsOverride, urlExtension, content);

            return response.StatusCode.ToString();
        }
        
        public string SendReply(string userIdentifier, string? mailFolderId, string messageId, string[] to, string[]? cc,
            string subject, string? body, Microsoft365BodyType? contentType, bool saveToSentItems,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
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
            HttpResponseMessage response = GraphRest.HttpResponsePost(settingsOverride, urlExtension, content);

            return response.StatusCode.ToString();
        }
        
        public string SendReplyToAll(string userIdentifier, string? mailFolderId, string messageId, string? comment,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetEmailUrl(userIdentifier, messageId, mailFolderId)}/replyAll";
            
            JsonContent content = JsonContent.Create(new Microsoft365EmailComment{Comment = comment});
            HttpResponseMessage response = GraphRest.HttpResponsePost(settingsOverride, urlExtension, content);

            return response.StatusCode.ToString();
        }
        
        public string ForwardEmail(string userIdentifier, string messageId, string? mailFolderId, string[] to, string comment,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetEmailUrl(userIdentifier, messageId, mailFolderId)}/forward";

            Microsoft365Recipient[] recipients = GetRecipients(to)!;
            Microsoft365ForwardRequest microsoft365ForwardRequest = new()
            {
                Comment = comment,
                ToRecipients = recipients
            };
            
            JsonContent content = JsonContent.Create(microsoft365ForwardRequest);
            HttpResponseMessage response = GraphRest.HttpResponsePost(settingsOverride, urlExtension, content);

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