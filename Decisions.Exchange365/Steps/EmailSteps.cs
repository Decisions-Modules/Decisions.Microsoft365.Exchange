using System.Net.Http.Json;
using Decisions.Exchange365.API;
using Decisions.Exchange365.Data;
using DecisionsFramework.Data.DataTypes;
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
        
        // TODO: configure to SEARCH for email
        private const string Url = $"{Exchange365Constants.GRAPH_URL}/users";
        public void SearchForEmail(string userIdentifier, string messageId)
        {
            string url = $"{GetUrl(userIdentifier)}/messages/{messageId}";
            string result = GraphRest.Get(url);
        }
        
        public EmailList ListEmails(string userIdentifier)
        {
            string url = $"{GetUrl(userIdentifier)}/messages";
            string result = GraphRest.Get(url);

            return JsonConvert.DeserializeObject<EmailList>(result) ?? new EmailList();
        }
        
        // TODO: test
        public string ForwardEmail(string userIdentifier, string? mailFolderId, string messageId, string[] to, string comment)
        {
            string url = (!string.IsNullOrEmpty(mailFolderId))
                ? $"{GetUrl(userIdentifier)}/mailFolders/{mailFolderId}/messages/{messageId}/reply"
                : $"{GetUrl(userIdentifier)}/messages/{messageId}/forward";

            Recipient[] recipients = GetRecipients(to);

            ForwardRequest forwardRequest = new()
            {
                Comment = comment,
                ToRecipients = recipients
            };
            
            JsonContent content = JsonContent.Create(comment);

            return GraphRest.HttpResponsePost(GetUrl(userIdentifier), content).StatusCode.ToString();
        }
        
        // TODO: configure to find UNREAD emails
        public EmailList ListUnreadEmails(string userIdentifier)
        {
            string url = $"{GetUrl(userIdentifier)}/messages";
            string result = GraphRest.Get(url);
            EmailList? response = JsonConvert.DeserializeObject<EmailList>(result);
            
            return response;
        }
        
        /* TODO: configure to mark email as READ */
        public string MarkEmailAsRead(string userIdentifier)
        {
            string url = $"{GetUrl(userIdentifier)}/???";
            JsonContent content = JsonContent.Create("???");
            
            return GraphRest.HttpResponsePost(url, content).StatusCode.ToString();
        }
        
        // TODO: test attachments
        public string SendEmail(string userIdentifier, string[] to, string[]? cc, string subject, string? body,
            BodyType? contentType, FileData[]? fileAttachments, bool saveToSentItems)
        {
            string url = $"{GetUrl(userIdentifier)}/sendMail";

            Recipient[] recipients = GetRecipients(to);
            Recipient[]? ccRecipients = GetRecipients(cc);
            Attachment[]? attachments = GetAttachments(fileAttachments);

            SendEmailRequest emailMessage = new()
            {
                Message = new()
                {
                    Attachments = attachments,
                    Body = new Body
                    {
                        ContentType = contentType.ToString() ?? BodyType.Text.ToString(),
                        Content = body
                    },
                    CcRecipients = ccRecipients,
                    Subject = subject,
                    ToRecipients = recipients
                },
                SaveToSentItems = saveToSentItems
            };
            
            JsonContent content = JsonContent.Create(emailMessage);

            return GraphRest.HttpResponsePost(url, content).StatusCode.ToString();
        }
        
        // TODO: test
        public string SendReply(string userIdentifier, string? mailFolderId, string messageId,
            string[] to, string[]? cc, string subject, string? body,
            BodyType? contentType, FileData[]? fileAttachments, bool saveToSentItems)
        {
            string url = (!string.IsNullOrEmpty(mailFolderId))
                ? $"{GetUrl(userIdentifier)}/mailFolders/{mailFolderId}/messages/{messageId}/reply"
                : $"{GetUrl(userIdentifier)}/messages/{messageId}/reply";
            
            Recipient[] recipients = GetRecipients(to);
            Recipient[]? ccRecipients = GetRecipients(cc);
            Attachment[]? attachments = GetAttachments(fileAttachments);

            SendEmailRequest emailMessage = new()
            {
                Message = new()
                {
                    Attachments = attachments,
                    Body = new Body
                    {
                        ContentType = contentType.ToString() ?? BodyType.Text.ToString(),
                        Content = body
                    },
                    CcRecipients = ccRecipients,
                    Subject = subject,
                    ToRecipients = recipients
                },
                SaveToSentItems = saveToSentItems
            };
            
            JsonContent content = JsonContent.Create(emailMessage);

            return GraphRest.HttpResponsePost(url, content).StatusCode.ToString();
        }
        
        // TODO: test
        public string SendReplyToAll(string userIdentifier, string? mailFolderId, string messageId,
            string[] to, string[]? cc, string subject, string? body,
            BodyType? contentType, FileData[]? fileAttachments, bool saveToSentItems)
        {
            string url = (!string.IsNullOrEmpty(mailFolderId))
                ? $"{GetUrl(userIdentifier)}/mailFolders/{mailFolderId}/messages/{messageId}/replyAll"
                : $"{GetUrl(userIdentifier)}/messages/{messageId}/replyAll";
            
            Recipient[] recipients = GetRecipients(to);
            Recipient[]? ccRecipients = GetRecipients(cc);
            Attachment[]? attachments = GetAttachments(fileAttachments);

            SendEmailRequest emailMessage = new()
            {
                Message = new()
                {
                    Attachments = attachments,
                    Body = new Body
                    {
                        ContentType = contentType.ToString() ?? BodyType.Text.ToString(),
                        Content = body
                    },
                    CcRecipients = ccRecipients,
                    Subject = subject,
                    ToRecipients = recipients
                },
                SaveToSentItems = saveToSentItems
            };
            
            JsonContent content = JsonContent.Create(emailMessage);

            return GraphRest.HttpResponsePost(url, content).StatusCode.ToString();
        }

        private string GetUrl(string userIdentifier)
        {
            return $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}";
        }
        
        private Recipient[] GetRecipients(string[] emailAddresses)
        {
            List<Recipient> recipients = new List<Recipient>();
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

        private Attachment[] GetAttachments(FileData[]? fileAttachments)
        {
            List<Attachment>? attachments = null;
            if (fileAttachments != null)
            {
                foreach (FileData file in fileAttachments)
                {
                    Attachment attachment = new Attachment
                    {
                        AdditionalData = new Dictionary<string, object>
                        {
                            { file.FileName, file.Contents }
                        },
                        Id = file.Id,
                        ContentType = file.FileType,
                        Name = file.FileName
                    };
                    attachments.Add(attachment);
                }
            }

            return attachments.ToArray();
        }
    }
}