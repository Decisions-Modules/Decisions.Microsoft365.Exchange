using System.Net.Http.Json;
using Decisions.Microsoft365.Common;
using Decisions.Microsoft365.Common.API.People;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;

namespace Decisions.Microsoft365.Exchange.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Microsoft365/Exchange/Contacts")]
    public class ContactSteps
    {
        public string CreateContact(string userIdentifier, string? contactFolderId, Microsoft365ContactRequest contactRequest,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = Microsoft365UrlHelper.GetContactUrl(userIdentifier, null, contactFolderId, null);
            
            JsonContent content = JsonContent.Create(contactRequest);
            HttpResponseMessage response = GraphRest.HttpResponsePost(settingsOverride, urlExtension, content);
            
            return response.StatusCode.ToString();
        }

        public string DeleteContact(string userIdentifier, string? contactId,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = Microsoft365UrlHelper.GetContactUrl(userIdentifier, contactId, null, null);
            HttpResponseMessage response = GraphRest.Delete(settingsOverride, urlExtension);
            
            return response.StatusCode.ToString();
        }

        public Microsoft365Contact? GetContact(string userIdentifier, string contactId, string? contactFolderId, string? childFolderId, string? expandQuery,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = Microsoft365UrlHelper.GetContactUrl(userIdentifier, contactId, contactFolderId, childFolderId);

            if (!string.IsNullOrEmpty(expandQuery))
            {
                urlExtension = $"?$expand={expandQuery}";
            }

            string result = GraphRest.Get(settingsOverride, urlExtension);
            
            return JsonHelper<Microsoft365Contact?>.JsonDeserialize(result);
        }

        public Microsoft365ContactList? ListContacts(string userIdentifier,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = Microsoft365UrlHelper.GetContactUrl(userIdentifier, null, null, null);
            string result = GraphRest.Get(settingsOverride, urlExtension);
            
            return JsonHelper<Microsoft365ContactList?>.JsonDeserialize(result);
        }
        
        public Microsoft365ContactList? SearchContacts(string userIdentifier, string searchQuery,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            if (string.IsNullOrEmpty(searchQuery))
            {
                throw new BusinessRuleException("searchQuery cannot be empty.");
            }
            
            string urlExtension = $"{Microsoft365UrlHelper.GetContactUrl(userIdentifier, null, null, null)}?$search={searchQuery}";
            string result = GraphRest.Get(settingsOverride, urlExtension);
            
            return JsonHelper<Microsoft365ContactList?>.JsonDeserialize(result);
        }

        public Microsoft365PeopleList? SearchGlobalContacts(string userIdentifier, string searchQuery,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            if (string.IsNullOrEmpty(searchQuery))
            {
                throw new BusinessRuleException("searchQuery cannot be empty.");
            }
            
            string urlExtension = $"{Microsoft365UrlHelper.GetUserUrl(userIdentifier)}/people?$search={searchQuery}";
            string result = GraphRest.Get(settingsOverride, urlExtension);
            
            return JsonHelper<Microsoft365PeopleList?>.JsonDeserialize(result);
        }
    }
}