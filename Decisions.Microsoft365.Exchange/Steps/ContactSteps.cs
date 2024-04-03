using System.Net.Http.Json;
using Decisions.Microsoft365.Exchange.API;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Microsoft365/Exchange/Contacts")]
    public class ContactSteps
    {
        public string CreateContact(string userIdentifier, string? contactFolderId, MicrosoftContact contact)
        {
            string urlExtension = $"/users/{userIdentifier}";
            urlExtension = (!string.IsNullOrEmpty(contactFolderId)) ? $"{urlExtension}/contactFolders/{contactFolderId}/contacts" : $"{urlExtension}/contacts";
            
            JsonContent content = JsonContent.Create(contact);
            
            return GraphRest.HttpResponsePost(urlExtension, content).StatusCode.ToString();
        }

        public string DeleteContact(string userIdentifier, string? contactId)
        {
            string urlExtension = $"/users/{userIdentifier}/contacts/{contactId}";
            
            return GraphRest.Delete(urlExtension).StatusCode.ToString();
        }

        public Contact ResolveContact(string userIdentifier, string contactId,
            string? contactFolderId, string? childFolderId, string? expandQuery)
        {
            if (string.IsNullOrEmpty(expandQuery))
            {
                throw new BusinessRuleException("expandQuery cannot be empty.");
            }
            
            string urlExtension = $"/users/{userIdentifier}";
            urlExtension = (!string.IsNullOrEmpty(contactFolderId) && !string.IsNullOrEmpty(childFolderId))
                ? $"{urlExtension}/contactFolders/{contactFolderId}/childFolders/{childFolderId}"
                : (!string.IsNullOrEmpty(contactFolderId)) ? $"{urlExtension}/contactFolders/{contactFolderId}" : urlExtension;
            urlExtension = (!string.IsNullOrEmpty(expandQuery)) ? $"{urlExtension}/contacts/{contactId}?$expand={expandQuery}" : $"{urlExtension}/contacts/{contactId}";
            
            string result = GraphRest.Get(urlExtension);
            return JsonConvert.DeserializeObject<Contact>(result) ?? new Contact();
        }

        public ContactList ListContacts(string userIdentifier)
        {
            string urlExtension = $"/users/{userIdentifier}/contacts";
            
            string result = GraphRest.Get(urlExtension);
            return JsonConvert.DeserializeObject<ContactList>(result) ?? new ContactList();
        }
        
        public ContactList SearchContacts(string userIdentifier, string searchQuery)
        {
            if (string.IsNullOrEmpty(searchQuery))
            {
                throw new BusinessRuleException("searchQuery cannot be empty.");
            }
            
            string urlExtension = $"/users/{userIdentifier}/contacts?$search={searchQuery}";
            
            string result = GraphRest.Get(urlExtension);
            return JsonConvert.DeserializeObject<ContactList>(result) ?? new ContactList();
        }

        public PeopleList SearchGlobalContacts(string userIdentifier, string searchQuery)
        {
            if (string.IsNullOrEmpty(searchQuery))
            {
                throw new BusinessRuleException("searchQuery cannot be empty.");
            }
            
            string urlExtension = $"/users/{userIdentifier}/people?$search={searchQuery}";
            
            string result = GraphRest.Get(urlExtension);
            return JsonConvert.DeserializeObject<PeopleList>(result) ?? new PeopleList();
        }
    }
}