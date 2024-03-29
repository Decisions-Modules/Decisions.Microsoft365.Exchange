using System.Net.Http.Json;
using Decisions.Exchange365.API;
using Decisions.Exchange365.Data;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Exchange365.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Exchange365/Contacts")]
    public class ContactSteps
    {
        public string CreateContact(string userIdentifier, string? contactFolderId, MicrosoftContact contact)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}";
            url = (!string.IsNullOrEmpty(contactFolderId)) ? $"{url}/contactFolders/{contactFolderId}/contacts" : $"{url}/contacts";
            
            JsonContent content = JsonContent.Create(contact);
            
            return GraphRest.HttpResponsePost(url, content).StatusCode.ToString();
        }

        public string DeleteContact(string userIdentifier, string? contactId)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/contacts/{contactId}";
            
            return GraphRest.Delete(url).StatusCode.ToString();
        }

        public Contact ResolveContact(string userIdentifier, string contactId,
            string? contactFolderId, string? childFolderId, string? expandQuery)
        {
            if (string.IsNullOrEmpty(expandQuery))
            {
                throw new BusinessRuleException("expandQuery cannot be empty.");
            }
            
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}";
            url = (!string.IsNullOrEmpty(contactFolderId) && !string.IsNullOrEmpty(childFolderId))
                ? $"{url}/contactFolders/{contactFolderId}/childFolders/{childFolderId}"
                : (!string.IsNullOrEmpty(contactFolderId)) ? $"{url}/contactFolders/{contactFolderId}" : url;
            url = (!string.IsNullOrEmpty(expandQuery)) ? $"{url}/contacts/{contactId}?$expand={expandQuery}" : $"{url}/contacts/{contactId}";
            
            string result = GraphRest.Get(url);
            return JsonConvert.DeserializeObject<Contact>(result) ?? new Contact();
        }

        public ContactList ListContacts(string userIdentifier)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/contacts";
            
            string result = GraphRest.Get(url);
            return JsonConvert.DeserializeObject<ContactList>(result) ?? new ContactList();
        }
        
        public ContactList SearchContacts(string userIdentifier, string searchQuery)
        {
            if (string.IsNullOrEmpty(searchQuery))
            {
                throw new BusinessRuleException("searchQuery cannot be empty.");
            }
            
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/contacts?$search={searchQuery}";
            
            string result = GraphRest.Get(url);
            return JsonConvert.DeserializeObject<ContactList>(result) ?? new ContactList();
        }

        public PeopleList SearchGlobalContacts(string userIdentifier, string searchQuery)
        {
            if (string.IsNullOrEmpty(searchQuery))
            {
                throw new BusinessRuleException("searchQuery cannot be empty.");
            }
            
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/people?$search={searchQuery}";
            
            string result = GraphRest.Get(url);
            return JsonConvert.DeserializeObject<PeopleList>(result) ?? new PeopleList();
        }
    }
}