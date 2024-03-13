using System.Net;
using System.Net.Http.Json;
using Decisions.Exchange365.API;
using Decisions.Exchange365.Data;
using DecisionsFramework.Design.Flow;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Exchange365.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Exchange365/Contacts")]
    public class ContactSteps
    {
        public string CreateContact(string userIdentifier, string? contactFolderId, Contact contact)
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

        public void ResolveContact()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
        }

        public ContactList SearchContacts(string userIdentifier)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/contacts";
            
            string result = GraphRest.Get(url);
            return JsonConvert.DeserializeObject<ContactList>(result) ?? new ContactList();
        }

        public void SearchGlobalContacts()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
        }
    }
}