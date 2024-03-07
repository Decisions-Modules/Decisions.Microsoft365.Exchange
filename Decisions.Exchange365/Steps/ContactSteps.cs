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
        public void CreateContact()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
        }

        public void DeleteContact()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
        }

        public void ResolveContact()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
        }

        public Contact[] SearchContacts(string userEmail)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userEmail}/contacts";
            
            try
            {
                Task<string> result = GraphRest.Get(url);
                Contact[] response = JsonConvert.DeserializeObject<Contact[]>(result.Result) ?? Array.Empty<Contact>();
                return response;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }

        public void SearchGlobalContacts()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
        }
    }
}