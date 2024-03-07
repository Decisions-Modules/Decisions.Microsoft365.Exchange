using Decisions.Exchange365.Data;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Exchange365.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Exchange365/Groups")]
    public class GroupSteps
    {
        public Group[] ListGroups()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/groups";
            
            try
            {
                Task<string> result = GraphRest.Get(url);
                Group[] response = JsonConvert.DeserializeObject<Group[]>(result.Result) ?? Array.Empty<Group>();
                return response;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
        
        public void CreateGroup()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
        }
        
        public void GetGroup()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
        }
        
        public void UpdateGroup()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
        }
        
        public void DeleteGroup()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
        }
        
        public void ListMembers()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
        }
        
        public void AddMember()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
        }
        
        public void RemoveMember()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
        }
        
        public void ListMembersOfGroup()
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/";
        }
    }
}