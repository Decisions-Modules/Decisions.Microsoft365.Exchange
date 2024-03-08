using System.Net;
using System.Net.Http.Json;
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
        private const string Url = $"{Exchange365Constants.GRAPH_URL}/groups";
        
        public Group[] ListGroups(bool filterUnified)
        {
            string url = (filterUnified) ? $"{Url}$filter=groupTypes/any(c:c+eq+'Unified')" : Url;
            
            try
            {
                string result = GraphRest.Get(url);
                Group[] response = JsonConvert.DeserializeObject<Group[]>(result) ?? Array.Empty<Group>();
                return response;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
        
        public void CreateGroup()
        {
            
        }
        
        public Group? GetGroup(string groupName)
        {
            try
            {
                string result = GraphRest.Get($"{Url}/{groupName}");
                Group? response = JsonConvert.DeserializeObject<Group>(result);
                return response;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("The request was unsuccessful.", ex);
            }
        }
        
        public HttpStatusCode UpdateGroup(string groupId, string groupContext)
        {
            JsonContent content = JsonContent.Create(groupContext);

            return GraphRest.HttpResponsePost($"{Url}/{groupId}", content).StatusCode;
        }
        
        public HttpStatusCode DeleteGroup(string groupId)
        {
            return GraphRest.Delete($"{Url}/{groupId}").StatusCode;
        }
        
        public void ListMembers()
        {
            
        }
        
        public void AddMember()
        {
            
        }
        
        public void RemoveMember()
        {
            
        }
        
        public void ListMembersOfGroup()
        {
            
        }
    }
}