using System.Net;
using System.Net.Http.Json;
using Decisions.Exchange365.Data;
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
            
            string result = GraphRest.Get(url);
            Group[] response = JsonConvert.DeserializeObject<Group[]>(result) ?? Array.Empty<Group>();
            
            return response;
        }
        
        public HttpStatusCode CreateGroup(Group group)
        {
            JsonContent content = JsonContent.Create(group);
            
            return GraphRest.HttpResponsePost(Url, content).StatusCode;
        }
        
        public Group? GetGroup(string groupName)
        {
            
            string result = GraphRest.Get($"{Url}/{groupName}");
            Group? response = JsonConvert.DeserializeObject<Group>(result);
            
            return response;
        }
        
        public HttpStatusCode UpdateGroup(string groupId, string groupContext)
        {
            JsonContent content = JsonContent.Create(groupContext);
            
            /* TODO: Utilize UpdateODataEntityStep features to dynamically build request data */

            return GraphRest.HttpResponsePost($"{Url}/{groupId}", content).StatusCode;
        }
        
        public HttpStatusCode DeleteGroup(string groupId)
        {
            return GraphRest.Delete($"{Url}/{groupId}").StatusCode;
        }
        
        public void ListMembers()
        {
            
        }
        
        public HttpStatusCode AddMember(string groupId, User user)
        {
            JsonContent content = JsonContent.Create(user);

            return GraphRest.HttpResponsePost($"{Url}/{groupId}/{user}", content).StatusCode;
        }
        
        public HttpStatusCode RemoveMember(string groupId, string userId)
        {
            return GraphRest.Delete($"{Url}/{groupId}/members{userId}").StatusCode;
        }
        
        public void ListMembersOfGroup()
        {
            
        }
    }
}