using System.Net.Http.Json;
using Decisions.Exchange365.API;
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
        
        public GroupList ListGroups(bool filterUnified)
        {
            string url = (filterUnified) ? $"{Url}$filter=groupTypes/any(c:c+eq+'Unified')" : Url;
            string result = GraphRest.Get(url);
            
            return JsonConvert.DeserializeObject<GroupList>(result) ?? new GroupList();
        }
        
        public string CreateGroup(MicrosoftGroup group)
        {
            JsonContent content = JsonContent.Create(group);
            
            return GraphRest.HttpResponsePost(Url, content).StatusCode.ToString();
        }
        
        public Group? GetGroup(string groupId)
        {
            string url = $"{Url}/{groupId}";
            string result = GraphRest.Get(url);
            
            return JsonConvert.DeserializeObject<Group>(result);
        }
        
        /* TODO: test */
        public string UpdateGroup(string groupId, MicrosoftGroup group)
        {
            string url = $"{Url}/{groupId}";
            
            string content = JsonConvert.SerializeObject(group, Formatting.Indented, new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore
            });

            return GraphRest.HttpResponsePost(url, content).StatusCode.ToString();
        }
        
        public string DeleteGroup(string groupId)
        {
            string url = $"{Url}/{groupId}";
            
            return GraphRest.Delete(url).StatusCode.ToString();
        }
        
        public MemberList ListMembers(string groupId)
        {
            string url = $"{Url}/{groupId}/members";
            string result = GraphRest.Get(url);
            
            return JsonConvert.DeserializeObject<MemberList>(result) ?? new MemberList();
        }
        
        public string AddMember(string groupId, string directoryObjectId)
        {
            string url = $"{Url}/{groupId}/members/$ref";
            
            ReferenceCreate reference = new ReferenceCreate
            {
                OdataId = $"{Exchange365Constants.GRAPH_URL}/directoryObjects/{directoryObjectId}"
            };

            JsonContent content = JsonContent.Create(reference);

            return GraphRest.HttpResponsePost(url, content).StatusCode.ToString();
        }
        
        public string RemoveMember(string groupId, string userId)
        {
            string url = $"{Url}/{groupId}/members{userId}";
            
            return GraphRest.Delete(url).StatusCode.ToString();
        }
        
        public DirectoryObject[] ListMemberOf(string userIdentifier)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/memberOf";
            string result = GraphRest.Get(url);
            
            return JsonConvert.DeserializeObject<DirectoryObject[]>(result) ?? Array.Empty<DirectoryObject>();
        }
    }
}