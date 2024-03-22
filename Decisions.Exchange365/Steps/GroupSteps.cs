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
        
        // TODO: update request
        public string CreateGroup(string description, string displayName, string[] groupTypes, bool isAssignableToRole,
            bool mailEnabled, string mailNickname, bool securityEnabled, string[] ownerIds, string[] memberIds)
        {
            List<string> owners = new List<string>();
            foreach (string ownerId in ownerIds)
            {
                owners.Add($"https://graph.microsoft.com/v1.0/users/{ownerId}");
            }

            List<string> members = new List<string>();
            foreach (string memberId in memberIds)
            {
                members.Add($"{Exchange365Constants.GRAPH_URL}/users/{memberId}");
            }
            
            MicrosoftGroup group = new MicrosoftGroup
            {
                Description = description,
                DisplayName = displayName,
                GroupTypes = groupTypes,
                IsAssignableToRole = isAssignableToRole,
                MailEnabled = mailEnabled,
                MailNickname = mailNickname,
                SecurityEnabled = securityEnabled,
                Owners = owners.ToArray(),
                Members = members.ToArray()
            };
            
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

            return GraphRest.Patch(url, content).StatusCode.ToString();
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
        
        public string AddMembers(string groupId, string[] directoryObjectIds)
        {
            string url = $"{Url}/{groupId}";

            List<string> memberList = new();
            foreach (string directoryObjectId in directoryObjectIds)
            {
                memberList.Add($"{Exchange365Constants.GRAPH_URL}/directoryObjects/{directoryObjectId}");
            }

            MicrosoftMembers members = new MicrosoftMembers
            {
                Members = memberList.ToArray()
            };

            JsonContent content = JsonContent.Create(members);

            return GraphRest.Patch(url, content).StatusCode.ToString();
        }
        
        public string RemoveMember(string groupId, string directoryObjectId)
        {
            string url = $"{Url}/{groupId}/members/{directoryObjectId}/$ref";
            
            return GraphRest.Delete(url).StatusCode.ToString();
        }
        
        public DirectoryObject ListMemberOf(string userIdentifier)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/memberOf";
            string result = GraphRest.Get(url);
            
            return JsonConvert.DeserializeObject<DirectoryObject>(result);
        }
    }
}