using System.Net.Http.Json;
using System.Text;
using Decisions.Exchange365.API;
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
        private const string GROUPS_URL = $"{Exchange365Constants.GRAPH_URL}/groups";
        
        public GroupList ListGroups(bool filterUnified)
        {
            string url = (filterUnified) ? $"{GROUPS_URL}$filter=groupTypes/any(c:c+eq+'Unified')" : GROUPS_URL;
            string result = GraphRest.Get(url);
            
            return JsonConvert.DeserializeObject<GroupList>(result) ?? new GroupList();
        }
        
        public Group CreateGroup(string? description, string displayName, string[]? groupTypes,
            bool mailEnabled, string mailNickname, bool securityEnabled, string[]? ownerIds, string[]? memberIds)
        {
            string[]? owners = (ownerIds != null) ? GetUserUrls(ownerIds) : Array.Empty<string>();
            string[]? members = (memberIds != null) ? GetUserUrls(memberIds) : Array.Empty<string>();
            
            MicrosoftGroup group = new MicrosoftGroup
            {
                Description = description,
                DisplayName = displayName,
                GroupTypes = groupTypes,
                MailEnabled = mailEnabled,
                MailNickname = mailNickname,
                SecurityEnabled = securityEnabled,
                Owners = owners,
                Members = members
            };
            
            HttpContent content = new StringContent(JsonConvert.SerializeObject(group), Encoding.UTF8, "application/json");
            
            string result = GraphRest.Post(GROUPS_URL, content);
            
            return JsonConvert.DeserializeObject<Group>(result) ?? throw new BusinessRuleException("Could not deserialize result");
        }
        
        public Group? GetGroup(string groupId)
        {
            string url = $"{GROUPS_URL}/{groupId}";
            string result = GraphRest.Get(url);
            
            return JsonConvert.DeserializeObject<Group>(result);
        }
        
        public string UpdateGroup(string groupId, string? description, string? displayName, string[]? groupTypes,
            bool? mailEnabled, bool? securityEnabled, string? visibility, bool? allowExternalSenders,
            AssignedLabel[]? assignedLabels, bool? autoSubscribeNewMembers, string? preferredDataLocation)
        {
            string url = $"{GROUPS_URL}/{groupId}";

            UpdateMicrosoftGroup group = new UpdateMicrosoftGroup
            {
                Description = description,
                DisplayName = displayName,
                GroupTypes = groupTypes,
                MailEnabled = mailEnabled,
                SecurityEnabled = securityEnabled,
                Visibility = visibility,
                AllowExternalSenders = allowExternalSenders,
                AssignedLabels = assignedLabels,
                AutoSubscribeNewMembers = autoSubscribeNewMembers,
                PreferredDataLocation = preferredDataLocation
            };
            
            HttpContent content = new StringContent(JsonConvert.SerializeObject(group, new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore
            }), Encoding.UTF8, "application/json");
            
            return GraphRest.HttpResponsePatch(url, content).StatusCode.ToString();
        }
        
        public string DeleteGroup(string groupId)
        {
            string url = $"{GROUPS_URL}/{groupId}";
            
            return GraphRest.Delete(url).StatusCode.ToString();
        }
        
        public MemberList ListMembers(string groupId)
        {
            string url = $"{GROUPS_URL}/{groupId}/members";
            string result = GraphRest.Get(url);
            
            return JsonConvert.DeserializeObject<MemberList>(result) ?? new MemberList();
        }
        
        public string AddMembers(string groupId, string[] directoryObjectIds)
        {
            string url = $"{GROUPS_URL}/{groupId}";

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

            return GraphRest.HttpResponsePatch(url, content).StatusCode.ToString();
        }
        
        public string RemoveMember(string groupId, string directoryObjectId)
        {
            string url = $"{GROUPS_URL}/{groupId}/members/{directoryObjectId}/$ref";
            
            return GraphRest.Delete(url).StatusCode.ToString();
        }
        
        public DirectoryObject ListMemberOf(string userIdentifier)
        {
            string url = $"{Exchange365Constants.GRAPH_URL}/users/{userIdentifier}/memberOf";
            string result = GraphRest.Get(url);
            
            return JsonConvert.DeserializeObject<DirectoryObject>(result);
        }

        private string[]? GetUserUrls(string[] users)
        {
            List<string> userUrls = new List<string>();
            foreach (string user in users)
            {
                userUrls.Add($"{Exchange365Constants.GRAPH_URL}/users/{user}");
            }

            return userUrls.ToArray();
        }
    }
}