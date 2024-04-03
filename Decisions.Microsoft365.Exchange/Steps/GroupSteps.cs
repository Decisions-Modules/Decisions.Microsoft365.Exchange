using System.Net.Http.Json;
using System.Text;
using Decisions.Microsoft365.Exchange.API;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.ServiceLayer;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Microsoft365/Exchange/Groups")]
    public class GroupSteps
    {
        private const string GROUPS_URL = $"/groups";
        
        public GroupList ListGroups(bool filterUnified)
        {
            string urlExtension = (filterUnified) ? $"{GROUPS_URL}$filter=groupTypes/any(c:c+eq+'Unified')" : GROUPS_URL;
            string result = GraphRest.Get(urlExtension);
            
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
            string urlExtension = $"{GROUPS_URL}/{groupId}";
            string result = GraphRest.Get(urlExtension);
            
            return JsonConvert.DeserializeObject<Group>(result);
        }
        
        public string UpdateGroup(string groupId, string? description, string? displayName, string[]? groupTypes,
            bool? mailEnabled, bool? securityEnabled, string? visibility, bool? allowExternalSenders,
            AssignedLabel[]? assignedLabels, bool? autoSubscribeNewMembers, string? preferredDataLocation)
        {
            string urlExtension = $"{GROUPS_URL}/{groupId}";

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
            
            return GraphRest.HttpResponsePatch(urlExtension, content).StatusCode.ToString();
        }
        
        public string DeleteGroup(string groupId)
        {
            string urlExtension = $"{GROUPS_URL}/{groupId}";
            
            return GraphRest.Delete(urlExtension).StatusCode.ToString();
        }
        
        public MemberList ListMembers(string groupId)
        {
            string urlExtension = $"{GROUPS_URL}/{groupId}/members";
            string result = GraphRest.Get(urlExtension);
            
            return JsonConvert.DeserializeObject<MemberList>(result) ?? new MemberList();
        }
        
        public string AddMembers(string groupId, string[] directoryObjectIds)
        {
            string urlExtension = $"{GROUPS_URL}/{groupId}";

            List<string> memberList = new();
            foreach (string directoryObjectId in directoryObjectIds)
            {
                memberList.Add($"{ModuleSettingsAccessor<ExchangeSettings>.GetSettings().GraphUrl}/directoryObjects/{directoryObjectId}");
            }

            MicrosoftMembers members = new MicrosoftMembers
            {
                Members = memberList.ToArray()
            };

            JsonContent content = JsonContent.Create(members);

            return GraphRest.HttpResponsePatch(urlExtension, content).StatusCode.ToString();
        }
        
        public string RemoveMember(string groupId, string directoryObjectId)
        {
            string urlExtension = $"{GROUPS_URL}/{groupId}/members/{directoryObjectId}/$ref";
            
            return GraphRest.Delete(urlExtension).StatusCode.ToString();
        }
        
        public DirectoryObject ListMemberOf(string userIdentifier)
        {
            string urlExtension = $"/users/{userIdentifier}/memberOf";
            string result = GraphRest.Get(urlExtension);
            
            return JsonConvert.DeserializeObject<DirectoryObject>(result);
        }

        private string[]? GetUserUrls(string[] users)
        {
            List<string> userUrls = new List<string>();
            foreach (string user in users)
            {
                userUrls.Add($"{ModuleSettingsAccessor<ExchangeSettings>.GetSettings().GraphUrl}/users/{user}");
            }

            return userUrls.ToArray();
        }
    }
}