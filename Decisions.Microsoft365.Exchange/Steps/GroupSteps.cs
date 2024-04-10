using System.Net.Http.Json;
using System.Text;
using Decisions.Microsoft365.Common;
using Decisions.Microsoft365.Common.API.Group;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.ServiceLayer;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Microsoft365/Exchange/Groups")]
    public class GroupSteps
    {
        private static JsonSerializerSettings IgnoreNullValues = new()
        {
            NullValueHandling = NullValueHandling.Ignore
        };
        
        public Microsoft365GroupList? ListGroups(bool filterUnified)
        {
            string urlExtension = (filterUnified) ? $"{Microsoft365UrlHelper.GetGroupUrl(null)}$filter=groupTypes/any(c:c+eq+'Unified')" : Microsoft365UrlHelper.GetGroupUrl(null);
            string result = GraphRest.Get(urlExtension);
            
            return JsonHelper<Microsoft365GroupList?>.JsonDeserialize(result);
        }
        
        public Microsoft365Group? CreateGroup(string? description, string displayName, string[]? groupTypes,
            bool mailEnabled, string mailNickname, bool securityEnabled, string[]? ownerIds, string[]? memberIds)
        {
            string[]? owners = (ownerIds != null) ? GetUserUrlStrings(ownerIds) : Array.Empty<string>();
            string[]? members = (memberIds != null) ? GetUserUrlStrings(memberIds) : Array.Empty<string>();
            
            Microsoft365GroupRequest groupRequest = new Microsoft365GroupRequest
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
            
            HttpContent content = new StringContent(groupRequest.JsonSerialize(), Encoding.UTF8, "application/json");
            string result = GraphRest.Post(Microsoft365UrlHelper.GetGroupUrl(null), content);

            return JsonHelper<Microsoft365Group?>.JsonDeserialize(result);
        }
        
        public Microsoft365Group? GetGroup(string groupId)
        {
            string urlExtension = Microsoft365UrlHelper.GetGroupUrl(groupId);
            string result = GraphRest.Get(urlExtension);

            return JsonHelper<Microsoft365Group?>.JsonDeserialize(result);
        }
        
        public string UpdateGroup(string groupId, Microsoft365UpdateGroup group)
        {
            string urlExtension = Microsoft365UrlHelper.GetGroupUrl(groupId);
            
            HttpContent content = new StringContent(JsonConvert.SerializeObject(group, IgnoreNullValues), Encoding.UTF8,
                "application/json");
            HttpResponseMessage response = GraphRest.HttpResponsePatch(urlExtension, content);
            
            return response.StatusCode.ToString();
        }
        
        public string DeleteGroup(string groupId)
        {
            string urlExtension = Microsoft365UrlHelper.GetGroupUrl(groupId);
            HttpResponseMessage response = GraphRest.Delete(urlExtension);
            
            return response.StatusCode.ToString();
        }
        
        public Microsoft365MemberList? ListMembers(string groupId)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetGroupUrl(groupId)}/members";
            string result = GraphRest.Get(urlExtension);
            
            return JsonHelper<Microsoft365MemberList?>.JsonDeserialize(result);
        }
        
        public string AddMembers(string groupId, string[] directoryObjectIds)
        {
            string urlExtension = Microsoft365UrlHelper.GetGroupUrl(groupId);

            List<string> memberList = new();
            foreach (string directoryObjectId in directoryObjectIds)
            {
                memberList.Add($"{ModuleSettingsAccessor<ExchangeSettings>.GetSettings().GraphUrl}/directoryObjects/{directoryObjectId}");
            }

            Microsoft365MembersRequest membersRequest = new Microsoft365MembersRequest
            {
                Members = memberList.ToArray()
            };

            JsonContent content = JsonContent.Create(membersRequest);
            HttpResponseMessage response = GraphRest.HttpResponsePatch(urlExtension, content);

            return response.StatusCode.ToString();
        }
        
        public string RemoveMember(string groupId, string directoryObjectId)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetGroupUrl(groupId)}/members/{directoryObjectId}/$ref";
            HttpResponseMessage response = GraphRest.Delete(urlExtension);
            
            return response.StatusCode.ToString();
        }
        
        public Microsoft365GroupCollection? ListMemberOf(string userIdentifier)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetUserUrl(userIdentifier)}/memberOf";
            string result = GraphRest.Get(urlExtension);
            
            return JsonHelper<Microsoft365GroupCollection?>.JsonDeserialize(result);
        }

        private string[]? GetUserUrlStrings(string[] users)
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