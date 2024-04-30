using System.Net.Http.Json;
using System.Text;
using Decisions.Microsoft365.Common;
using Decisions.Microsoft365.Common.API.Group;
using DecisionsFramework;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.ServiceLayer;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.Steps
{
    [AutoRegisterMethodsOnClass(true, "Integration/Microsoft365/Exchange/Groups")]
    public class GroupSteps
    {
        private static readonly JsonSerializerSettings IgnoreNullValues = new()
        {
            NullValueHandling = NullValueHandling.Ignore
        };
        
        public Microsoft365GroupList? ListGroups(bool filterUnified,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride = null)
        {
            string urlExtension = (filterUnified) ? $"{Microsoft365UrlHelper.GetGroupUrl(null)}$filter=groupTypes/any(c:c+eq+'Unified')" : Microsoft365UrlHelper.GetGroupUrl(null);
            string result = GraphRest.Get(settingsOverride, urlExtension);
            
            return JsonHelper<Microsoft365GroupList?>.JsonDeserialize(result);
        }
        
        public Microsoft365Group? CreateGroup(string? description, string displayName, string[]? groupTypes,
            bool? mailEnabled, string mailNickname, bool? securityEnabled, string[]? ownerIds, string[]? memberIds,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            if (string.IsNullOrEmpty(mailNickname))
            {
                throw new BusinessRuleException("Mail Nickname cannot be null or empty. Please set a unique Mail Nickname.");
            }
            
            mailEnabled ??= false;
            securityEnabled ??= false;
            groupTypes ??= Array.Empty<string>();
            
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
            string result = GraphRest.Post(settingsOverride, Microsoft365UrlHelper.GetGroupUrl(null), content);

            return JsonHelper<Microsoft365Group?>.JsonDeserialize(result);
        }
        
        public Microsoft365Group? GetGroup(string groupId,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = Microsoft365UrlHelper.GetGroupUrl(groupId);
            string result = GraphRest.Get(settingsOverride, urlExtension);

            return JsonHelper<Microsoft365Group?>.JsonDeserialize(result);
        }
        
        public string UpdateGroup(string groupId, Microsoft365UpdateGroup group,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = Microsoft365UrlHelper.GetGroupUrl(groupId);

            // Some objects can only be patched individually, so separate requests will be made.
            LesserGroupUpdates(settingsOverride, urlExtension, group);
            
            HttpContent content = new StringContent(JsonConvert.SerializeObject(group, IgnoreNullValues),
                Encoding.UTF8, "application/json");
            HttpResponseMessage response = GraphRest.HttpResponsePatch(settingsOverride, urlExtension, content);
            
            return response.StatusCode.ToString();
        }
        
        public string DeleteGroup(string groupId,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = Microsoft365UrlHelper.GetGroupUrl(groupId);
            HttpResponseMessage response = GraphRest.Delete(settingsOverride, urlExtension);
            
            return response.StatusCode.ToString();
        }
        
        public Microsoft365MemberList? ListMembers(string groupId,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetGroupUrl(groupId)}/members";
            string result = GraphRest.Get(settingsOverride, urlExtension);
            
            return JsonHelper<Microsoft365MemberList?>.JsonDeserialize(result);
        }
        
        public string AddMembers(string groupId, string[] directoryObjectIds,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
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
            HttpResponseMessage response = GraphRest.HttpResponsePatch(settingsOverride, urlExtension, content);

            return response.StatusCode.ToString();
        }
        
        public string RemoveMember(string groupId, string directoryObjectId,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetGroupUrl(groupId)}/members/{directoryObjectId}/$ref";
            HttpResponseMessage response = GraphRest.Delete(settingsOverride, urlExtension);
            
            return response.StatusCode.ToString();
        }
        
        public Microsoft365GroupCollection? ListMemberOf(string userIdentifier,
            [PropertyClassification(0, "Settings Override", "Settings")] ExchangeSettings? settingsOverride)
        {
            string urlExtension = $"{Microsoft365UrlHelper.GetUserUrl(userIdentifier)}/memberOf";
            string result = GraphRest.Get(settingsOverride, urlExtension);
            
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

        private void LesserGroupUpdates(ExchangeSettings? settingsOverride, string urlExtension, Microsoft365UpdateGroup group)
        {
            List<Microsoft365UpdateGroup> lesserGroups = new List<Microsoft365UpdateGroup>();
            if (group.AllowExternalSenders != null)
            {
                lesserGroups.Add(new() { AllowExternalSenders = group.AllowExternalSenders });
            }
            if (group.AutoSubscribeNewMembers != null)
            {
                lesserGroups.Add(new() { AutoSubscribeNewMembers = group.AutoSubscribeNewMembers});
            }

            foreach (Microsoft365UpdateGroup lesserGroup in lesserGroups)
            {
                HttpContent content = new StringContent(JsonConvert.SerializeObject(lesserGroup, IgnoreNullValues),
                    Encoding.UTF8, "application/json");
                GraphRest.HttpResponsePatch(settingsOverride, urlExtension, content);
            }
        }
    }
}