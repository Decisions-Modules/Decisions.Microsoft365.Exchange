using System;
using DecisionsFramework;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.Group
{
    [Writable]
    public class Microsoft365GroupRequest
    {
        [WritableValue]
        [JsonProperty("description")]
        public string? Description { get; set; }

        [WritableValue]
        [JsonProperty("displayName")]
        public string? DisplayName { get; set; }

        [WritableValue]
        [JsonProperty("groupTypes")]
        public string[]? GroupTypes { get; set; }

        [WritableValue]
        [JsonProperty("mailEnabled")]
        public bool? MailEnabled { get; set; }

        [WritableValue]
        [JsonProperty("mailNickname")]
        public string? MailNickname { get; set; }

        [WritableValue]
        [JsonProperty("securityEnabled")]
        public bool? SecurityEnabled { get; set; }
        
        [WritableValue]
        [JsonProperty("owners@odata.bind")]
        public string[]? Owners { get; set; }
        
        [WritableValue]
        [JsonProperty("members@odata.bind")]
        public string[]? Members { get; set; }
        
        public string JsonSerialize()
        {
            try
            {
                string request = JsonConvert.SerializeObject(this);
                return request;
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("There was a problem serializing request.", ex);
            }
        }
    }
}
