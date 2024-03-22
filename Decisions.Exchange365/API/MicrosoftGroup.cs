using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Exchange365.API
{
    [Writable]
    public class MicrosoftGroup
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
        [JsonProperty("isAssignableToRole")]
        public bool? IsAssignableToRole { get; set; }

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
    }
}
