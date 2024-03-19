using Newtonsoft.Json;

namespace Decisions.Exchange365.API
{
    public partial class MicrosoftGroup
    {
        [JsonProperty("description")]
        public string Description { get; set; }

        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [JsonProperty("groupTypes")]
        public string[] GroupTypes { get; set; }

        [JsonProperty("mailEnabled")]
        public bool MailEnabled { get; set; }

        [JsonProperty("mailNickname")]
        public string MailNickname { get; set; }

        [JsonProperty("securityEnabled")]
        public bool SecurityEnabled { get; set; }
    }
}
