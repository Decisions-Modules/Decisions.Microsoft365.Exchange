using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Exchange365.API
{
    [Writable]
    public class MicrosoftMembers
    {
        [WritableValue]
        [JsonProperty("members@odata.bind")]
        public string[] Members;
    }
}