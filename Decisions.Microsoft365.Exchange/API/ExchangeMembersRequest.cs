using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API
{
    [Writable]
    public class ExchangeMembersRequest
    {
        [WritableValue]
        [JsonProperty("members@odata.bind")]
        public string[] Members;
    }
}