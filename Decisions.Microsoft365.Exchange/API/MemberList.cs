using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API
{
    [Writable]
    public class MemberList
    {
        [WritableValue]
        [JsonProperty("@odata.context")]
        public Uri OdataContext { get; set; }

        [WritableValue]
        [JsonProperty("value")]
        public DirectoryObject[] Value { get; set; }
    }
}