using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API
{
    [Writable]
    public class ContactList
    {
        [WritableValue]
        [JsonProperty("@odata.context")]
        public Uri OdataContext { get; set; }

        [WritableValue]
        [JsonProperty("value")]
        public Contact[] Value { get; set; }
    }
}