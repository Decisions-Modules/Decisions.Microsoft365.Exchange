using DecisionsFramework;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API
{
    [Writable]
    public class ExchangeContactList
    {
        [WritableValue]
        [JsonProperty("@odata.context")]
        public Uri OdataContext { get; set; }

        [WritableValue]
        [JsonProperty("value")]
        public Contact[] Value { get; set; }
        
        public static ExchangeContactList? JsonDeserialize(string content)
        {
            try
            {
                return JsonConvert.DeserializeObject<ExchangeContactList>(content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("Could not deserialize result.", ex);
            }
        }
    }
}