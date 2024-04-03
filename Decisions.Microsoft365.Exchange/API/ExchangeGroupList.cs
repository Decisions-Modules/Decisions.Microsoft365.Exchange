using DecisionsFramework;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API
{
    [Writable]
    public class ExchangeGroupList
    {
        [WritableValue]
        [JsonProperty("@odata.context")]
        public Uri OdataContext { get; set; }

        [WritableValue]
        [JsonProperty("value")]
        public Group[] Value { get; set; }
        
        public static ExchangeGroupList? JsonDeserialize(string content)
        {
            try
            {
                return JsonConvert.DeserializeObject<ExchangeGroupList>(content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("Could not deserialize result.", ex);
            }
        }
    }
}