using DecisionsFramework;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API
{
    [Writable]
    public class ExchangePeopleList
    {
        [WritableValue]
        [JsonProperty("value")]
        public Person[] Value { get; set; }
        
        public static ExchangePeopleList? JsonDeserialize(string content)
        {
            try
            {
                return JsonConvert.DeserializeObject<ExchangePeopleList>(content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("Could not deserialize result.", ex);
            }
        }
    }
}