using System;
using DecisionsFramework;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.People
{
    [Writable]
    public class Microsoft365ContactList
    {
        [WritableValue]
        [JsonProperty("@odata.context")]
        public string? OdataContext { get; set; }

        [WritableValue]
        [JsonProperty("value")]
        public Microsoft365Contact[]? Value { get; set; }
        
        public static Microsoft365ContactList? JsonDeserialize(string content)
        {
            try
            {
                return JsonConvert.DeserializeObject<Microsoft365ContactList>(content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("Could not deserialize result.", ex);
            }
        }
    }
}