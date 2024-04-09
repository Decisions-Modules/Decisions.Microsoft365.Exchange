using System;
using DecisionsFramework;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.Calendar
{
    [Writable]
    public class Microsoft365EventList
    {
        [WritableValue]
        [JsonProperty("@odata.context")]
        public string? OdataContext { get; set; }

        [WritableValue]
        [JsonProperty("value")]
        public Microsoft365Event[]? Value { get; set; }
        
        public static Microsoft365EventList? JsonDeserialize(string content)
        {
            try
            {
                return JsonConvert.DeserializeObject<Microsoft365EventList>(content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("Could not deserialize result.", ex);
            }
        }
    }
}