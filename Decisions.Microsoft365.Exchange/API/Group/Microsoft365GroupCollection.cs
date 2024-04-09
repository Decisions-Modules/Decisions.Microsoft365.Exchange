using System;
using DecisionsFramework;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.Group
{
    public class Microsoft365GroupCollection
    {
        [JsonProperty("@odata.context")]
        public string? OdataContext { get; set; }

        [JsonProperty("value")]
        public Microsoft365Group[]? Value { get; set; }
        
        public static Microsoft365GroupCollection? JsonDeserialize(string content)
        {
            try
            {
                return JsonConvert.DeserializeObject<Microsoft365GroupCollection>(content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("Could not deserialize result.", ex);
            }
        }
    }
}