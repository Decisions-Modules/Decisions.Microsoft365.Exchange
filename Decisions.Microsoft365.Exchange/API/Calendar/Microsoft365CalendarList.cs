using System;
using DecisionsFramework;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.Calendar
{
    [Writable]
    public class Microsoft365CalendarList
    {
        [WritableValue]
        [JsonProperty("@odata.context")]
        public string? OdataContext { get; set; }

        [WritableValue]
        [JsonProperty("value")]
        public Microsoft365Calendar[]? Value { get; set; }
        
        public static Microsoft365CalendarList? JsonDeserialize(string content)
        {
            try
            {
                return JsonConvert.DeserializeObject<Microsoft365CalendarList>(content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("Could not deserialize result.", ex);
            }
        }
    }
}