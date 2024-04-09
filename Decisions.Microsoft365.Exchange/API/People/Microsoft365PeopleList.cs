using System;
using DecisionsFramework;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.People
{
    [Writable]
    public class Microsoft365PeopleList
    {
        [WritableValue]
        [JsonProperty("value")]
        public Microsoft365Person[] Value { get; set; }
        
        public static Microsoft365PeopleList? JsonDeserialize(string content)
        {
            try
            {
                return JsonConvert.DeserializeObject<Microsoft365PeopleList>(content);
            }
            catch (Exception ex)
            {
                throw new BusinessRuleException("Could not deserialize result.", ex);
            }
        }
    }
}