using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace Decisions.Exchange365.API
{
    [Writable]
    public class PeopleList
    {
        [WritableValue]
        [JsonProperty("value")]
        public Person[] Value { get; set; }
    }
}