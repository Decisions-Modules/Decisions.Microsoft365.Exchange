using Decisions.Microsoft365.Exchange.API.Email;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.People
{
    [Writable]
    public class Microsoft365ContactRequest
    {
        [WritableValue]
        [JsonProperty("givenName")]
        public string? GivenName { get; set; }

        [WritableValue]
        [JsonProperty("surname")]
        public string? Surname { get; set; }

        [WritableValue]
        [JsonProperty("emailAddresses")]
        public Microsoft365EmailAddress[]? EmailAddresses { get; set; }

        [WritableValue]
        [JsonProperty("businessPhones")]
        public string[]? BusinessPhones { get; set; }
    }
}
