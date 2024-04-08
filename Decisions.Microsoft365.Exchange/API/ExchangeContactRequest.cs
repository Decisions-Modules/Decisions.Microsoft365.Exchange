using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API
{
    [Writable]
    public class ExchangeContactRequest
    {
        [WritableValue]
        [JsonProperty("givenName")]
        public string? GivenName { get; set; }

        [WritableValue]
        [JsonProperty("surname")]
        public string? Surname { get; set; }

        [WritableValue]
        [JsonProperty("emailAddresses")]
        public MicrosoftEmailAddress[]? EmailAddresses { get; set; }

        [WritableValue]
        [JsonProperty("businessPhones")]
        public string[]? BusinessPhones { get; set; }
    }
}
