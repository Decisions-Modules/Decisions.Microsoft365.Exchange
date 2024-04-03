using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API
{
    [Writable]
    public class MicrosoftContact
    {
        [WritableValue]
        [JsonProperty("givenName")]
        public string GivenName { get; set; }

        [WritableValue]
        [JsonProperty("surname")]
        public string? Surname { get; set; }

        [WritableValue]
        [JsonProperty("emailAddresses")]
        public EmailAddressName[] EmailAddresses { get; set; }

        [WritableValue]
        [JsonProperty("businessPhones")]
        public string[]? BusinessPhones { get; set; }
    }

    [Writable]
    public class EmailAddressName
    {
        [WritableValue]
        [JsonProperty("address")]
        public string Address { get; set; }

        [WritableValue]
        [JsonProperty("name")]
        public string? Name { get; set; }
    }
}
