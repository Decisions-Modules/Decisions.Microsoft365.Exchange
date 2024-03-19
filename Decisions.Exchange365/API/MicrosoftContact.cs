using Newtonsoft.Json;

namespace Decisions.Exchange365.API
{
    public partial class MicrosoftContact
    {
        [JsonProperty("givenName")]
        public string GivenName { get; set; }

        [JsonProperty("surname")]
        public string? Surname { get; set; }

        [JsonProperty("emailAddresses")]
        public EmailAddressName[] EmailAddresses { get; set; }

        [JsonProperty("businessPhones")]
        public string[]? BusinessPhones { get; set; }
    }

    public partial class EmailAddressName
    {
        [JsonProperty("address")]
        public string Address { get; set; }

        [JsonProperty("name")]
        public string? Name { get; set; }
    }
}
