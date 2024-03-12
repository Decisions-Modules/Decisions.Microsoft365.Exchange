using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Exchange365.API
{
    [Writable]
    public class EmailList
    {
        [WritableValue]
        [JsonProperty("@odata.context")]
        public Uri OdataContext { get; set; }

        [WritableValue]
        [JsonProperty("value")]
        public EmailValue[] Value { get; set; }
    }

    [Writable]
    public class EmailValue
    {
        [WritableValue]
        [JsonProperty("@odata.etag")]
        public string OdataEtag { get; set; }

        [WritableValue]
        [JsonProperty("id")]
        public Guid Id { get; set; }

        [WritableValue]
        [JsonProperty("subject")]
        public string Subject { get; set; }

        [WritableValue]
        [JsonProperty("sender")]
        public Sender Sender { get; set; }
    }

    [Writable]
    public class Sender
    {
        [WritableValue]
        [JsonProperty("emailAddress")]
        public EmailAddress EmailAddress { get; set; }
    }

    [Writable]
    public class EmailAddress
    {
        [WritableValue]
        [JsonProperty("name")]
        public string Name { get; set; }

        [WritableValue]
        [JsonProperty("address")]
        public string Address { get; set; }
    }
}
