using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Exchange365.Data
{
    [Writable]
    public partial class EmailDataList
    {
        [WritableValue]
        [JsonProperty("@odata.context")]
        public Uri OdataContext { get; set; }

        [WritableValue]
        [JsonProperty("value")]
        public Value[] Value { get; set; }
    }

    [Writable]
    public class Value
    {
        [WritableValue]
        [JsonProperty("@odata.etag")]
        public string OdataEtag { get; set; }

        [WritableValue]
        [JsonProperty("id")]
        public string Id { get; set; }

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
