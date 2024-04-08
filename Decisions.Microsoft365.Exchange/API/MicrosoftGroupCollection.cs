using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API
{
    public class MicrosoftGroupCollection
    {
        [JsonProperty("@odata.context")]
        public string? OdataContext { get; set; }

        [JsonProperty("value")]
        public MicrosoftGroup[]? Value { get; set; }
    }
}