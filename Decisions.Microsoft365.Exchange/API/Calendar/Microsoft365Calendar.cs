using Decisions.Microsoft365.Exchange.API.Email;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using Newtonsoft.Json;

namespace Decisions.Microsoft365.Exchange.API.Calendar
{
    [Writable]
    public class Microsoft365Calendar
    {
        [WritableValue]
        [JsonProperty("allowedOnlineMeetingProviders")]
        public string[]? AllowedOnlineMeetingProviders { get; set; }

        [WritableValue]
        [JsonProperty("canEdit")]
        public bool? CanEdit { get; set; }

        [WritableValue]
        [JsonProperty("canShare")]
        public bool? CanShare { get; set; }

        [WritableValue]
        [JsonProperty("canViewPrivateItems")]
        public bool? CanViewPrivateItems { get; set; }

        [WritableValue]
        [JsonProperty("changeKey")]
        public string? ChangeKey { get; set; }

        [WritableValue]
        [JsonProperty("color")]
        public string? Color { get; set; }

        [WritableValue]
        [JsonProperty("defaultOnlineMeetingProvider")]
        public string? DefaultOnlineMeetingProvider { get; set; }

        [WritableValue]
        [JsonProperty("hexColor")]
        public string? HexColor { get; set; }

        [WritableValue]
        [JsonProperty("id")]
        public string? Id { get; set; }

        [WritableValue]
        [JsonProperty("isDefaultCalendar")]
        public bool? IsDefaultCalendar { get; set; }

        [WritableValue]
        [JsonProperty("isRemovable")]
        public bool? IsRemovable { get; set; }

        [WritableValue]
        [JsonProperty("isTallyingResponses")]
        public bool? IsTallyingResponses { get; set; }

        [WritableValue]
        [JsonProperty("name")]
        public string? Name { get; set; }

        [WritableValue]
        [JsonProperty("owner")]
        public Microsoft365EmailAddress? Owner { get; set; }
    }
}